"""API wrapper for Microsoft Planner."""
from __future__ import annotations

import logging
from typing import Any
from urllib.parse import quote

import msal
import requests

_LOGGER = logging.getLogger(__name__)

GRAPH_API_ENDPOINT = "https://graph.microsoft.com/v1.0"


class PlannerAPI:
    """Microsoft Planner API wrapper."""

    def __init__(self, client_id: str, client_secret: str, tenant_id: str) -> None:
        """Initialize the API wrapper."""
        self.client_id = client_id
        self.client_secret = client_secret
        self.tenant_id = tenant_id
        self.access_token = None

    def authenticate(self) -> None:
        """Authenticate with Microsoft Graph using client credentials flow."""
        authority = f"https://login.microsoftonline.com/{self.tenant_id}"
        scope = ["https://graph.microsoft.com/.default"]

        app = msal.ConfidentialClientApplication(
            self.client_id,
            authority=authority,
            client_credential=self.client_secret,
        )

        result = app.acquire_token_for_client(scopes=scope)

        if "access_token" in result:
            self.access_token = result["access_token"]
            _LOGGER.info("Successfully authenticated with Microsoft Graph")
            
            # Log token info for debugging (without exposing the actual token)
            if "expires_in" in result:
                _LOGGER.debug("Token expires in %s seconds", result["expires_in"])
            
            # Try to decode and log the scopes/permissions
            try:
                import base64
                import json
                # JWT tokens have 3 parts separated by dots
                token_parts = self.access_token.split('.')
                if len(token_parts) >= 2:
                    # Decode the payload (second part)
                    # Add padding if needed
                    payload = token_parts[1]
                    payload += '=' * (4 - len(payload) % 4)
                    decoded = base64.b64decode(payload)
                    token_data = json.loads(decoded)
                    
                    if "roles" in token_data:
                        _LOGGER.info("Token has application roles: %s", token_data["roles"])
                    if "scp" in token_data:
                        _LOGGER.info("Token has delegated scopes: %s", token_data["scp"])
                    
                    _LOGGER.debug("Token issued for app: %s", token_data.get("appid", "unknown"))
            except Exception as decode_err:
                _LOGGER.debug("Could not decode token info: %s", decode_err)
        else:
            error = result.get("error")
            error_description = result.get("error_description")
            _LOGGER.error(
                "Failed to acquire token: %s - %s", error, error_description
            )
            raise Exception(f"Authentication failed: {error} - {error_description}")

    def _get_headers(self) -> dict[str, str]:
        """Get headers for API requests."""
        if not self.access_token:
            self.authenticate()
        return {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json",
        }

    def _make_request(self, endpoint: str, retry_auth: bool = True) -> Any:
        """Make a request to the Microsoft Graph API."""
        url = f"{GRAPH_API_ENDPOINT}/{endpoint}"
        _LOGGER.debug("Making request to: %s", url)
        
        try:
            response = requests.get(url, headers=self._get_headers(), timeout=30)
            _LOGGER.debug("Response status: %s", response.status_code)
            
            # If we get 401 and haven't retried yet, refresh token and retry
            if response.status_code == 401 and retry_auth:
                _LOGGER.debug("Token expired, refreshing...")
                self.access_token = None  # Force re-authentication
                self.authenticate()
                # Retry the request once with new token
                return self._make_request(endpoint, retry_auth=False)
            
            response.raise_for_status()
            return response.json()
        except requests.exceptions.HTTPError as err:
            _LOGGER.error("HTTP Error for %s: %s", url, err)
            if err.response is not None:
                _LOGGER.error("Response body: %s", err.response.text)
            raise

    def get_user_display_name(self, user_id: str) -> str:
        """Get display name for a user ID."""
        try:
            user_response = self._make_request(f"users/{user_id}")
            return user_response.get("displayName", user_id)
        except Exception as err:
            _LOGGER.warning("Could not resolve user ID %s: %s", user_id, err)
            return user_id

    @staticmethod
    def _escape_odata_string(value: str) -> str:
        """Escape quotes for OData filters."""
        return value.replace("'", "''")

    def get_user_id_by_name(self, display_name: str) -> str | None:
        """Resolve a user ID from UPN, mail nickname, or display name."""
        identifier = (display_name or "").strip()
        if not identifier:
            _LOGGER.warning("Empty user identifier provided for lookup")
            return None

        # Try direct lookup first – Graph accepts object ID or UPN on /users/{id}
        try:
            user_response = self._make_request(f"users/{quote(identifier)}")
            if user_response:
                return user_response.get("id")
        except requests.exceptions.HTTPError as err:
            if err.response is not None and err.response.status_code == 404:
                _LOGGER.debug("Direct lookup for '%s' returned 404", identifier)
            else:
                _LOGGER.debug("Direct lookup for '%s' failed: %s", identifier, err)
        except Exception as err:
            _LOGGER.debug("Direct lookup for '%s' errored: %s", identifier, err)

        escaped_value = self._escape_odata_string(identifier)
        filter_expressions = [
            f"userPrincipalName eq '{escaped_value}'",
            f"mail eq '{escaped_value}'",
            f"mailNickname eq '{escaped_value}'",
            f"displayName eq '{escaped_value}'",
            f"startswith(mailNickname,'{escaped_value}')",
        ]

        for filter_query in filter_expressions:
            try:
                users_response = self._make_request(f"users?$filter={filter_query}")
                users = users_response.get("value", [])
                if users:
                    return users[0].get("id")
            except Exception as err:
                _LOGGER.debug("Filter lookup '%s' failed: %s", filter_query, err)

        _LOGGER.warning("User '%s' not found", identifier)
        return None

    def get_task_assignments(self, task_id: str) -> list[str]:
        """Get the list of assignees for a task."""
        try:
            # Get task details which include assignments
            details_response = self._make_request(f"planner/tasks/{task_id}/details")
            
            # The assignments are in the task itself, not details
            # We need to get the full task with assignments
            task_response = self._make_request(f"planner/tasks/{task_id}")
            assignments = task_response.get("assignments", {})
            
            assignees = []
            for user_id in assignments.keys():
                if assignments[user_id]:  # Assignment is not null/empty
                    display_name = self.get_user_display_name(user_id)
                    assignees.append(display_name)
            
            return assignees
        except Exception as err:
            _LOGGER.warning("Could not get assignments for task %s: %s", task_id, err)
            return []

    def list_all_groups(self) -> list[dict[str, Any]]:
        """List all groups accessible to the app."""
        try:
            groups_response = self._make_request("groups")
            groups = groups_response.get("value", [])
            _LOGGER.debug("Found %d groups", len(groups))
            for group in groups:
                _LOGGER.debug("Group: %s (ID: %s)", group.get("displayName"), group.get("id"))
            return groups
        except requests.exceptions.HTTPError as err:
            if err.response is not None and err.response.status_code == 401:
                _LOGGER.error(
                    "Error listing groups: %s. "
                    "This usually means the app doesn't have proper permissions. "
                    "Required permissions: Group.Read.All, Tasks.Read (or Tasks.ReadWrite). "
                    "Make sure you clicked 'Grant admin consent' in Azure Portal after adding permissions.",
                    err
                )
            else:
                _LOGGER.error("Error listing groups: %s", err)
            return []
        except Exception as err:
            _LOGGER.error("Error listing groups: %s", err)
            return []

    def list_all_plans(self) -> list[dict[str, Any]]:
        """List all plans across all groups."""
        all_plans = []
        groups = self.list_all_groups()
        
        for group in groups:
            group_id = group.get("id")
            try:
                plans_response = self._make_request(f"groups/{group_id}/planner/plans")
                plans = plans_response.get("value", [])
                for plan in plans:
                    plan["group_name"] = group.get("displayName")
                    all_plans.append(plan)
                    _LOGGER.debug(
                        "Found plan: '%s' in group '%s' (Plan ID: %s)",
                        plan.get("title"),
                        group.get("displayName"),
                        plan.get("id")
                    )
            except requests.exceptions.HTTPError as err:
                if err.response.status_code == 403:
                    _LOGGER.debug("No access to plans in group: %s", group.get("displayName"))
                else:
                    _LOGGER.debug("Error getting plans for group %s: %s", group.get("displayName"), err)
            except Exception as err:
                _LOGGER.debug("Error getting plans for group %s: %s", group.get("displayName"), err)
        
        return all_plans

    def get_plan_by_name(self, plan_name: str) -> dict[str, Any] | None:
        """Get a plan by its name."""
        try:
            _LOGGER.debug("Searching for plan: '%s'", plan_name)
            all_plans = self.list_all_plans()
            
            _LOGGER.debug("Total plans found: %d", len(all_plans))
            
            for plan in all_plans:
                plan_title = plan.get("title", "")
                if plan_title == plan_name:
                    _LOGGER.debug("Found matching plan: %s with ID: %s", plan_name, plan.get("id"))
                    return plan
            
            _LOGGER.warning("Plan '%s' not found among %d plans", plan_name, len(all_plans))
            _LOGGER.debug("Available plans: %s", [p.get("title") for p in all_plans])
            return None
            
        except requests.exceptions.HTTPError as err:
            if err.response.status_code == 403:
                _LOGGER.error(
                    "Permission denied. Make sure your app has Group.Read.All and Tasks.Read permissions"
                )
            _LOGGER.error("HTTP Error: %s - %s", err.response.status_code, err.response.text)
            raise
        except Exception as err:
            _LOGGER.error("Error in get_plan_by_name: %s", err, exc_info=True)
            raise

    def get_plan_tasks(self, plan_name: str) -> dict[str, Any]:
        """Get all tasks for a specific plan."""
        plan = self.get_plan_by_name(plan_name)
        
        if not plan:
            return {
                "plan_name": plan_name,
                "plan_id": None,
                "open_tasks": [],
                "total_open": 0,
                "error": f"Plan '{plan_name}' not found",
            }

        plan_id = plan.get("id")
        
        try:
            # Get all tasks for the plan
            tasks_response = self._make_request(f"planner/plans/{plan_id}/tasks")
            all_tasks = tasks_response.get("value", [])
            
            # Filter for open tasks (not completed) and add assignees
            open_tasks = []
            for task in all_tasks:
                if task.get("percentComplete", 0) < 100:
                    task_id = task.get("id")
                    
                    # Get assignees for this task
                    assignees = []
                    assignments = task.get("assignments", {})
                    for user_id in assignments.keys():
                        if assignments[user_id]:  # Assignment is not null/empty
                            display_name = self.get_user_display_name(user_id)
                            assignees.append(display_name)
                    
                    open_tasks.append({
                        "id": task_id,
                        "title": task.get("title"),
                        "percentComplete": task.get("percentComplete", 0),
                        "priority": task.get("priority", 5),
                        "dueDateTime": task.get("dueDateTime"),
                        "createdDateTime": task.get("createdDateTime"),
                        "bucketId": task.get("bucketId"),
                        "assignees": assignees,
                    })
            
            return {
                "plan_name": plan_name,
                "plan_id": plan_id,
                "open_tasks": open_tasks,
                "total_open": len(open_tasks),
            }
            
        except Exception as err:
            _LOGGER.error("Error fetching tasks: %s", err)
            return {
                "plan_name": plan_name,
                "plan_id": plan_id,
                "open_tasks": [],
                "total_open": 0,
                "error": str(err),
            }

    def get_plan_buckets(self, plan_name: str) -> dict[str, Any]:
        """Return all buckets for the given plan."""
        plan = self.get_plan_by_name(plan_name)

        if not plan:
            _LOGGER.error("Cannot list buckets: Plan '%s' not found", plan_name)
            return {
                "success": False,
                "plan_name": plan_name,
                "plan_id": None,
                "buckets": [],
                "error": f"Plan '{plan_name}' not found",
            }

        plan_id = plan.get("id")

        try:
            buckets_response = self._make_request(f"planner/plans/{plan_id}/buckets")
            buckets = [
                {
                    "id": bucket.get("id"),
                    "name": bucket.get("name"),
                    "planId": bucket.get("planId"),
                    "orderHint": bucket.get("orderHint"),
                }
                for bucket in buckets_response.get("value", [])
            ]

            return {
                "success": True,
                "plan_name": plan_name,
                "plan_id": plan_id,
                "buckets": buckets,
            }

        except Exception as err:
            _LOGGER.error("Error fetching buckets for plan %s: %s", plan_name, err)
            return {
                "success": False,
                "plan_name": plan_name,
                "plan_id": plan_id,
                "buckets": [],
                "error": str(err),
            }

    def resolve_bucket_id(self, plan_name: str, bucket_value: str | None) -> dict[str, Any]:
        """Resolve a bucket name or ID to an ID for the given plan."""
        cleaned_value = (bucket_value or "").strip()
        if not cleaned_value:
            return {
                "success": False,
                "plan_name": plan_name,
                "error": "Bucket value is empty",
            }

        buckets_result = self.get_plan_buckets(plan_name)
        if not buckets_result.get("success"):
            return buckets_result

        target_lower = cleaned_value.lower()
        for bucket in buckets_result.get("buckets", []):
            bucket_id = (bucket.get("id") or "").strip()
            bucket_name = (bucket.get("name") or "").strip()

            if bucket_id and bucket_id.lower() == target_lower:
                return {
                    "success": True,
                    "plan_name": plan_name,
                    "plan_id": buckets_result.get("plan_id"),
                    "bucket_id": bucket_id,
                    "bucket_name": bucket_name,
                }

            if bucket_name and bucket_name.lower() == target_lower:
                return {
                    "success": True,
                    "plan_name": plan_name,
                    "plan_id": buckets_result.get("plan_id"),
                    "bucket_id": bucket_id,
                    "bucket_name": bucket_name,
                }

        return {
            "success": False,
            "plan_name": plan_name,
            "plan_id": buckets_result.get("plan_id"),
            "error": f"Bucket '{cleaned_value}' not found",
            "available_buckets": [
                {"id": b.get("id"), "name": b.get("name")}
                for b in buckets_result.get("buckets", [])
            ],
        }

    def create_task(
        self,
        plan_name: str,
        title: str,
        due_date: str | None = None,
        assignees: list[str] | None = None,
        priority: int = 5,
        bucket_id: str | None = None,
    ) -> dict[str, Any]:
        """Create a new task in the plan.
        
        Args:
            plan_name: Name of the plan
            title: Task title/subject
            due_date: Due date in ISO format (e.g., "2025-10-20T10:00:00Z")
            assignees: List of display names to assign the task to
            priority: Task priority (1=urgent, 5=normal, 9=low)
            bucket_id: Target Planner bucket ID (defaults to plan default bucket)
        
        Returns:
            Dictionary with task creation result
        """
        plan = self.get_plan_by_name(plan_name)
        
        if not plan:
            _LOGGER.error("Cannot create task: Plan '%s' not found", plan_name)
            return {"success": False, "error": f"Plan '{plan_name}' not found"}
        
        plan_id = plan.get("id")
        
        try:
            # Build task data
            task_data = {
                "planId": plan_id,
                "title": title,
                "priority": priority,
            }
            
            # Add due date if provided
            if due_date:
                task_data["dueDateTime"] = due_date

            if bucket_id:
                task_data["bucketId"] = bucket_id
            
            # Build assignments dictionary
            assignments = {}
            if assignees:
                for assignee_name in assignees:
                    user_id = self.get_user_id_by_name(assignee_name)
                    if user_id:
                        assignments[user_id] = {
                            "@odata.type": "#microsoft.graph.plannerAssignment",
                            "orderHint": " !"
                        }
                    else:
                        _LOGGER.warning("Could not find user '%s', skipping assignment", assignee_name)
            
            if assignments:
                task_data["assignments"] = assignments
            
            # Create the task via POST request
            url = f"{GRAPH_API_ENDPOINT}/planner/tasks"
            _LOGGER.info("Creating task '%s' in plan '%s'", title, plan_name)
            _LOGGER.debug("Task data: %s", task_data)
            
            response = requests.post(
                url,
                headers=self._get_headers(),
                json=task_data,
                timeout=30
            )
            
            if response.status_code == 401:
                # Token expired, retry with new token
                _LOGGER.debug("Token expired, refreshing...")
                self.access_token = None
                self.authenticate()
                response = requests.post(
                    url,
                    headers=self._get_headers(),
                    json=task_data,
                    timeout=30
                )
            
            response.raise_for_status()
            created_task = response.json()
            
            _LOGGER.info(
                "Successfully created task '%s' with ID: %s",
                title,
                created_task.get("id")
            )
            
            return {
                "success": True,
                "task_id": created_task.get("id"),
                "title": created_task.get("title"),
            }
            
        except requests.exceptions.HTTPError as err:
            _LOGGER.error("HTTP error creating task: %s", err)
            if err.response is not None:
                _LOGGER.error("Response body: %s", err.response.text)
            return {
                "success": False,
                "error": f"HTTP error: {err}"
            }
        except Exception as err:
            _LOGGER.error("Error creating task: %s", err, exc_info=True)
            return {
                "success": False,
                "error": str(err)
            }

    def delete_task(self, task_id: str) -> dict[str, Any]:
        """Delete a task from Planner."""

        task_url = f"{GRAPH_API_ENDPOINT}/planner/tasks/{task_id}"

        try:
            get_response = requests.get(task_url, headers=self._get_headers(), timeout=30)

            if get_response.status_code == 401:
                self.access_token = None
                self.authenticate()
                get_response = requests.get(task_url, headers=self._get_headers(), timeout=30)

            get_response.raise_for_status()
            etag = get_response.headers.get("ETag") or get_response.json().get("@odata.etag")

            if not etag:
                return {"success": False, "error": "Task ETag missing; cannot delete"}

            headers = self._get_headers()
            headers["If-Match"] = etag

            delete_response = requests.delete(task_url, headers=headers, timeout=30)

            if delete_response.status_code == 401:
                self.access_token = None
                self.authenticate()
                headers = self._get_headers()
                headers["If-Match"] = etag
                delete_response = requests.delete(task_url, headers=headers, timeout=30)

            delete_response.raise_for_status()
            _LOGGER.info("Deleted task %s", task_id)
            return {"success": True}

        except requests.exceptions.HTTPError as err:
            _LOGGER.error("HTTP error deleting task %s: %s", task_id, err)
            if err.response is not None:
                _LOGGER.error("Response body: %s", err.response.text)
            return {"success": False, "error": f"HTTP error: {err}"}
        except Exception as err:
            _LOGGER.error("Error deleting task %s: %s", task_id, err, exc_info=True)
            return {"success": False, "error": str(err)}

    def update_task(
        self,
        task_id: str,
        title: str | None = None,
        due_date: str | None = None,
        assignees: list[str] | None = None,
        percent_complete: int | None = None,
        completed: bool | None = None,
        bucket_id: str | None = None,
    ) -> dict[str, Any]:
        """Update properties on an existing task.

        Args:
            task_id: ID of the task to update
            title: Optional new title
            due_date: Optional ISO8601 due date string
            assignees: Optional list of assignee display names (set exact list)
            percent_complete: Optional numeric completion percentage (0-100)
            completed: Optional boolean to quickly mark complete/incomplete
            bucket_id: Optional Planner bucket ID to move the task into
        """

        if not any(
            value is not None
            for value in (
                title,
                due_date,
                assignees,
                percent_complete,
                completed,
                bucket_id,
            )
        ):
            return {"success": False, "error": "No update fields were provided"}

        task_url = f"{GRAPH_API_ENDPOINT}/planner/tasks/{task_id}"

        try:
            # Fetch current task to get ETag and existing assignments
            get_response = requests.get(task_url, headers=self._get_headers(), timeout=30)

            if get_response.status_code == 401:
                self.access_token = None
                self.authenticate()
                get_response = requests.get(task_url, headers=self._get_headers(), timeout=30)

            get_response.raise_for_status()
            task_data = get_response.json()

            etag = (
                get_response.headers.get("ETag")
                or task_data.get("@odata.etag")
            )

            if not etag:
                return {
                    "success": False,
                    "error": "Task ETag missing; cannot update",
                }

            update_payload: dict[str, Any] = {}

            if title is not None:
                update_payload["title"] = title

            if due_date is not None:
                update_payload["dueDateTime"] = due_date

            if percent_complete is not None:
                clamped = max(0, min(100, percent_complete))
                update_payload["percentComplete"] = clamped
            elif completed is not None:
                update_payload["percentComplete"] = 100 if completed else 0

            current_assignments = task_data.get("assignments", {})
            if assignees is not None:
                new_assignments: dict[str, Any] = {}
                resolved_ids: list[str] = []

                for name in assignees:
                    user_id = self.get_user_id_by_name(name)
                    if user_id:
                        resolved_ids.append(user_id)
                        new_assignments[user_id] = {
                            "@odata.type": "#microsoft.graph.plannerAssignment",
                            "orderHint": " !",
                        }
                    else:
                        _LOGGER.warning("Could not resolve assignee '%s'", name)

                # Remove any previous assignments not in the new list
                for existing_id in current_assignments.keys():
                    if existing_id not in resolved_ids:
                        new_assignments[existing_id] = None

                update_payload["assignments"] = new_assignments

            if bucket_id is not None:
                update_payload["bucketId"] = bucket_id

            if not update_payload:
                return {"success": False, "error": "No valid fields to update"}

            headers = self._get_headers()
            headers["If-Match"] = etag

            patch_response = requests.patch(
                task_url,
                headers=headers,
                json=update_payload,
                timeout=30,
            )

            if patch_response.status_code == 401:
                self.access_token = None
                self.authenticate()
                headers = self._get_headers()
                headers["If-Match"] = etag
                patch_response = requests.patch(
                    task_url,
                    headers=headers,
                    json=update_payload,
                    timeout=30,
                )

            patch_response.raise_for_status()

            _LOGGER.info("Updated task %s successfully", task_id)
            return {
                "success": True,
                "task_id": task_id,
                "updated_fields": list(update_payload.keys()),
            }

        except requests.exceptions.HTTPError as err:
            _LOGGER.error("HTTP error updating task %s: %s", task_id, err)
            if err.response is not None:
                _LOGGER.error("Response body: %s", err.response.text)
            return {"success": False, "error": f"HTTP error: {err}"}
        except Exception as err:
            _LOGGER.error("Error updating task %s: %s", task_id, err, exc_info=True)
            return {"success": False, "error": str(err)}
