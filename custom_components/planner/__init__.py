"""The Microsoft Planner integration."""
from __future__ import annotations

import logging
from datetime import timedelta

from homeassistant.config_entries import ConfigEntry
from homeassistant.const import Platform
from homeassistant.core import HomeAssistant
from homeassistant.exceptions import ConfigEntryNotReady
from homeassistant.helpers.update_coordinator import DataUpdateCoordinator, UpdateFailed

from .const import DOMAIN
from .planner_api import PlannerAPI

_LOGGER = logging.getLogger(__name__)

PLATFORMS: list[Platform] = [Platform.SENSOR, Platform.TODO]

UPDATE_INTERVAL = timedelta(minutes=5)


async def async_setup_entry(hass: HomeAssistant, entry: ConfigEntry) -> bool:
    """Set up Microsoft Planner from a config entry."""
    hass.data.setdefault(DOMAIN, {})

    client_id = entry.data["client_id"]
    client_secret = entry.data["client_secret"]
    tenant_id = entry.data["tenant_id"]
    plan_name = entry.data["plan_name"]

    api = PlannerAPI(client_id, client_secret, tenant_id)

    # Test authentication
    try:
        await hass.async_add_executor_job(api.authenticate)
    except Exception as err:
        _LOGGER.error("Failed to authenticate with Microsoft Graph: %s", err)
        raise ConfigEntryNotReady from err

    async def async_update_data():
        """Fetch data from API."""
        try:
            return await hass.async_add_executor_job(
                api.get_plan_tasks, plan_name
            )
        except Exception as err:
            raise UpdateFailed(f"Error communicating with API: {err}") from err

    coordinator = DataUpdateCoordinator(
        hass,
        _LOGGER,
        name=f"{DOMAIN}_{plan_name}",
        update_method=async_update_data,
        update_interval=UPDATE_INTERVAL,
    )

    # Fetch initial data
    await coordinator.async_config_entry_first_refresh()

    hass.data[DOMAIN][entry.entry_id] = {
        "coordinator": coordinator,
        "api": api,
        "plan_name": plan_name,
    }

    await hass.config_entries.async_forward_entry_setups(entry, PLATFORMS)

    # Register services
    async def handle_create_task(call):
        """Handle the create_task service call."""
        title = call.data.get("title")
        due_date = call.data.get("due_date")
        assignees = call.data.get("assignees", [])
        priority = call.data.get("priority", 5)
        bucket_id = call.data.get("bucket_id")
        bucket_value = call.data.get("bucket")
        
        # Use the first configured plan if not specified
        target_plan = call.data.get("plan_name", plan_name)

        resolved_bucket_id = bucket_id
        if not resolved_bucket_id and bucket_value:
            lookup = await hass.async_add_executor_job(
                api.resolve_bucket_id,
                target_plan,
                bucket_value,
            )

            if not lookup.get("success"):
                _LOGGER.error(
                    "Failed to resolve bucket '%s' for plan '%s': %s",
                    bucket_value,
                    target_plan,
                    lookup.get("error"),
                )
                return lookup

            resolved_bucket_id = lookup.get("bucket_id")
        
        _LOGGER.info("Service call to create task: %s", title)
        
        result = await hass.async_add_executor_job(
            api.create_task,
            target_plan,
            title,
            due_date,
            assignees,
            priority,
            resolved_bucket_id,
        )
        
        if result.get("success"):
            _LOGGER.info("Task created successfully: %s", result.get("task_id"))
            # Refresh coordinator to show new task
            await coordinator.async_request_refresh()
        else:
            _LOGGER.error("Failed to create task: %s", result.get("error"))
        
        return result

    hass.services.async_register(DOMAIN, "create_task", handle_create_task)

    async def handle_update_task(call):
        """Handle the update_task service call."""
        task_id = call.data.get("task_id")
        title = call.data.get("title")
        due_date = call.data.get("due_date")
        assignees = call.data.get("assignees")
        percent_complete = call.data.get("percent_complete")
        completed = call.data.get("completed")
        bucket_id = call.data.get("bucket_id")
        bucket_value = call.data.get("bucket")
        target_plan = call.data.get("plan_name", plan_name)

        resolved_bucket_id = bucket_id
        if not resolved_bucket_id and bucket_value:
            lookup = await hass.async_add_executor_job(
                api.resolve_bucket_id,
                target_plan,
                bucket_value,
            )

            if not lookup.get("success"):
                _LOGGER.error(
                    "Failed to resolve bucket '%s' for plan '%s': %s",
                    bucket_value,
                    target_plan,
                    lookup.get("error"),
                )
                return lookup

            resolved_bucket_id = lookup.get("bucket_id")

        if not task_id:
            _LOGGER.error("update_task service requires task_id")
            return {"success": False, "error": "task_id missing"}

        _LOGGER.info("Service call to update task: %s", task_id)

        result = await hass.async_add_executor_job(
            api.update_task,
            task_id,
            title,
            due_date,
            assignees,
            percent_complete,
            completed,
            resolved_bucket_id,
        )

        if result.get("success"):
            await coordinator.async_request_refresh()
        else:
            _LOGGER.error(
                "Failed to update task %s: %s",
                task_id,
                result.get("error"),
            )

        return result

    hass.services.async_register(DOMAIN, "update_task", handle_update_task)

    async def handle_list_buckets(call):
        """Handle listing buckets for a plan."""
        target_plan = call.data.get("plan_name", plan_name)

        result = await hass.async_add_executor_job(
            api.get_plan_buckets,
            target_plan,
        )

        if not result.get("success"):
            _LOGGER.error(
                "Failed to list buckets for plan '%s': %s",
                target_plan,
                result.get("error"),
            )

        return result

    hass.services.async_register(DOMAIN, "list_buckets", handle_list_buckets)

    return True


async def async_unload_entry(hass: HomeAssistant, entry: ConfigEntry) -> bool:
    """Unload a config entry."""
    if unload_ok := await hass.config_entries.async_unload_platforms(entry, PLATFORMS):
        hass.data[DOMAIN].pop(entry.entry_id)

    return unload_ok
