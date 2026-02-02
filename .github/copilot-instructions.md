# Copilot Instructions

## Big Picture
- Home Assistant custom integration surfaces Microsoft Planner tasks via sensor + todo entities; REST + Graph operations live in [custom_components/planner](custom_components/planner).
- Polling cadence is fixed at 5 min through DataUpdateCoordinator, so favor push refreshes (see [custom_components/planner/__init__.py](custom_components/planner/__init__.py)).
- README describes Azure setup and HA usage; keep docs aligned when behavior changes (see [README.md](README.md)).

## Architecture & Data Flow
- Entry flow: config entry → PlannerAPI auth/test → coordinator → entity setup in [custom_components/planner/__init__.py](custom_components/planner/__init__.py).
- [custom_components/planner/planner_api.py](custom_components/planner/planner_api.py) wraps Graph calls (`msal` auth, retry-on-401, bucket/name resolution); any new network call should reuse `_make_request` or follow its timeout/refresh pattern.
- Config validation in [custom_components/planner/config_flow.py](custom_components/planner/config_flow.py) authenticates and ensures the plan exists by enumerating groups; preserve this to fail early on mis-typed plan names.
- Coordinator payload schema: `{plan_name, plan_id, open_tasks[], total_open, error?}`; both sensor and todo entities read directly from it, so extend carefully.

## Entities & Services
- Sensor logic in [custom_components/planner/sensor.py](custom_components/planner/sensor.py) exposes count + structured attributes (`tasks`, `high_priority_tasks`, `last_updated`); keep attribute contract stable for dashboards/intents.
- Todo bridge in [custom_components/planner/todo.py](custom_components/planner/todo.py) mirrors open tasks; `TodoItem.description` is the assignee list, and due dates must stay in UTC `Z` format.
- Services `create_task`, `update_task`, `list_buckets` are registered in [custom_components/planner/__init__.py](custom_components/planner/__init__.py) and documented in [custom_components/planner/services.yaml](custom_components/planner/services.yaml); update both when adding/changing fields.
- Bucket handling: prefer `bucket` (human name) and let `resolve_bucket_id` translate; only accept IDs directly when the user supplies them.

## Implementation Patterns
- All Graph calls are synchronous and run inside `hass.async_add_executor_job`; never await `requests` directly from the event loop.
- Always refresh the coordinator (`await coordinator.async_request_refresh()`) after mutating tasks so sensor/todo stay in sync.
- User lookup relies on display names/UPNs via `get_user_id_by_name`; when adding assignment features, reuse this helper to keep matching behavior consistent.
- Planner deletions/updates require the current ETag; follow the fetch→If-Match workflow in `delete_task`/`update_task` to avoid 412s.
- Error reporting returns `{success: bool, error?: str}`; service handlers log and bubble that dict, so stick to the same envelope for new API helpers.

## Developer Workflow
- Install deps via Home Assistant (manifest already pins `msal==1.34.0`); no standalone requirements.txt.
- For manual testing, copy `custom_components/planner` into your HA instance, restart HA, then use Settings → Devices & Services → Reload to pick up changes.
- Exercise services through Developer Tools → Services; use the YAML snippets in [README.md](README.md) for sample payloads.
- When adding translation-visible fields, update `strings.json` and `translations/en.json` alongside `services.yaml` to keep the UI form names aligned.

## Gotchas
- Plan discovery iterates every group via Graph; keep logging level sane and avoid excessive retries to prevent throttling.
- Coordinator data is shared across entities; avoid mutating `coordinator.data` structures in-place inside entities to prevent subtle cache bugs.
- Authentication failures raise `ConfigEntryNotReady`; catch network issues early rather than letting the coordinator loop fail silently.
- Keep file encoding ASCII; Microsoft Graph payloads should stay UTF-8-safe but avoid inserting emojis in log messages.

## GitHub
- Always split branches by feature/fix for PRs; avoid working directly on `main`.
- Write clear, descriptive commit messages; reference related issues/PRs when applicable.
- Provide the links to the PR in chat for the users to test and review.
- Tag releases in GitHub matching the version in `custom_components/planner/manifest.json`.
- increment the version in `manifest.json` for every PR that changes functionality.