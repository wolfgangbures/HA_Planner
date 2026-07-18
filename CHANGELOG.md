# Changelog

All notable changes to this project will be documented in this file.

## [1.3.0-beta] - 2026-06-14
### Added
- `planner.add_task_comment` service to append or overwrite the Notes field of an existing Planner task.
  - `task_id` (required): ID of the task to update.
  - `comment` (required): The text to write into the task's Notes.
  - `append` (optional, default `true`): When `true` the comment is appended to existing notes; when `false` the existing notes are replaced.

### Fixed
- Added Home Assistant custom integration brand assets under `custom_components/planner/brand/` so the integration icon renders correctly in Home Assistant and satisfies HACS brand checks.

## [0.4.0] - 2025-12-30
### Added
- Bucket-aware task creation and updates, including optional `bucket` name resolution and direct `bucket_id` targeting.
- `planner.list_buckets` service to expose the buckets available in a plan.
- Planner API helpers to fetch buckets and convert bucket names into IDs.
- Documentation updates covering the new bucket workflow and service fields.

## [0.3.0] - 2025-12-12
### Added
- Home Assistant todo entity support alongside the task sensor.
- Task creation, update, and deletion through the todo platform.
