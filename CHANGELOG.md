# Changelog

All notable changes to this project will be documented in this file.

## [1.2.3] - 2026-06-14
### Fixed
- Guard against `AttributeError` when `HTTPError.response` is `None` in
  `list_all_plans()` and `get_plan_by_name()`. Previously a network-level
  error (e.g. connection reset) that produced an `HTTPError` without an
  attached response object would raise an unhandled `AttributeError` instead
  of the expected log message, masking the real root cause.

## [1.2.2] - 2025-12-30
### Changed
- Various stability and permission-handling improvements.

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
