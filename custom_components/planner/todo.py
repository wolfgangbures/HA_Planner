"""Todo platform for the Microsoft Planner integration."""
from __future__ import annotations

from datetime import datetime, timezone
import logging
from typing import Any

from homeassistant.components.todo import (
    TodoItem,
    TodoItemPriority,
    TodoItemStatus,
    TodoListEntity,
    TodoListEntityFeature,
)
from homeassistant.config_entries import ConfigEntry
from homeassistant.core import HomeAssistant
from homeassistant.helpers.entity_platform import AddEntitiesCallback
from homeassistant.helpers.update_coordinator import CoordinatorEntity

from .const import DOMAIN

_LOGGER = logging.getLogger(__name__)


async def async_setup_entry(
    hass: HomeAssistant,
    entry: ConfigEntry,
    async_add_entities: AddEntitiesCallback,
) -> None:
    """Set up the todo platform."""
    data = hass.data[DOMAIN][entry.entry_id]

    async_add_entities(
        [
            PlannerTodoList(
                data["coordinator"],
                data["api"],
                entry,
                data["plan_name"],
            )
        ]
    )


class PlannerTodoList(CoordinatorEntity, TodoListEntity):
    """Representation of the Planner todo list."""

    _attr_has_entity_name = True
    _attr_supported_features = (
        TodoListEntityFeature.CREATE_TODO_ITEM
        | TodoListEntityFeature.UPDATE_TODO_ITEM
        | TodoListEntityFeature.DELETE_TODO_ITEM
        | TodoListEntityFeature.SET_DUE_DATE
    )

    def __init__(self, coordinator, api, entry: ConfigEntry, plan_name: str) -> None:
        """Initialize the todo entity."""
        super().__init__(coordinator)
        self._api = api
        self._plan_name = plan_name
        self._entry_id = entry.entry_id
        self._attr_unique_id = f"{entry.entry_id}_todo"
        self._attr_name = f"{plan_name} Tasks"

    @property
    def device_info(self) -> dict[str, Any]:
        """Return device information shared with the sensor."""
        return {
            "identifiers": {(DOMAIN, self._entry_id)},
            "name": f"Planner: {self._plan_name}",
            "manufacturer": "Microsoft",
            "model": "Planner",
        }

    async def async_get_todo_items(self) -> list[TodoItem]:
        """Return the current Planner tasks as todo items."""
        data = self.coordinator.data or {}
        tasks: list[dict[str, Any]] = data.get("open_tasks", [])
        items: list[TodoItem] = []

        for task in tasks:
            items.append(
                TodoItem(
                    summary=task.get("title", "Unnamed task"),
                    uid=task.get("id"),
                    status=self._status_from_task(task),
                    due=self._parse_due_date(task.get("dueDateTime")),
                    description=self._build_description(task),
                    priority=self._priority_from_planner(task.get("priority")),
                )
            )

        return items

    async def async_create_todo_item(self, item: TodoItem) -> TodoItem | None:
        """Create a Planner task from a todo item."""
        title = item.summary or "New Task"
        due_date = self._format_due_date(item.due)
        priority = self._priority_to_planner(item.priority)

        result = await self.hass.async_add_executor_job(
            self._api.create_task,
            self._plan_name,
            title,
            due_date,
            None,
            priority,
        )

        if not result.get("success"):
            _LOGGER.error("Failed to create Planner task: %s", result.get("error"))
            return None

        await self.coordinator.async_request_refresh()

        return TodoItem(
            summary=title,
            uid=result.get("task_id"),
            status=TodoItemStatus.NEEDS_ACTION,
            due=item.due,
            priority=item.priority,
        )

    async def async_update_todo_item(self, item: TodoItem) -> TodoItem | None:
        """Update an existing Planner task."""
        if not item.uid:
            _LOGGER.error("Cannot update Planner task without uid")
            return None

        due_date = self._format_due_date(item.due)
        completed = item.status == TodoItemStatus.COMPLETED

        result = await self.hass.async_add_executor_job(
            self._api.update_task,
            item.uid,
            item.summary,
            due_date,
            None,
            None,
            completed,
        )

        if not result.get("success"):
            _LOGGER.error("Failed to update Planner task %s: %s", item.uid, result.get("error"))
            return None

        await self.coordinator.async_request_refresh()

        return item

    async def async_delete_todo_item(self, uid: str) -> None:
        """Delete a Planner task when a todo item is removed."""
        result = await self.hass.async_add_executor_job(self._api.delete_task, uid)

        if not result.get("success"):
            _LOGGER.error("Failed to delete Planner task %s: %s", uid, result.get("error"))
            return

        await self.coordinator.async_request_refresh()

    @staticmethod
    def _build_description(task: dict[str, Any]) -> str | None:
        assignees = task.get("assignees")
        if assignees:
            return ", ".join(assignees)
        return None

    @staticmethod
    def _parse_due_date(value: str | None) -> datetime | None:
        if not value:
            return None
        try:
            if value.endswith("Z"):
                value = value.replace("Z", "+00:00")
            return datetime.fromisoformat(value)
        except ValueError:
            _LOGGER.debug("Could not parse due date: %s", value)
            return None

    @staticmethod
    def _format_due_date(value: datetime | None) -> str | None:
        if not value:
            return None
        if value.tzinfo is None:
            value = value.replace(tzinfo=timezone.utc)
        value = value.astimezone(timezone.utc)
        iso_value = value.isoformat().replace("+00:00", "Z")
        return iso_value

    @staticmethod
    def _priority_from_planner(value: int | None) -> TodoItemPriority:
        if value is None:
            return TodoItemPriority.NORMAL
        if value <= 3:
            return TodoItemPriority.HIGH
        if value <= 6:
            return TodoItemPriority.NORMAL
        return TodoItemPriority.LOW

    @staticmethod
    def _priority_to_planner(priority: TodoItemPriority | None) -> int:
        if priority == TodoItemPriority.HIGH:
            return 1
        if priority == TodoItemPriority.LOW:
            return 9
        return 5

    @staticmethod
    def _status_from_task(task: dict[str, Any]) -> TodoItemStatus:
        if task.get("percentComplete", 0) >= 100:
            return TodoItemStatus.COMPLETED
        return TodoItemStatus.NEEDS_ACTION
