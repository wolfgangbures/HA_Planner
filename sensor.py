"""Sensor platform for Microsoft Planner integration."""
from __future__ import annotations

from datetime import datetime
import logging

from homeassistant.components.sensor import SensorEntity, SensorStateClass
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
    """Set up the sensor platform."""
    coordinator = hass.data[DOMAIN][entry.entry_id]["coordinator"]
    plan_name = hass.data[DOMAIN][entry.entry_id]["plan_name"]

    async_add_entities([PlannerTasksSensor(coordinator, entry, plan_name)])


class PlannerTasksSensor(CoordinatorEntity, SensorEntity):
    """Representation of a Planner Tasks sensor."""

    _attr_has_entity_name = True
    _attr_state_class = SensorStateClass.MEASUREMENT

    def __init__(self, coordinator, entry: ConfigEntry, plan_name: str) -> None:
        """Initialize the sensor."""
        super().__init__(coordinator)
        self._attr_unique_id = f"{entry.entry_id}_tasks"
        self._attr_name = "Open Tasks"
        self._plan_name = plan_name
        self._entry_id = entry.entry_id

    @property
    def device_info(self):
        """Return device information about this entity."""
        return {
            "identifiers": {(DOMAIN, self._entry_id)},
            "name": f"Planner: {self._plan_name}",
            "manufacturer": "Microsoft",
            "model": "Planner",
        }

    @property
    def native_value(self) -> int:
        """Return the state of the sensor."""
        if self.coordinator.data:
            return self.coordinator.data.get("total_open", 0)
        return 0

    @property
    def extra_state_attributes(self) -> dict:
        """Return the state attributes."""
        if not self.coordinator.data:
            return {}

        data = self.coordinator.data
        attributes = {
            "plan_name": data.get("plan_name"),
            "plan_id": data.get("plan_id"),
            "total_open_tasks": data.get("total_open", 0),
        }

        # Add error if present
        if "error" in data:
            attributes["error"] = data["error"]

        # Add task details
        tasks = data.get("open_tasks", [])
        if tasks:
            # Sort tasks by priority (lower number = higher priority)
            sorted_tasks = sorted(tasks, key=lambda x: x.get("priority", 5))
            
            # Create a list of task summaries
            task_list = []
            for task in sorted_tasks:
                task_info = {
                    "id": task.get("id"),
                    "title": task.get("title"),
                    "priority": task.get("priority", 5),
                    "percent_complete": task.get("percentComplete", 0),
                }
                
                # Add due date if present
                if task.get("dueDateTime"):
                    task_info["due_date"] = task["dueDateTime"]
                
                # Add assignees if present
                if task.get("assignees"):
                    task_info["assignees"] = task["assignees"]
                
                task_list.append(task_info)
            
            attributes["tasks"] = task_list
            
            # Add high priority task count (priority 1-3)
            high_priority_count = sum(
                1 for task in tasks if task.get("priority", 5) <= 3
            )
            attributes["high_priority_tasks"] = high_priority_count

        attributes["last_updated"] = datetime.now().isoformat()

        return attributes

    @property
    def icon(self) -> str:
        """Return the icon to use in the frontend."""
        return "mdi:clipboard-check-multiple-outline"

    @property
    def available(self) -> bool:
        """Return True if entity is available."""
        return self.coordinator.last_update_success
