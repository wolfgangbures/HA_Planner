"""Config flow for Microsoft Planner integration."""
from __future__ import annotations

import logging
from typing import Any

import voluptuous as vol

from homeassistant import config_entries
from homeassistant.core import HomeAssistant
from homeassistant.data_entry_flow import FlowResult
from homeassistant.exceptions import HomeAssistantError

from .const import DOMAIN, CONF_CLIENT_ID, CONF_CLIENT_SECRET, CONF_TENANT_ID, CONF_PLAN_NAME
from .planner_api import PlannerAPI

_LOGGER = logging.getLogger(__name__)

# Default values intentionally blank to avoid leaking credentials
DEFAULT_CLIENT_ID = ""
DEFAULT_TENANT_ID = ""
DEFAULT_CLIENT_SECRET = ""
DEFAULT_PLAN_NAME = ""

STEP_USER_DATA_SCHEMA = vol.Schema(
    {
        vol.Required(CONF_CLIENT_ID, default=DEFAULT_CLIENT_ID): str,
        vol.Required(CONF_CLIENT_SECRET, default=DEFAULT_CLIENT_SECRET): str,
        vol.Required(CONF_TENANT_ID, default=DEFAULT_TENANT_ID): str,
        vol.Required(CONF_PLAN_NAME, default=DEFAULT_PLAN_NAME): str,
    }
)


async def validate_input(hass: HomeAssistant, data: dict[str, Any]) -> dict[str, Any]:
    """Validate the user input allows us to connect.

    Data has the keys from STEP_USER_DATA_SCHEMA with values provided by the user.
    """
    api = PlannerAPI(
        data[CONF_CLIENT_ID],
        data[CONF_CLIENT_SECRET],
        data[CONF_TENANT_ID],
    )

    # Test authentication
    try:
        await hass.async_add_executor_job(api.authenticate)
        _LOGGER.info("Authentication successful for tenant: %s", data[CONF_TENANT_ID])
    except Exception as err:
        _LOGGER.error("Authentication failed: %s", err)
        raise InvalidAuth from err

    # Test if we can find the plan
    try:
        _LOGGER.info("Attempting to find plan: %s", data[CONF_PLAN_NAME])
        plan = await hass.async_add_executor_job(
            api.get_plan_by_name, data[CONF_PLAN_NAME]
        )
        if not plan:
            # List available plans for debugging
            all_plans = await hass.async_add_executor_job(api.list_all_plans)
            available_plans = [p.get("title") for p in all_plans]
            _LOGGER.error(
                "Plan '%s' not found. Available plans: %s", 
                data[CONF_PLAN_NAME], 
                available_plans
            )
            raise CannotConnect(f"Plan '{data[CONF_PLAN_NAME]}' not found. Available plans: {available_plans}")
    except CannotConnect:
        raise
    except Exception as err:
        _LOGGER.error("Failed to retrieve plan: %s", err, exc_info=True)
        raise CannotConnect from err

    # Return info that you want to store in the config entry.
    return {"title": f"Planner: {data[CONF_PLAN_NAME]}"}


class ConfigFlow(config_entries.ConfigFlow, domain=DOMAIN):
    """Handle a config flow for Microsoft Planner."""

    VERSION = 1

    async def async_step_user(
        self, user_input: dict[str, Any] | None = None
    ) -> FlowResult:
        """Handle the initial step."""
        errors: dict[str, str] = {}
        
        if user_input is not None:
            try:
                info = await validate_input(self.hass, user_input)
            except CannotConnect:
                errors["base"] = "cannot_connect"
            except InvalidAuth:
                errors["base"] = "invalid_auth"
            except Exception:  # pylint: disable=broad-except
                _LOGGER.exception("Unexpected exception")
                errors["base"] = "unknown"
            else:
                # Create a unique ID based on tenant and plan name
                await self.async_set_unique_id(
                    f"{user_input[CONF_TENANT_ID]}_{user_input[CONF_PLAN_NAME]}"
                )
                self._abort_if_unique_id_configured()
                
                return self.async_create_entry(title=info["title"], data=user_input)

        return self.async_show_form(
            step_id="user", data_schema=STEP_USER_DATA_SCHEMA, errors=errors
        )


class CannotConnect(HomeAssistantError):
    """Error to indicate we cannot connect."""


class InvalidAuth(HomeAssistantError):
    """Error to indicate there is invalid auth."""
