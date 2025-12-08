# Microsoft Planner Integration for Home Assistant

This custom integration allows you to retrieve and monitor tasks from Microsoft Planner plans in Home Assistant.
This code was generatde by GPT5.1-Codex.

## Features

- ✅ Retrieve open tasks from a Microsoft Planner plan
- ✅ Create new tasks with title, due date, and assignees
- ✅ Update task details (title, due date, assignees, completion)
- ✅ Display total number of open tasks
- ✅ Show task details including title, priority, due date, assignees, and completion percentage
- ✅ Filter and count high-priority tasks
- ✅ Automatic updates every 5 minutes
- ✅ Use in voice intents to get task overviews and create tasks

## Prerequisites

Before using this integration, you need to set up an Azure App Registration:

1. **Create an Azure App Registration:**
   - Go to the [Azure Portal](https://portal.azure.com/)
   - Navigate to "Azure Active Directory" → "App registrations"
   - Click "New registration"
   - Give it a name (e.g., "Home Assistant Planner")
   - Select "Accounts in this organizational directory only"
   - Click "Register"

2. **Configure API Permissions:**
   - In your app registration, go to "API permissions"
   - Click "Add a permission" → "Microsoft Graph" → "Application permissions"
   - Add these permissions:
     - `Tasks.ReadWrite` (required to read and create tasks)
     - `Group.Read.All` (needed to find plans)
     - `User.Read.All` (needed to resolve assignee names)
   - Click "Grant admin consent" (requires admin privileges)

3. **Create a Client Secret:**
   - Go to "Certificates & secrets"
   - Click "New client secret"
   - Add a description and set expiration
   - Copy the **Value** (not the Secret ID) - you'll need this for configuration

4. **Get your IDs:**
   - **Client ID**: Found on the app's "Overview" page (Application/client ID)
   - **Tenant ID**: Found on the app's "Overview" page (Directory/tenant ID)

## Installation

1. Copy the `planner` folder to your Home Assistant `custom_components` directory
2. Restart Home Assistant
3. Go to Settings → Devices & Services → Add Integration
4. Search for "Microsoft Planner"
5. Enter your credentials:
   - **Client ID**: Your Azure app's Application (client) ID
   - **Client Secret**: The client secret value you created
   - **Tenant ID**: Your Azure AD Tenant ID
   - **Plan Name**: The exact name of the Planner plan you want to monitor

## Usage

### Sensor Entity

After configuration, a sensor entity will be created:
- **Entity ID**: `sensor.planner_<plan_name>_open_tasks`
- **State**: Number of open tasks
- **Attributes**:
  - `plan_name`: Name of the monitored plan
  - `plan_id`: Microsoft Graph ID of the plan
  - `total_open_tasks`: Total number of open tasks
  - `high_priority_tasks`: Count of high-priority tasks (priority 1-3)
  - `tasks`: List of all open tasks with details
  - `last_updated`: Timestamp of last update

### Example Task Attributes

```yaml
tasks:
  - title: "Complete project documentation"
    priority: 1
    percent_complete: 50
    due_date: "2025-10-20T00:00:00Z"
    assignees: ["Wolfgang", "Maria"]
  - title: "Review pull request"
    priority: 3
    percent_complete: 0
    assignees: ["Wolfgang"]
```

### Voice Intent Example

You can use this sensor in voice intents to get task overviews. Example intent script:

```yaml
# intent_scripts/planner_intents.yaml
GetPlannerTasks:
  speech:
    text: >
      You have {{ states('sensor.planner_my_plan_open_tasks') }} open tasks.
      {% set high_priority = state_attr('sensor.planner_my_plan_open_tasks', 'high_priority_tasks') %}
      {% if high_priority > 0 %}
      There are {{ high_priority }} high priority tasks.
      {% endif %}
```

### Creating Tasks

You can create new tasks using the `planner.create_task` service:

#### Via Developer Tools → Services:
```yaml
service: planner.create_task
data:
  title: "Buy groceries"
  due_date: "2025-10-25T18:00:00Z"
  assignees:
    - Wolfgang
    - Maria
  priority: 5
```

### Updating Tasks

Use the `planner.update_task` service to change properties on an existing task.

#### Via Developer Tools → Services:
```yaml
service: planner.update_task
data:
  task_id: "AAMkADk3..."
  title: "Finish documentation"
  due_date: "2025-10-28T18:00:00Z"
  assignees:
    - Wolfgang
  completed: true
```

You can supply any combination of `title`, `due_date`, `assignees`, `percent_complete`, or `completed`. Fields you leave out remain unchanged. Passing an empty list for `assignees` removes every assignment from the task.

### Completing Tasks via Voice

Pair the Planner integration with Assist by enabling the `CompletePlannerTask` intent. The helper script `script.complete_planner_task` fetches the open tasks from `sensor.planner_aufgaben_open_tasks`, lets the AI match a spoken utterance ("I finished the status report"), and then calls `planner.update_task` with `completed: true` when the match confidence is high enough. If the model is unsure it will ask for clarification instead of closing the wrong task.


#### In Automations:
```yaml
automation:
  - alias: "Create task from voice command"
    trigger:
      - platform: event
        event_type: call_service
        event_data:
          domain: conversation
          service: process
    action:
      - service: planner.create_task
        data:
          title: "{{ trigger.event.data.text }}"
          assignees:
            - Wolfgang
          priority: 3
```

#### In Scripts:
```yaml
script:
  add_shopping_task:
    alias: "Add Shopping Task"
    sequence:
      - service: planner.create_task
        data:
          title: "{{ item }}"
          due_date: "{{ now().date() + timedelta(days=1) }}T18:00:00Z"
          assignees:
            - Maria
          priority: 5
```

### Automation Example

```yaml
automation:
  - alias: "Notify about high priority tasks"
    trigger:
      - platform: state
        entity_id: sensor.planner_my_plan_open_tasks
        attribute: high_priority_tasks
    condition:
      - condition: template
        value_template: "{{ state_attr('sensor.planner_my_plan_open_tasks', 'high_priority_tasks') | int > 0 }}"
    action:
      - service: notify.mobile_app
        data:
          message: >
            You have {{ state_attr('sensor.planner_my_plan_open_tasks', 'high_priority_tasks') }} 
            high priority tasks in your planner!
```

## Troubleshooting

### Authentication Issues

- Verify your Client ID, Client Secret, and Tenant ID are correct
- Ensure the client secret hasn't expired
- Check that admin consent was granted for the required permissions

### Plan Not Found

- Verify the plan name is exactly as it appears in Microsoft Planner (case-sensitive)
- Ensure your Azure app has the necessary permissions
- Check that the plan is accessible with the credentials provided

### Permission Errors

If you see 403 Forbidden errors:
- Verify that `Tasks.Read` and `Group.Read.All` permissions are granted
- Ensure admin consent was granted (blue checkmark in Azure Portal)
- Wait a few minutes after granting permissions for changes to propagate

## Future Enhancements

Planned features for future versions:
- Create new tasks via services
- Update existing tasks (mark as complete, change priority)
- Support for task buckets
- Task assignment information
- Task comments and attachments

## License

This integration is provided as-is for personal use.

## Support

For issues, questions, or feature requests, please check the Home Assistant logs for detailed error messages.
