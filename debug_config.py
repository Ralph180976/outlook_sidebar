import sys
from sidebar.core.config_manager import ConfigManager

cm = ConfigManager()
print("reminder_show_tasks:", cm.reminder_show_tasks)
print("reminder_task_dates:", cm.reminder_task_dates)
