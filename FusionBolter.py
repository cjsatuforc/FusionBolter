# Author-Patrick Rainsberry
# Description-Calculates bolt holes

# Importing  Fusion Commands
from .FusionBolterCommand import FusionBolterCommand, FusionBolterFindCommand
from .FusionBolterPrepCommand import FusionBolterPrepCommand
from .FusionBolterConfigCommand import FusionBolterConfigCommand


commands = []
command_definitions = []

# Define parameters for 1st command
cmd = {
    'cmd_name': 'Add Through Hole Fastener Stack',
    'cmd_description': 'Adds Bolt, washers and nut to a through hole',
    'cmd_id': 'cmdID_FusionBolterCommand_1',
    'cmd_resources': './resources',
    'workspace': 'FusionSolidEnvironment',
    'toolbar_panel_id': 'Bolter',
    'command_promoted': True,
    'class': FusionBolterCommand
}
command_definitions.append(cmd)

# Define parameters for 1st command
cmd = {
    'cmd_name': 'Count Holes',
    'cmd_description': 'Calculate Bolt Holes with lengths',
    'cmd_id': 'cmdID_FusionBolterCommand_2',
    'cmd_resources': './resources',
    'workspace': 'FusionSolidEnvironment',
    'toolbar_panel_id': 'Bolter',
    'command_promoted': False,
    'class': FusionBolterFindCommand
}
command_definitions.append(cmd)

# Define parameters for 1st command
cmd = {
    'cmd_name': 'Fusion Bolter - Prepare Hardware',
    'cmd_description': 'Calculate Bolt Holes',
    'cmd_id': 'cmdID_FusionBolterPrepCommand',
    'cmd_resources': './resources',
    'workspace': 'FusionSolidEnvironment',
    'toolbar_panel_id': 'Bolter',
    'class': FusionBolterPrepCommand
}
command_definitions.append(cmd)

# Define parameters for 1st command
cmd = {
    'cmd_name': 'Configure Fusion Bolter',
    'cmd_description': 'Initialize Fusion Bolter',
    'cmd_id': 'cmdID_FusionBolterConfigCommand',
    'cmd_resources': './resources',
    'workspace': 'FusionSolidEnvironment',
    'toolbar_panel_id': 'Bolter',
    'class': FusionBolterConfigCommand
}
command_definitions.append(cmd)

# Set to True to display various useful messages when debugging your app
debug = False


# Don't change anything below here:
for cmd_def in command_definitions:
    command = cmd_def['class'](cmd_def, debug)
    commands.append(command)


def run(context):
    for run_command in commands:
        run_command.on_run()


def stop(context):
    for stop_command in commands:
        stop_command.on_stop()