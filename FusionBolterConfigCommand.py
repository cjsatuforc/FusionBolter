import webbrowser
import shutil
import os

from .Fusion360Utilities.Fusion360Utilities import get_app_objects
from .Fusion360Utilities.Fusion360CommandBase import Fusion360CommandBase

from .SheetsService import sheets_copy_spreadsheet

from .BolterUtilities import write_settings, read_settings, get_default_model_dir

# default_source_sheet_id = '1csJofpAG3oD204rxkVoYzI5WGjvrMeWsaq2oe2J7hig'
default_source_sheet_id = '1HgvACfX5B7QFzvgQC0Z76pPZ6jCNw8hxVzgbDc0AkCA'


# Copy sample files to configuration directory
def copy_files():

    local_model_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'models')
    model_list = os.listdir(local_model_path)

    model_dir = get_default_model_dir()

    for model in model_list:
        src = os.path.join(local_model_path, model)
        shutil.copy2(src, model_dir)


class FusionBolterConfigCommand(Fusion360CommandBase):

    def on_execute(self, command, inputs, args, input_values):

        if input_values['new_or_existing'] == 'Create New Sheet':
            spreadsheet = sheets_copy_spreadsheet(default_source_sheet_id)
            new_id = spreadsheet['spreadsheetId']

        elif input_values['new_or_existing'] == 'Link to Existing Sheet':
            new_id = input_values['existing_sheet_id']
            # TODO add check to see if sheet id is valid

        else:
            raise Exception('Something went wrong creating sheet with your inputs')

        write_settings(new_id)

        copy_files()

        url = 'https://docs.google.com/spreadsheets/d/%s/edit#gid=0' % new_id

        webbrowser.open(url, new=2)

    # Update dialog based on user selections
    def on_input_changed(self, command_, command_inputs, changed_input, input_values):

        if changed_input.id == 'new_or_existing':
            if changed_input.selectedItem.name == 'Create New Sheet':

                input_values['instructions_input'].isVisible = False
                input_values['existing_sheet_id_input'].isVisible = False

            elif changed_input.selectedItem.name == 'Link to Existing Sheet':
                input_values['instructions_input'].isVisible = True
                input_values['existing_sheet_id_input'].isVisible = True
                input_values['existing_sheet_id_input'].value = read_settings()

    def on_create(self, command, command_inputs):

        command.setDialogInitialSize(600, 800)

        command_inputs.addTextBoxCommandInput('new_title', '', '<b>Create New Sheet or Link to existing?</b>', 1, True)
        new_option_group = command_inputs.addRadioButtonGroupCommandInput('new_or_existing')

        new_option_group.listItems.add('Create New Sheet', True)
        new_option_group.listItems.add('Link to Existing Sheet', False)

        instructions_text = '<br> <b>You need the spreadsheetID from your existing sheets document. </b><br><br>' \
                            'This is found by examining the hyperlink in your browser:<br><br>' \
                            '<i>https://docs.google.com/spreadsheets/d/<b>****spreadshetID*****</b>/edit#gid=0 </i>' \
                            '<br><br>Copy just the long character string in place of ****spreadshetID*****.<br><br>'

        instructions_input = command_inputs.addTextBoxCommandInput('instructions', '',
                                                                   instructions_text, 16, True)
        instructions_input.isVisible = False

        existing_sheet_id_input = command_inputs.addStringValueInput('existing_sheet_id', 'Existing Sheet ID:', '')

        existing_sheet_id_input.isVisible = False
