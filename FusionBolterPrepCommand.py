import adsk.core
import adsk.fusion
import traceback

import os

from .Fusion360Utilities.Fusion360Utilities import get_app_objects
from .Fusion360Utilities.Fusion360CommandBase import Fusion360CommandBase

from .SheetsService import sheets_append_values, sheets_update_values, sheets_update_spreadsheet

from .BolterUtilities import get_default_model_dir, read_settings


def tag_bolter_parameters():
    # Gets necessary application objects from utility function
    app_objects = get_app_objects()

    component = app_objects['root_comp']

    # Set all parameter values based on the input form
    parameters = component.modelParameters

    for parameter in parameters:
        parameter.attributes.add('FusionBolter', 'FusionBolterParameter', parameter.name)


def dup_check(name):
    if os.path.exists(name):
        base, ext = os.path.splitext(name)
        base += '-dup'
        name = base + ext
        dup_check(name)
    return name


def export_active_doc(folder, doc_name):
    app = adsk.core.Application.get()
    design = app.activeProduct
    export_mgr = design.exportManager

    export_name = folder + doc_name + '.f3d'
    export_name = dup_check(export_name)
    export_options = export_mgr.createFusionArchiveExportOptions(export_name)
    export_mgr.execute(export_options)


# Create Parameters Sheet
def create_sheet_parameters(sheet_id, model_name):
    app_objects = get_app_objects()
    um = app_objects['units_manager']
    design = app_objects['design']

    parameters = design.allParameters

    headers = []
    dims = []

    headers.append('number')
    dims.append(design.rootComponent.partNumber)

    headers.append('description')
    dims.append(design.rootComponent.description)

    for parameter in parameters:
        headers.append(parameter.name)
        dims.append(um.formatInternalValue(parameter.value, parameter.unit, False))

    headers.append('min')
    dims.append(0)

    headers.append('max')
    dims.append(0)

    range_body = {"range": model_name,
                  "values": [headers, dims]}

    sheet_range = model_name

    sheets_update_values(sheet_id, sheet_range, range_body)


class FusionBolterPrepCommand(Fusion360CommandBase):
    def on_execute(self, command, inputs, args, input_values):

        app = adsk.core.Application.get()
        folder = get_default_model_dir()
        doc_name = app.activeDocument.name
        doc_name = doc_name[:doc_name.rfind(' v')]

        spreadsheet_id = read_settings()

        if input_values['model_type'] == 'Bolt':
            item_attributes = input_values['bolt_selection'][0].attributes
            item_attributes.add('FusionBolter', 'BoltFace', '1')

        elif input_values['model_type'] == 'Washer':
            item_attributes = input_values['washer_top_selection'][0].attributes
            item_attributes.add('FusionBolter', 'WasherTopFace', '1')
            item_attributes = input_values['washer_bottom_selection'][0].attributes
            item_attributes.add('FusionBolter', 'WasherBottomFace', '1')

        elif input_values['model_type'] == 'Nut':
            item_attributes = input_values['nut_selection'][0].attributes
            item_attributes.add('FusionBolter', 'NutBottomFace', '1')

        tag_bolter_parameters()

        export_active_doc(folder, doc_name)

        range_body = {'range': 'Master',
                      'values': [[input_values['model_type'],
                                  input_values['description'],
                                  doc_name + '.f3d']]
                      }

        sheets_append_values(spreadsheet_id, 'Master', range_body)

        update_request_list = [{
            "addSheet": {
                "properties": {
                    "title": input_values['description']
                }
            }
        }]

        sheets_update_spreadsheet(spreadsheet_id, update_request_list)

        create_sheet_parameters(spreadsheet_id, input_values['description'])

    def on_create(self, command, command_inputs):

        command_inputs.addTextBoxCommandInput('type_title', '', '<b>Model Type</b>', 1, True)
        model_type_group = command_inputs.addRadioButtonGroupCommandInput('model_type')
        model_type_group.listItems.add('Bolt', True)
        model_type_group.listItems.add('Washer', False)
        model_type_group.listItems.add('Nut', False)

        command_inputs.addStringValueInput('description', 'Description', 'My New Model')

        bolt_selection = command_inputs.addSelectionInput('bolt_selection', 'Bolt Mating Face:',
                                                          'Select the under side of the Bolt head')
        bolt_selection.addSelectionFilter('SolidFaces')
        bolt_selection.isVisible = True
        bolt_selection.setSelectionLimits(0, 1)

        washer_top_selection = command_inputs.addSelectionInput('washer_top_selection', 'Washer Top Face:',
                                                                'Select the top face of the washer.')
        washer_top_selection.isVisible = False
        washer_top_selection.isEnabled = False
        washer_top_selection.setSelectionLimits(0, 1)

        washer_bottom_selection = command_inputs.addSelectionInput('washer_bottom_selection', 'Washer Bottom Face:',
                                                                   'Select the bottom face of the washer.')
        washer_bottom_selection.isVisible = False
        washer_bottom_selection.isEnabled = False
        washer_bottom_selection.setSelectionLimits(0, 1)

        nut_selection = command_inputs.addSelectionInput('nut_selection', 'Nut Mating Face:',
                                                         'Select the under side of the Nut')
        nut_selection.isVisible = False
        nut_selection.isEnabled = False
        nut_selection.setSelectionLimits(0, 1)

    def on_input_changed(self, command_, command_inputs, changed_input, input_values):

        if changed_input.id == 'model_type':
            if changed_input.selectedItem.name == 'Bolt':

                command_inputs.itemById('bolt_selection').isVisible = True
                command_inputs.itemById('washer_top_selection').isVisible = False
                command_inputs.itemById('washer_bottom_selection').isVisible = False
                command_inputs.itemById('nut_selection').isVisible = False

            elif changed_input.selectedItem.name == 'Washer':

                command_inputs.itemById('bolt_selection').isVisible = False
                command_inputs.itemById('washer_top_selection').isVisible = True
                command_inputs.itemById('washer_bottom_selection').isVisible = True
                command_inputs.itemById('nut_selection').isVisible = False

            elif changed_input.selectedItem.name == 'Nut':

                command_inputs.itemById('bolt_selection').isVisible = False
                command_inputs.itemById('washer_top_selection').isVisible = False
                command_inputs.itemById('washer_bottom_selection').isVisible = False
                command_inputs.itemById('nut_selection').isVisible = True
