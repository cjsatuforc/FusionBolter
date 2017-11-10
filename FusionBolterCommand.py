import adsk.core
import adsk.fusion
import traceback

import math
import os
import csv
from collections import defaultdict, namedtuple

import json
from urllib import request, error

from .Fusion360Utilities.Fusion360Utilities import get_app_objects
from .Fusion360Utilities.Fusion360CommandBase import Fusion360CommandBase
from .Fusion360Utilities.Fusion360DebugUtilities import perf_log, perf_message

from .SheetsService import sheets_get_range, sheets_get_ranges

from .BolterUtilities import read_settings

Hole_Edge = namedtuple('Hole_Edge', ('edge', 'translation_matrix', 'radius', 'cylinderFace', 'planeFace'))
Hole_Pair = namedtuple('Hole_Pair', ('top_hole', 'bottom_hole', 'length'))
Hardware_Item = namedtuple('Hardware_Item', ('type', 'description', 'model'))
joint_check = []

MASTER_FILE = 'Hardware_Sizes - Master.csv'

ONLINE = True


# Get default directory
def get_default_model_dir():
    # Get user's home directory
    default_dir = os.path.expanduser("~")

    # Create a subdirectory for this application
    default_dir = os.path.join(default_dir, 'FusionBolter', '')

    # Create the folder if it does not exist
    if not os.path.exists(default_dir):
        os.makedirs(default_dir)

    # Create a subdirectory for this application
    default_dir = os.path.join(default_dir, 'Models', '')

    # Create the folder if it does not exist
    if not os.path.exists(default_dir):
        os.makedirs(default_dir)

    return default_dir


# Function to convert a csv file to a list of dictionaries.  Takes in one variable called "data_file_name"
def csv_dict_list(data_file_name):
    # + '.csv'
    csv_file = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'data', data_file_name)

    # Open variable-based csv, iterate over the rows and map values to a list of dictionaries containing key/value pairs
    reader = csv.DictReader(open(csv_file, 'r'))
    dict_list = []
    for line in reader:
        dict_list.append(line)
    return dict_list


def sheets_to_dict(spreadsheet_id, sheet_number):
    url = 'http://gsx2json.com/api?id=%s&sheet=%s&columns=false&integers=false' % (spreadsheet_id, sheet_number)

    try:

        # Open URL
        response = request.urlopen(url)

        # Get response and process json
        data = response.read()
        encoding = response.info().get_content_charset('utf-8')
        out = json.loads(data.decode(encoding))

        # ao=get_app_objects()
        # ao['ui'].messageBox(str(out['rows']))
        # from pprint import pprint
        # pprint(out['rows'])

        # Return all rows
        return out['rows']

    except error.HTTPError as e:
        print('HTTPError = ' + str(e.code))
    except error.URLError as e:
        print('URLError = ' + str(e.reason))
    # except error.HTTPException as e:
    #     print('HTTPException')
    except Exception:
        import traceback
        print('generic exception: ' + traceback.format_exc())

    return None


def get_parts_from_sheets2(spreadsheet_id):

    # Todo from config file
    range_name = 'Master'

    master_list = sheet_range_to_dict2(spreadsheet_id, range_name)

    if not master_list:
        raise Exception('failed to get data from Google Sheet')

    hardware_list = defaultdict(list)

    for item in master_list:
        hardware_type = item.get('type', None)
        description = item.get('description', None)
        model = item.get('model', None)

        hardware_item = Hardware_Item(hardware_type, description, model)

        hardware_list[hardware_item.type].append(hardware_item)

    return hardware_list


# Returns list of Dict's for each row after header row
def sheet_range_to_dict2(spreadsheet_id, range_name):

    result = sheets_get_range(spreadsheet_id, range_name)

    rows = result.get('values', [])

    dict_list = []

    for row in rows[1:]:
        row_dict = dict(zip(rows[0], row))
        dict_list.append(row_dict)

    return dict_list


# Returns list of Dict's for each row after header row
def sheet_ranges_to_dict2(value_ranges, range_index):

    rows = value_ranges[range_index].get('values', [])

    dict_list = []

    for row in rows[1:]:
        row_dict = dict(zip(rows[0], row))
        dict_list.append(row_dict)

    return dict_list


def find_hole_edges(target_faces):
    # Initialize a list that's used to return information about the found "hole" edges.
    hole_positions = []

    for target_face in target_faces:
        face = adsk.fusion.BRepFace.cast(target_face)

        # Iterate over all the edges looking for "hole" edges.
        for edge in face.edges:
            if not edge.isDegenerate:
                # Check to see if the edge is a circle.
                if edge.geometry.curveType == adsk.core.Curve3DTypes.Circle3DCurveType:
                    # Get the two faces that are connected by the edge.
                    cylinderFace = None
                    planeFace = None
                    for face in edge.faces:
                        if face.geometry.surfaceType == adsk.core.SurfaceTypes.CylinderSurfaceType:
                            cylinderFace = face
                        elif face.geometry.surfaceType == adsk.core.SurfaceTypes.PlaneSurfaceType:
                            planeFace = face

                    if cylinderFace and planeFace:
                        # Check to see if the circular edge is an inner loop of the planar face
                        # and that the loop consists of a single curve, which is the circular edge.
                        for planeLoop in planeFace.loops:
                            if not planeLoop.isOuter and planeLoop.edges.count == 1 and planeLoop.edges[0] == edge:
                                # Get an arbitrary point on the cylinder.
                                pnt = cylinderFace.pointOnFace

                                # Get a normal from the cylinder.
                                (result_success, normal) = cylinderFace.evaluator.getNormalAtPoint(pnt)

                                # Change the length of the normal to be the radius of the cylinder.
                                normal.normalize()
                                cyl = cylinderFace.geometry
                                normal.scaleBy(cyl.radius)

                                # Translate the point on the cylinder along the normal vector.
                                pnt.translateBy(normal)

                                # Check to see if the point lies along the cylinder axis.
                                # if it does, then this is a hole and not a boss.
                                if distance_point_to_line(pnt, cyl.origin, cyl.axis) < cyl.radius:
                                    # Create a matrix that will define the position of the cork.
                                    (result_success, zDir) = planeFace.evaluator.getNormalAtPoint(planeFace.pointOnFace)
                                    xDir = normal
                                    yDir = zDir.crossProduct(xDir)
                                    xDir.normalize()
                                    yDir.normalize()
                                    zDir.normalize()
                                    translation_matrix = adsk.core.Matrix3D.create()
                                    translation_matrix.setWithCoordinateSystem(edge.geometry.center, xDir, yDir, zDir)

                                    # Save the edge, matrix, and radius.
                                    hole_positions.append(Hole_Edge(edge, translation_matrix, edge.geometry.radius,
                                                                    cylinderFace, planeFace))

                                break

    return hole_positions


# Computes the distance between the given infinite line and a point.
def distance_point_to_line(point, lineRootPoint, lineDirection):
    try:
        # Compute the distance between the point and the origin point of the line.
        dist = lineRootPoint.distanceTo(point)

        # Special case when the point is the same as the line root point.
        if dist < 0.000001:
            return 0

        # Create a vector that defines the direction from the lines origin to the input point.
        pointVec = lineRootPoint.vectorTo(point)

        # Get the angle between the two vectors.
        angle = lineDirection.angleTo(pointVec)

        # Calculate the side height of the triangle.
        return dist * math.sin(angle)
    except:
        app = adsk.core.Application.get()
        ui = app.userInterface
        ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))


def group_length_results(hole_pairs):
    ao = get_app_objects()
    units_manager = ao['units_manager']

    hole_map = defaultdict(list)

    for hole in hole_pairs:
        diameter_value = round(units_manager.convert(hole.top_hole.radius * 2, 'cm',
                                                     units_manager.defaultLengthUnits), 5)
        length_value = ao['units_manager'].formatInternalValue(hole.length, "DefaultDistance", False)

        key_string = str(diameter_value) + ' X ' + str(length_value)
        hole_map[key_string].append(hole)

    return hole_map


def find_pairs(top_holes, bottom_holes):
    hole_pairs = []

    for top_hole in top_holes:

        for bottom_hole in bottom_holes:
            distance = distance_point_to_line(top_hole.edge.geometry.center,
                                              bottom_hole.cylinderFace.geometry.origin,
                                              bottom_hole.cylinderFace.geometry.axis)
            if distance < .0001:
                length = top_hole.edge.geometry.center.distanceTo(bottom_hole.edge.geometry.center)
                hole_pairs.append(Hole_Pair(top_hole, bottom_hole, length))
                break

    return hole_pairs


# Opens the correct model
def open_model(model_name: str, target_comp: adsk.fusion.Component) -> adsk.fusion.Occurrence:
    ao = get_app_objects()

    # Get import manager
    import_manager = ao['import_manager']

    # import_file_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'models', model_name)
    import_file_name = os.path.join(get_default_model_dir(), model_name)

    archive_options = import_manager.createFusionArchiveImportOptions(import_file_name)

    # Import f3d file to component
    import_manager.importToTarget(archive_options, target_comp)

    # Return last occurrence in target component.  This is the new one
    return target_comp.occurrences.asList.item(target_comp.occurrences.count - 1)


def get_size_from_sheet(radius, spreadsheet_id, sheet_number):

    # Get a reference to all relevant application objects in a dictionary
    app_objects = get_app_objects()
    um = app_objects['units_manager']

    all_parts = sheet_range_to_dict2(spreadsheet_id, sheet_number)

    comp_def = None

    if all_parts is None:
        raise Exception('Failed to load part size')

    for part in all_parts:

        # TODO broken for radius calc
        if um.evaluateExpression(part['min'], 'in') <= radius * 2 < um.evaluateExpression(part['max'], 'in'):
            comp_def = part
            return comp_def
        last_number = part['number']
        last_size = um.evaluateExpression(part['min'], 'in')

    app_objects['ui'].messageBox('Could not find size radius =   ' + str(last_size))

    return comp_def


# Updates length for bolt component
def update_length(comp_def, length):
    # Get a reference to all relevant application objects in a dictionary
    app_objects = get_app_objects()
    um = app_objects['units_manager']

    comp_def['length'] = str(um.convert(length, "internalUnits", 'in'))


# Places a new Component
def place_comp2(target_comp, model_name, comp_def, comp_name):
    # log = []
    # perf_log(log, 'place_comp2', 'Start', comp_name)
    occ = open_model(model_name, target_comp)
    # perf_log(log, 'place_comp2', 'open_model', comp_name)

    occ.component.name = comp_def.get('name', comp_name)
    occ.component.description = comp_def.get('description', comp_name)
    occ.component.partNumber = comp_def.get('number', '')

    update_parameters(occ, comp_def)
    # perf_log(log, 'place_comp2', 'update_parameters', comp_name)

    # perf_message(log)

    return occ


def update_parameters(occ, comp_size):
    app_objects = get_app_objects()
    um = app_objects['units_manager']

    parameters = occ.component.modelParameters

    # ao = get_app_objects()
    # ao['ui'].messageBox(str(comp_size))
    # for parameter in all_parameters:
    #
    #     new_value = size.get(parameter.name)
    #     if new_value is not None:
    #         unit_type = parameter.unit
    #
    #         if len(unit_type) > 0:
    #
    #             if um.isValidExpression(new_value, unit_type):
    #                 sheet_value = um.evaluateExpression(new_value, unit_type)
    #             else:
    #                 continue
    #         else:
    #             sheet_value = float(new_value)
    #
    #         # TODO handle units with an attribute that is written on create.  Can be set for link
    #         if parameter.value != sheet_value:
    #             parameter.value = sheet_value

    for parameter in parameters:
        attribute = parameter.attributes.itemByName('FusionBolter', 'FusionBolterParameter')
        if attribute is not None:

            # Todo Units are messed up
            new_value = comp_size.get(attribute.value)
            if new_value is not None:
                # ui.messageBox(str(new_value))
                # ui.messageBox(str(type(new_value)))
                # ui.messageBox(new_value)
                parameter.value = um.evaluateExpression(new_value, 'in')


def place_existing_component(target_comp, translation_matrix, comp):
    occ = target_comp.occurrences.addExistingComponent(comp, translation_matrix)
    return occ


# Check if component already exists
def check_number(comp_number, design):
    # Look for an existing component by looking for a pre-defined name.
    for check_comp in design.allComponents:
        if check_comp.partNumber == comp_number:
            return check_comp

    return None


def get_tagged_edge(occ: adsk.fusion.Occurrence, tag_name: str) -> adsk.fusion.BRepEdge:
    body = occ.component.bRepBodies.item(0)
    joint_edge = None

    for edge in body.edges:

        attribute = edge.attributes.itemByName('FusionBolter', tag_name)

        if attribute is not None:
            joint_edge = edge
            break

    joint_edge = joint_edge.createForAssemblyContext(occ)

    return joint_edge


def get_tagged_face(occ: adsk.fusion.Occurrence, tag_name: str) -> adsk.fusion.JointGeometry:
    body = occ.component.bRepBodies.item(0)
    target_face = None

    for face in body.faces:

        attribute = face.attributes.itemByName('FusionBolter', tag_name)

        if attribute is not None:
            target_face = face
            target_face = target_face.createForAssemblyContext(occ)
            joint_geometry = adsk.fusion.JointGeometry.createByPlanarFace(target_face, None,
                                                                          adsk.fusion.JointKeyPointTypes.CenterKeyPoint)
            return joint_geometry


    return None


def joint_to_hole(target_comp: adsk.fusion.Component, joint_face, hole):
    design = target_comp.parentDesign

    (top_target_edge, top_translation_matrix, top_radius, top_cylinder_face, top_hole_face) = hole

    cyl = top_cylinder_face.geometry
    test_dist = cyl.origin.distanceTo(top_target_edge.geometry.center)

    if test_dist < .0001:
        key_point = adsk.fusion.JointKeyPointTypes.EndKeyPoint
        # ui.messageBox('Top = End Point')

    else:
        key_point = adsk.fusion.JointKeyPointTypes.StartKeyPoint
        # ui.messageBox('Top = Start Point')

    top_hole_joint_geo = adsk.fusion.JointGeometry.createByNonPlanarFace(top_cylinder_face, key_point)

    joint_input = target_comp.joints.createInput(joint_face, top_hole_joint_geo)

    top_joint = target_comp.joints.add(joint_input)

    top_hole_face = adsk.fusion.BRepFace.cast(top_hole_face)

    test, top_hole_face_normal = top_hole_face.evaluator.getNormalAtPoint(top_hole_face.pointOnFace)

    top_washer_face = joint_face.entityOne
    test, top_washer_normal = top_washer_face.evaluator.getNormalAtPoint(top_washer_face.pointOnFace)

    face_angle = top_hole_face_normal.angleTo(top_washer_normal)

    if face_angle < 3:
        # Get the current position of the timeline.
        start_position = design.timeline.markerPosition

        # Roll back the time line to the joint.
        top_joint.timelineObject.rollTo(True)

        top_joint.isFlipped = not top_joint.isFlipped

        # Move the marker back
        design.timeline.markerPosition = start_position


def place_components3(spreadsheet_id, target_comp: adsk.fusion.Component, hole_map, bolt_extension, bolt, washer, nut):
    global joint_check

    # Get a reference to all relevant application objects in a dictionary
    app_objects = get_app_objects()
    ui = app_objects['ui']
    # Get a reference to all relevant application objects in a dictionary
    um = app_objects['units_manager']

    design = target_comp.parentDesign
    root_comp = design.rootComponent

    bolt_source = bolt.model
    washer_source = washer.model
    nut_source = nut.model

    log = []

    perf_log(log, 'place components', 'start', '')

    # TODO want to do a perf log on this
    for size, hole_pairs in hole_map.items():

        # Get size data from first hole.  Not great.
        radius = hole_pairs[0].top_hole.radius
        length = hole_pairs[0].length

        # Get components
        # TODO check for existing bolt
        bolt_def = get_size_from_sheet(radius, spreadsheet_id, bolt.description)
        perf_log(log, 'place components', 'get_size', 'bolt')

        washer_def = get_size_from_sheet(radius, spreadsheet_id, washer.description)
        washer_comp_top = check_number(washer_def['number'], design)
        perf_log(log, 'place components', 'get_size', 'washer1')

        # TODO Creating second instance no matter what because top isn't created yet
        # TODO Allow for different top and bottom washer
        # TODO Allow for multiple washers in top and bottom stack
        washer_comp_bottom = check_number(washer_def['number'], design)
        perf_log(log, 'place components', 'get_size', 'washer2')

        nut_def = get_size_from_sheet(
            radius, spreadsheet_id, nut.description)
        nut_comp = check_number(nut_def['number'], design)
        perf_log(log, 'place components', 'get_size', 'nut')

        length += um.evaluateExpression(washer_def['thickness'], 'in')
        length += um.evaluateExpression(washer_def['thickness'], 'in')
        length += um.evaluateExpression(nut_def['thickness'], 'in')
        length += bolt_extension

        # TODO maybe if there is no number use name?
        # TODO won't instance right.  Need to check number AND length
        bolt_comp = check_number(bolt_def['number'] + str(length), design)
        update_length(bolt_def, length)
        perf_log(log, 'place components', 'check_number', 'bolt')
        perf_count = 0

        for hole_pair in hole_pairs:  # type: Hole_Pair

            (top_target_edge, top_translation_matrix, top_radius, top_cylinder_face,
             top_hole_face) = hole_pair.top_hole
            (bottom_target_edge, bottom_translation_matrix, bottom_radius, bottom_cylinder_face,
             bottom_hole_face) = hole_pair.bottom_hole

            if bolt_comp is None:
                # todo is this necessary?
                bolt_def['length'] = str(um.convert(length, "internalUnits", 'in'))
                bolt_occ = place_comp2(target_comp, bolt_source, bolt_def, 'Bolt_' + size)
                bolt_comp = adsk.fusion.Component.cast(bolt_occ.component)
                perf_log(log, 'place components', 'place_comp2', 'Bolt')
            else:
                bolt_occ = place_existing_component(target_comp, top_translation_matrix, bolt_comp)
                perf_log(log, 'place components', 'place_existing', 'Bolt' + str(perf_count))

            if washer_comp_top is None:
                washer_occ_top = place_comp2(target_comp, washer_source, washer_def, 'Washer_' + size)
                washer_comp_top = adsk.fusion.Component.cast(washer_occ_top.component)
                perf_log(log, 'place components', 'place_comp2', 'Washer1')
            else:
                washer_occ_top = place_existing_component(target_comp, top_translation_matrix, washer_comp_top)
                perf_log(log, 'place components', 'place_existing', 'Washer1' + str(perf_count))

            if washer_comp_bottom is None:
                washer_occ_bottom = place_comp2(target_comp, washer_source, washer_def, 'Washer_' + size)
                washer_comp_bottom = adsk.fusion.Component.cast(washer_occ_bottom.component)
                perf_log(log, 'place components', 'place_comp2', 'Washer2')

            else:
                washer_occ_bottom = place_existing_component(target_comp, bottom_translation_matrix, washer_comp_bottom)
                perf_log(log, 'place components', 'place_existing', 'Washer2' + str(perf_count))

            if nut_comp is None:
                nut_occ = place_comp2(target_comp, nut_source, nut_def, 'Nut_' + size)
                nut_comp = adsk.fusion.Component.cast(nut_occ.component)
                perf_log(log, 'place components', 'place_comp2', 'Nut')
            else:
                nut_occ = place_existing_component(target_comp, top_translation_matrix, nut_comp)
                perf_log(log, 'place components', 'place_existing', 'Nut' + str(perf_count))

            # Get Joint Geometry for bolt
            bolt_joint_face = get_tagged_face(bolt_occ, 'BoltFace')

            # Get Joint Geometry for Washer
            top_washer_top_face = get_tagged_face(washer_occ_top, 'WasherTopFace')
            top_washer_bottom_face = get_tagged_face(washer_occ_top, 'WasherBottomFace')
            perf_log(log, 'place components', 'get_tagged_faces', 'Top')

            # Create Joint for Top hole to top washer
            joint_to_hole(target_comp, top_washer_bottom_face, hole_pair.top_hole)
            perf_log(log, 'place components', 'Joint', 'Top_hole' + str(perf_count))

            # Create joint Bolt to top washer
            joint_input = target_comp.joints.createInput(bolt_joint_face, top_washer_top_face)
            joint_input.isFlipped = True
            root_comp.joints.add(joint_input)
            perf_log(log, 'place components', 'Joint', 'Bolt' + str(perf_count))

            # Bottom Joints
            bottom_washer_top_face = get_tagged_face(washer_occ_bottom, 'WasherTopFace')
            bottom_washer_bottom_face = get_tagged_face(washer_occ_bottom, 'WasherBottomFace')
            nut_joint_face = get_tagged_face(nut_occ, 'NutBottomFace')
            perf_log(log, 'place components', 'get_tagged_faces', 'Bottom')

            # Create Joint for bottom hole to bottom washer
            joint_to_hole(target_comp, bottom_washer_bottom_face, hole_pair.bottom_hole)
            perf_log(log, 'place components', 'Joint', 'Bottom_hole' + str(perf_count))

            # Create joint input
            joint_input = target_comp.joints.createInput(nut_joint_face, bottom_washer_top_face)
            joint_input.isFlipped = True
            root_comp.joints.add(joint_input)
            perf_log(log, 'place components', 'Joint', 'Nut' + str(perf_count))

            perf_count += 1

    # perf_message(log)


class FusionBolterCommand(Fusion360CommandBase):

    def __init__(self, cmd_def, debug):
        super().__init__(cmd_def, debug)
        self.hardware_list = {}

    # Run whenever a user makes any change to a value or selection in the addin UI
    # Commands in here will be run through the Fusion processor and changes will be reflected in  Fusion graphics area
    def on_preview(self, command, inputs, args, input_values):
        pass

    # Run after the command is finished.
    # Can be used to launch another command automatically or do other clean up.
    def on_destroy(self, command, inputs, reason, input_values):
        pass

    # Run when any input is changed.
    # Can be used to check a value and then update the add-in UI accordingly
    def on_input_changed(self, command_, command_inputs, changed_input, input_values):
        pass

    # Run when the user presses OK
    # This is typically where your main program logic would go
    def on_execute(self, command, inputs, args, input_values):
        # Get a reference to all relevant application objects in a dictionary
        app_objects = get_app_objects()
        ui = app_objects['ui']

        root_comp = app_objects['root_comp']

        # Get the hole edges in the body.
        top_holes = find_hole_edges(input_values['top_hole_faces'])

        bottom_holes = find_hole_edges(input_values['bottom_hole_faces'])

        spreadsheet_id = read_settings()

        for item in self.hardware_list['Bolt']:
            if input_values['Bolt'] == item.description:
                bolt_hardware = item

        for item in self.hardware_list['Washer']:
            if input_values['Washer'] == item.description:
                washer_hardware = item

        for item in self.hardware_list['Nut']:
            if input_values['Nut'] == item.description:
                nut_hardware = item

        # TODO left off here next pass hardware to place components to select proper bolt

        hole_pairs = find_pairs(top_holes, bottom_holes)

        # hole_map = group_results(top_holes)

        # message_sring = ''
        # for diameter, hole in hole_map.items():
        #     message_sring += 'Diameter:  ' + diameter
        #     message_sring += '    \t'
        #     message_sring += 'Quantity:  ' + str(len(hole))
        #     message_sring += '\n'
        #
        # ui.messageBox(message_sring)

        hole_map = group_length_results(hole_pairs)
        message_sring = ''
        for key, holes in hole_map.items():
            message_sring += key
            message_sring += '    \t'
            message_sring += 'Quantity:  ' + str(len(holes))
            message_sring += '\n'

        place_components3(spreadsheet_id, root_comp, hole_map, input_values['bolt_extension'], bolt_hardware, washer_hardware,
                          nut_hardware)

        ui.messageBox(message_sring)

    # Run when the user selects your command icon from the Fusion 360 UI
    # Typically used to create and display a command dialog box
    # The following is a basic sample of a dialog UI
    def on_create(self, command, command_inputs):

        # TODO pull sheets once, put in self.
        # Gets necessary application objects
        app_objects = get_app_objects()
        default_units = app_objects['units_manager'].defaultLengthUnits

        top_faces_input = command_inputs.addSelectionInput('top_hole_faces', 'Select Top faces to find holes',
                                                           'Hole top faces')
        top_faces_input.addSelectionFilter('SolidFaces')
        top_faces_input.setSelectionLimits(1, 0)

        bot_faces_input = command_inputs.addSelectionInput('bottom_hole_faces', 'Select Bottom faces to find holes',
                                                           'Hole top faces')
        bot_faces_input.addSelectionFilter('SolidFaces')
        bot_faces_input.setSelectionLimits(1, 0)

        command_inputs.addValueInput('bolt_extension', 'Bolt Extension', default_units,
                                     adsk.core.ValueInput.createByString('.25 in'))

        spreadsheet_id = read_settings()

        self.hardware_list = get_parts_from_sheets2(spreadsheet_id)

        bolt_drop_down = command_inputs.addDropDownCommandInput('Bolt', 'Bolt Family',
                                                                adsk.core.DropDownStyles.LabeledIconDropDownStyle)
        for item in self.hardware_list['Bolt']:
            bolt_drop_down.listItems.add(item.description, False)

        washer_drop_down = command_inputs.addDropDownCommandInput('Washer', 'Washer Family',
                                                                  adsk.core.DropDownStyles.LabeledIconDropDownStyle)
        for item in self.hardware_list['Washer']:
            washer_drop_down.listItems.add(item.description, False)

        nut_drop_down = command_inputs.addDropDownCommandInput('Nut', 'Nut Family',
                                                               adsk.core.DropDownStyles.LabeledIconDropDownStyle)
        for item in self.hardware_list['Nut']:
            nut_drop_down.listItems.add(item.description, False)


class FusionBolterFindCommand(Fusion360CommandBase):

    def on_execute(self, command, inputs, args, input_values):
        # Get a reference to all relevant application objects in a dictionary
        app_objects = get_app_objects()
        ui = app_objects['ui']


        # Get the hole edges in the body.
        top_holes = find_hole_edges(input_values['top_hole_faces'])

        bottom_holes = find_hole_edges(input_values['bottom_hole_faces'])

        hole_pairs = find_pairs(top_holes, bottom_holes)

        hole_map = group_length_results(hole_pairs)
        message_sring = ''
        for key, holes in hole_map.items():
            message_sring += key
            message_sring += '    \t'
            message_sring += 'Quantity:  ' + str(len(holes))
            message_sring += '\n'

        ui.messageBox(message_sring)

    def on_create(self, command, command_inputs):

        # TODO pull sheets once, put in self.
        # Gets necessary application objects
        app_objects = get_app_objects()
        default_units = app_objects['units_manager'].defaultLengthUnits

        top_faces_input = command_inputs.addSelectionInput('top_hole_faces', 'Select Top faces to find holes',
                                                           'Hole top faces')
        top_faces_input.addSelectionFilter('SolidFaces')
        top_faces_input.setSelectionLimits(1, 0)

        bot_faces_input = command_inputs.addSelectionInput('bottom_hole_faces', 'Select Bottom faces to find holes',
                                                           'Hole top faces')
        bot_faces_input.addSelectionFilter('SolidFaces')
        bot_faces_input.setSelectionLimits(1, 0)


