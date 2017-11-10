from xml.etree import ElementTree
from xml.etree.ElementTree import SubElement

from os.path import expanduser
import os


# Get default directory
def get_default_dir():
    # Get user's home directory
    default_dir = os.path.expanduser("~")

    # Create a subdirectory for this application
    default_dir = os.path.join(default_dir, 'FusionBolter', '')

    # Create the folder if it does not exist
    if not os.path.exists(default_dir):
        os.makedirs(default_dir)

    return default_dir


# Get default directory
def get_default_model_dir():

    default_dir = get_default_dir()

    # Create a subdirectory for the models
    default_dir = os.path.join(default_dir, 'Models', '')

    # Create the folder if it does not exist
    if not os.path.exists(default_dir):
        os.makedirs(default_dir)

    return default_dir


# Creates directory and returns file name for settings file
def get_settings_file():
    default_dir = get_default_dir()
    # Create file name in this path
    settings_file_name = os.path.join(default_dir, 'settings.xml')
    return settings_file_name


# Writes user settings to a file in local home directory
def write_settings(spreadsheet_id):

    xml_file_name = get_settings_file()

    # If file doesn't exist create it
    if not os.path.isfile(xml_file_name):
        new_file = open(xml_file_name, 'w')
        new_file.write('<?xml version="1.0"?>')
        new_file.write("<FusionBolter /> ")
        new_file.close()
        tree = ElementTree.parse(xml_file_name)
        root = tree.getroot()

    # Otherwise delete existing settings
    else:
        tree = ElementTree.parse(xml_file_name)
        root = tree.getroot()
        root.remove(root.find('settings'))

    # Write settings
    settings = SubElement(root, 'settings')
    SubElement(settings, 'spreadsheet_id', value=spreadsheet_id)
    tree.write(xml_file_name)


# Read user settings in from XML file
def read_settings():

    xml_file_name = get_settings_file()

    if os.path.exists(xml_file_name):

        # Get the root of the XML tree
        tree = ElementTree.parse(xml_file_name)
        root = tree.getroot()

        # Get the settings values
        spreadsheet_id = root.find('settings/spreadsheet_id').attrib['value']
        return spreadsheet_id

    else:
        raise Exception('You need to configure Bolter before running it!')
        return''



