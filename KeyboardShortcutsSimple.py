#Author-Thomas Axelsson
#Description-Lists all keyboard shortcuts

# This file is part of KeyboardShortcutsSimple, a Fusion 360 add-in for naming
# features directly after creation.
#
# Copyright (C) 2020  Thomas Axelsson
#
# KeyboardShortcutsSimple is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# KeyboardShortcutsSimple is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with KeyboardShortcutsSimple.  If not, see <https://www.gnu.org/licenses/>.

import adsk.core, adsk.fusion, adsk.cam, traceback

from collections import defaultdict
import ctypes
import json
import os
import pathlib
import sys
import xml.etree.ElementTree as ET

NAME = 'Keyboard Shortcuts Simple'

FILE_DIR = os.path.dirname(os.path.realpath(__file__))

if FILE_DIR not in sys.path:
    sys.path.append(FILE_DIR)
import thomasa88lib, thomasa88lib.events, thomasa88lib.timeline

# Force modules to be fresh during development
import importlib
importlib.reload(thomasa88lib)
importlib.reload(thomasa88lib.events)
importlib.reload(thomasa88lib.timeline)

LIST_CMD_ID = 'thomasa88_keyboardShortcutsSimpleList'

app_ = None
ui_ = None
events_manager_ = thomasa88lib.events.EventsManger(NAME)
list_cmd_def_ = None
cmd_def_workspaces_map_ = None

class Hotkey:
    pass

def list_command_created_handler(args):
    eventArgs = adsk.core.CommandCreatedEventArgs.cast(args)

    # The nifty thing with cast is that code completion then knows the object type
    cmd = adsk.core.Command.cast(args.command)
    cmd.isRepeatable = False
    cmd.isExecutedWhenPreEmpted = False
    cmd.isOKButtonVisible = False
    cmd.setDialogMinimumSize(500, 500)

    inputs = cmd.commandInputs
    inputs.addTextBoxCommandInput('list', '', get_hotkeys_str(), 30, False)

def get_hotkeys_str():
    options_files = find_options_files()
    # TODO: Pick the correct user/profile if there are multiple options files
    hotkeys = parse_hotkeys(options_files[0])
    string = '<table>'
    build_cmd_def_workspaces_map()
    hotkeys = map_command_names(hotkeys)
    ns_hotkeys = namespace_group_hotkeys(hotkeys)
    #gör dedup i workspaces!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!1
    for workspace_id, hotkeys in ns_hotkeys.items():
        if workspace_id:
            workspace_name = ui_.workspaces.itemById(workspace_id).name
        else:
            workspace_name = 'General'
        string += f'<tr><td><b>{workspace_name}<b></td><td></td></tr>'
        for hotkey in sorted(hotkeys, key=lambda  h: h.command_name):
            string += f'<tr><td>{hotkey.command_name}</td><td> {hotkey.key_sequence}</td></tr>'
    string += '</table>'
    return string

def map_command_names(hotkeys):
    for hotkey in hotkeys:
        command = ui_.commandDefinitions.itemById(hotkey.command_id)
        if command:
            command_name = command.name
        else:
            command_name = hotkey.command_id
        if hotkey.command_argument:
            command_name += '-&gt;' + hotkey.command_argument
        hotkey.command_name = command_name
    return hotkeys

def namespace_group_hotkeys(hotkeys):
    ns_hotkeys = defaultdict(list)
    for hotkey in hotkeys:
        workspaces = find_cmd_workspaces(hotkey.command_id)
        for workspace in workspaces:
            ns_hotkeys[workspace].append(hotkey)
        else:
            ns_hotkeys[None].append(hotkey)
    return ns_hotkeys

def build_cmd_def_workspaces_map():
    global cmd_def_workspaces_map_
    cmd_def_workspaces_map_ = defaultdict(set)
    for workspace in ui_.workspaces:
        try:
            if workspace.productType == '':
                continue
        except Exception:
            continue
        for panel in workspace.toolbarPanels:
            control = explore_control(panel.controls, workspace)
            if control:
                return control
    return None

def explore_control(controls, workspace):
    for control in controls:
        if control.objectType == adsk.core.CommandControl.classType():
            try:
                cmd_id = control.commandDefinition.id
                cmd_def_workspaces_map_[cmd_id].add(workspace.id)
                #print("READ", cmd_id)
            except Exception as e:
                #print(f"Could not read commandDefintion for {control.id}", control)
                pass
        elif control.objectType == adsk.core.DropDownControl.classType():
            return explore_control(control.controls, workspace)
    return None

def find_cmd_workspaces(cmd_id):
    return cmd_def_workspaces_map_.get(cmd_id, [])

def deduplicate_hotkeys(hotkeys):
    ids = set()
    filtered = []
    for hotkey in hotkeys:
        # Using command_name instead of command_id, as the names duplicate
        hid = (hotkey.command_name, hotkey.command_argument)
        if hid in ids:
            print("DUP: ", hotkey.command_id, hotkey.command_argument, hotkey.key_sequence)
            continue
        ids.add(hid)
        filtered.append(hotkey)
    return filtered

def find_options_files():
    CSIDL_APPDATA = 26
    SHGFP_TYPE_CURRENT = 0
    # SHGetFolderPath is deprecated, but SHGetKnownFolderPath is much more cumbersome to use.
    # Could just use win32com in that case, to simplify things.
    roaming_path_buf = ctypes.create_unicode_buffer(ctypes.wintypes.MAX_PATH)
    ctypes.windll.shell32.SHGetFolderPathW(0, CSIDL_APPDATA, 0, SHGFP_TYPE_CURRENT, roaming_path_buf)

    roaming_path = pathlib.Path(roaming_path_buf.value)
    options_path = roaming_path / 'Autodesk' / 'Neutron Platform' / 'Options'
    options_files = list(options_path.glob('*/NGlobalOptions.xml'))
    return options_files

def parse_hotkeys(options_file):
    hotkeys = []
    xml = ET.parse(options_file)
    root = xml.getroot()
    json_element = root.find('./HotKeyGroup/HotKeyJSONString')
    value = json.loads(json_element.attrib['Value'])
    for h in value['hotkeys']:
        if 'hotkey_sequence' not in h:
            continue
        key_sequence = h['hotkey_sequence']

        # Move data extraction to separate function?
        # E.g. ! is used for shift+1, so we need to pull out the virtual keycode, to get just one key
        #vk, _ = fusion_key_to_vk(key_sequence)

        for hotkey_command in h['commands']:
            hotkey = Hotkey()
            hotkey.command_id = hotkey_command['command_id']
            hotkey.command_argument = hotkey_command['command_argument']
            hotkey.is_default = hotkey_command['isDefault']
            hotkey.key_sequence = key_sequence
            hotkeys.append(hotkey)
    return hotkeys

def delete_command_def():
    cmd_def = ui_.commandDefinitions.itemById(LIST_CMD_ID)
    if cmd_def:
        cmd_def.deleteMe()

def run(context):
    global app_
    global ui_
    global list_cmd_def_
    try:
        app_ = adsk.core.Application.get()
        ui_ = app_.userInterface
        
        ui_.terminateActiveCommand()
        delete_command_def()
        list_cmd_def_ = ui_.commandDefinitions.addButtonDefinition(LIST_CMD_ID,
                                                                   NAME,
                                                                   '',)

        events_manager_.add_handler(list_cmd_def_.commandCreated,
                                    adsk.core.CommandCreatedEventHandler,
                                    list_command_created_handler)

        list_cmd_def_.execute()

        # Keep the script running.
        adsk.autoTerminate(False)
    except:
        if ui_:
            ui_.messageBox('Failed:\n{}'.format(traceback.format_exc()))
