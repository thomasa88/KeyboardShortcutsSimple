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
import operator
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
UNKNOWN_WORKSPACE = 'UNKNOWN'

app_ = None
ui_ = None
events_manager_ = thomasa88lib.events.EventsManger(NAME)
list_cmd_def_ = None
cmd_def_workspaces_map_ = None
used_workspaces_ids_ = None
sorted_workspaces_ = None
ws_filter_map_ = None

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

    events_manager_.add_handler(cmd.inputChanged,
                                adsk.core.InputChangedEventHandler,
                                input_changed_handler)
    events_manager_.add_handler(cmd.execute,
                                adsk.core.CommandEventHandler,
                                execute_handler)

    inputs = cmd.commandInputs

    workspace_input = inputs.addDropDownCommandInput('workspace',
                                                     'Show',
                                                     adsk.core.DropDownStyles.LabeledIconDropDownStyle)
    global ws_filter_map_
    ws_filter_map_ = []
    workspace_input.listItems.add('All', True, '', -1)
    ws_filter_map_.append(None)
    #workspace_input.listItems.addSeparator(-1)
    workspace_input.listItems.add('----------------------------------', False, '', -1)
    ws_filter_map_.append('SEPARATOR')
    workspace_input.listItems.add('General', False, '', -1)
    ws_filter_map_.append(UNKNOWN_WORKSPACE)
    for workspace in sorted_workspaces_:
        workspace_input.listItems.add(workspace.name, False, '', -1)
        ws_filter_map_.append(workspace.id)
    
    only_user_input = inputs.addBoolValueInput('only_user', 'Only user-defined          ', True, '', True)

    inputs.addTextBoxCommandInput('list', '', get_hotkeys_str(only_user=only_user_input.value), 30, False)
    inputs.addTextBoxCommandInput('list_info', '', '* = User-defined', 1, True)

def input_changed_handler(args):
    eventArgs = adsk.core.InputChangedEventArgs.cast(args)

    if eventArgs.input.id == 'list':
        retun

    inputs = eventArgs.inputs
    only_user_input = inputs.itemById('only_user')
    
    workspace_input = inputs.itemById('workspace')
    workspace_filter = ws_filter_map_[workspace_input.selectedItem.index]
    
    list_input = inputs.itemById('list')
    list_input.formattedText = get_hotkeys_str(only_user_input.value, workspace_filter)

def execute_handler(args):
    # Force the termination of the command.
    adsk.terminate()

def get_hotkeys_str(only_user=False, workspace_filter=None):
    options_files = find_options_files()
    # TODO: Pick the correct user/profile if there are multiple options files
    hotkeys = parse_hotkeys(options_files[0])

    # Make sure to filter before any de-dup operation
    if only_user:
        hotkeys = [hotkey for hotkey in hotkeys if not hotkey.is_default]

    # HTML table is hard to copy-paste. Use fixed-width font instead.
    # Supported HTML in QT: https://doc.qt.io/archives/qt-4.8/richtext-html-subset.html
    string = '<pre>'
    
    hotkeys = map_command_names(hotkeys)
    ns_hotkeys = namespace_group_hotkeys(hotkeys)
    for workspace_id, hotkeys in ns_hotkeys.items():
        if workspace_filter and workspace_id != workspace_filter:
            continue
        hotkeys = deduplicate_hotkeys(hotkeys)
        if workspace_id != UNKNOWN_WORKSPACE:
            workspace_name = ui_.workspaces.itemById(workspace_id).name
        else:
            workspace_name = 'General'
        string += f'<b>{workspace_name}</b><br>'
        for hotkey in sorted(hotkeys, key=lambda  h: h.command_name):
            name = hotkey.command_name
            if not hotkey.is_default:
                name += '*'
            string += f'{name:30} {hotkey.key_sequence}<br>'
        string += '<br>'
    string += '</pre>'
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
        if not workspaces:
            ns_hotkeys[UNKNOWN_WORKSPACE].append(hotkey)
    return ns_hotkeys

def build_cmd_def_workspaces_map():
    global cmd_def_workspaces_map_
    global used_workspaces_ids_
    cmd_def_workspaces_map_ = defaultdict(set)
    used_workspaces_ids_ = set()
    for workspace in ui_.workspaces:
        try:
            if workspace.productType == '':
                continue
        except Exception:
            continue
        for panel in workspace.toolbarPanels:
            explore_controls(panel.controls, workspace)
            

def explore_controls(controls, workspace):
    global used_workspaces_ids_
    for control in controls:
        if control.objectType == adsk.core.CommandControl.classType():
            try:
                cmd_id = control.commandDefinition.id
            except RuntimeError as e:
                #print(f"Could not read commandDefintion for {control.id}", control)
                continue
            cmd_def_workspaces_map_[cmd_id].add(workspace.id)
            used_workspaces_ids_.add(workspace.id)
            #print("READ", cmd_id)
        elif control.objectType == adsk.core.DropDownControl.classType():
            return explore_controls(control.controls, workspace)
    return None

def find_cmd_workspaces(cmd_id):
    return cmd_def_workspaces_map_.get(cmd_id, [ UNKNOWN_WORKSPACE ])

def deduplicate_hotkeys(hotkeys):
    ids = set()
    filtered = []
    for hotkey in hotkeys:
        # Using command_name instead of command_id, as the names duplicate
        hid = (hotkey.command_name, hotkey.command_argument)
        if hid in ids:
            #print("DUP: ", hotkey.command_id, hotkey.command_argument, hotkey.key_sequence)
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

        build_cmd_def_workspaces_map()

        global sorted_workspaces_
        sorted_workspaces_ = sorted([ui_.workspaces.itemById(w_id) for w_id in used_workspaces_ids_],
                                    key=lambda w: w.name)

        list_cmd_def_.execute()

        # Keep the script running.
        adsk.autoTerminate(False)
    except:
        if ui_:
            ui_.messageBox('Failed:\n{}'.format(traceback.format_exc()))
