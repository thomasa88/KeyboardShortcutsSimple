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
from tkinter import Tk
import xml.etree.ElementTree as ET

NAME = 'Keyboard Shortcuts Simple'

FILE_DIR = os.path.dirname(os.path.realpath(__file__))


# Import relative path to avoid namespace pollution
from .thomasa88lib import utils
from .thomasa88lib import events
from .thomasa88lib import error

from .version import VERSION
if os.name == 'nt':
    from . import windows as platform
else:
    from . import mac as platform

# Force modules to be fresh during development
import importlib
importlib.reload(thomasa88lib.utils)
importlib.reload(thomasa88lib.events)
importlib.reload(thomasa88lib.error)
importlib.reload(platform)


LIST_CMD_ID = 'thomasa88_keyboardShortcutsSimpleList'
UNKNOWN_WORKSPACE = 'UNKNOWN'

app_ = None
ui_ = None
error_catcher_ = thomasa88lib.error.ErrorCatcher(msgbox_in_debug=False)
events_manager_ = thomasa88lib.events.EventsManager(error_catcher_)
list_cmd_def_ = None
cmd_def_workspaces_map_ = None
used_workspaces_ids_ = None
sorted_workspaces_ = None
ws_filter_map_ = None
ns_hotkeys_ = None
copy_button_args_ = ('copy', False,
                     thomasa88lib.utils.get_fusion_deploy_folder() + '/Electron/UI/Resources/Icons/Copy',
                     -1)

class Hotkey:
    pass

def list_command_created_handler(args):
    eventArgs = adsk.core.CommandCreatedEventArgs.cast(args)

    get_data()

    # The nifty thing with cast is that code completion then knows the object type
    cmd = adsk.core.Command.cast(args.command)
    cmd.isRepeatable = False
    cmd.isExecutedWhenPreEmpted = False
    cmd.isOKButtonVisible = False
    cmd.setDialogMinimumSize(350, 200)
    cmd.setDialogInitialSize(400, 500)

    events_manager_.add_handler(cmd.inputChanged,
                                adsk.core.InputChangedEventHandler,
                                input_changed_handler)
    events_manager_.add_handler(cmd.destroy,
                                adsk.core.CommandEventHandler,
                                destroy_handler)

    inputs = cmd.commandInputs

    workspace_input = inputs.addDropDownCommandInput('workspace',
                                                     'Workspace',
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
    shortcut_sort_input = inputs.addBoolValueInput('shortcut_sort', 'Sort by shortcut keys', True, '', False)

    inputs.addTextBoxCommandInput('list', '', get_hotkeys_str(only_user=only_user_input.value), 30, False)
    inputs.addTextBoxCommandInput('list_info', '', '* = User-defined', 1, True)

    copy_input = inputs.addButtonRowCommandInput('copy', 'Copy to clipboard', False)
    copy_input.isMultiSelectEnabled = False
    copy_input.listItems.add(*copy_button_args_)

def get_data():
    # Build on every invocation, in case keys have changed
    # (Does not really matter for a Script)
    build_cmd_def_workspaces_map()
    global sorted_workspaces_
    sorted_workspaces_ = sorted([ui_.workspaces.itemById(w_id) for w_id in used_workspaces_ids_],
                                 key=lambda w: w.name)

    global ns_hotkeys_
    options_files = find_options_files()
    # TODO: Pick the correct user/profile if there are multiple options files
    hotkeys = parse_hotkeys(options_files[0])
    hotkeys = map_command_names(hotkeys)
    ns_hotkeys_ = namespace_group_hotkeys(hotkeys)

def input_changed_handler(args):
    eventArgs = adsk.core.InputChangedEventArgs.cast(args)

    inputs = eventArgs.inputs
    if eventArgs.input.id == 'list':
        return
    
    list_input = inputs.itemById('list')
    only_user_input = inputs.itemById('only_user')
    shortcut_sort_input = inputs.itemById('shortcut_sort')    
    workspace_input = inputs.itemById('workspace')

    workspace_filter = ws_filter_map_[workspace_input.selectedItem.index]

    if eventArgs.input.id == 'copy':
        copy_input = eventArgs.input
        # Does not work: copy_input.listItems[0].isSelected = False
        #copy_button = copy_input.listItems[0]
        copy_input.listItems.clear()
        copy_input.listItems.add(*copy_button_args_)
        copy_to_clipboard(get_hotkeys_str(only_user_input.value, workspace_filter,
                                          sort_by_shortcut=shortcut_sort_input.value,
                                          html=False))
    else:
        # Update list
        list_input.formattedText = get_hotkeys_str(only_user_input.value, workspace_filter,
                                                   sort_by_shortcut=shortcut_sort_input.value)

def destroy_handler(args):
    # Force the termination of the command.
    adsk.terminate()
    events_manager_.clean_up()

def get_hotkeys_str(only_user=False, workspace_filter=None, sort_by_shortcut=False, html=True):
    # HTML table is hard to copy-paste. Use fixed-width font instead.
    # Supported HTML in QT: https://doc.qt.io/archives/qt-4.8/richtext-html-subset.html

    def header(text, text_underline='-'):
        if html:
            return f'<b>{text}</b><br>'
        else:
            return f'{text}\n{text_underline * len(text)}\n'
        
    def newline():
        return '<br>' if html else '\n'

    def sort_key(hotkey):
        if not sort_by_shortcut:
            return hotkey.command_name
        else:
            return hotkey.keyboard_key_sequence

    string = ''
    if html:
        string += '<pre>'
    else:
        string += header('Fusion 360 Keyboard Shortcuts', '=') + '\n'
    
    for workspace_id, hotkeys in ns_hotkeys_.items():
        if workspace_filter and workspace_id != workspace_filter:
            continue
        # Make sure to filter before any de-dup operation
        if only_user:
            hotkeys = [hotkey for hotkey in hotkeys if not hotkey.is_default]
        if not hotkeys:
            continue
        hotkeys = deduplicate_hotkeys(hotkeys)
        if workspace_id != UNKNOWN_WORKSPACE:
            workspace_name = ui_.workspaces.itemById(workspace_id).name
        else:
            workspace_name = 'General'

        if html:
            string += f'<b>{workspace_name}</b><br>'
        else:
            string += f'{workspace_name}\n{"=" * len(workspace_name)}\n'

        for hotkey in sorted(hotkeys, key=sort_key):
            name = hotkey.command_name
            if not hotkey.is_default:
                name += '*'
            string += f'{name:30} {hotkey.keyboard_key_sequence}'
            string += newline()
        string += newline()
    
    if html:
        string += '</pre>'
    else:
        string += '* = User-defined'
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
            #print("DUP: ", hotkey.command_id, hotkey.command_argument, hotkey.fusion_key_sequence)
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
        fusion_key_sequence = h['hotkey_sequence']

        # Move data extraction to separate function?
        # E.g. ! is used for shift+1, so we need to pull out the virtual keycode,
        # to get the actual key that the user needs to press. (E.g. '=' is placed
        # on different keys on different keyboards and some use shift.)
        keyboard_key_sequence = platform.fusion_key_to_keyboard_key(fusion_key_sequence)

        for hotkey_command in h['commands']:
            hotkey = Hotkey()
            hotkey.command_id = hotkey_command['command_id']
            hotkey.command_argument = hotkey_command['command_argument']
            hotkey.is_default = hotkey_command['isDefault']
            hotkey.fusion_key_sequence = fusion_key_sequence
            hotkey.keyboard_key_sequence = keyboard_key_sequence
            hotkeys.append(hotkey)
    return hotkeys

def copy_to_clipboard(string):
    # From https://stackoverflow.com/a/25476462/106019
    r = Tk()
    r.withdraw()
    r.clipboard_clear()
    r.clipboard_append(string)
    r.update() # now it stays on the clipboard after the window is closed
    r.destroy()

def delete_command_def():
    cmd_def = ui_.commandDefinitions.itemById(LIST_CMD_ID)
    if cmd_def:
        cmd_def.deleteMe()

def run(context):
    global app_
    global ui_
    global list_cmd_def_
    with error_catcher_:
        app_ = adsk.core.Application.get()
        ui_ = app_.userInterface
        
        if ui_.activeCommand == LIST_CMD_ID:
            ui_.terminateActiveCommand()
        delete_command_def()
        list_cmd_def_ = ui_.commandDefinitions.addButtonDefinition(LIST_CMD_ID,
                                                                   f'{NAME} {VERSION}',
                                                                   '',)

        events_manager_.add_handler(list_cmd_def_.commandCreated,
                                    adsk.core.CommandCreatedEventHandler,
                                    list_command_created_handler)

        list_cmd_def_.execute()

        # Keep the script running.
        adsk.autoTerminate(False)
