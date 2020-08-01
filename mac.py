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

# Platform-specific code

import pathlib

def fusion_key_to_keyboard_key(key_sequence):
    # TODO
    keys = key_sequence.split('+')
    return key_sequence, keys[-1]

def find_options_file(app):
    # Seems that Macs can have the files in different locations:
    # https://forums.autodesk.com/t5/fusion-360-support/cannot-find-fuision360-documents-locally-in-macos/td-p/8324149

    # * /Users/<username>/Library/Application Support/Autodesk
    # * /Users/<username>/Library/Containers/com.autodesk.mas.fusion360/Data/Library/Application Support/Autodesk
    # append: /Neutron Platform/Options/<user id>\<file.xml>
    # Idea: If neededn, check where our add-in is running to determine which path to use.

    autodesk_paths = [
        pathlib.Path.home() / 'Library/Application Support/Autodesk',
        pathlib.Path.home() / 'Library/Containers/com.autodesk.mas.fusion360/Data/Library/Application Support/Autodesk'
    ]

    for autodesk_path in autodesk_paths:
        if autodesk_path.exists():
            break
    else:
        raise Exception("Could not find Autodesk directory")

    options_path = autodesk_path / 'Neutron Platform' / 'Options' / app.userId / 'NGlobalOptions.xml'
    return options_path
