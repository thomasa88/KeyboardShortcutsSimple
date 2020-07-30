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

import ctypes

# Virtual Keycodes are available here:
# https://docs.microsoft.com/en-us/windows/win32/inputdev/virtual-key-codes
FUSION_VK_MAPPING = { 'Slash': ord('/'),
                      'backspace': 0x08,
                      'delete': 0x2e,
                      'escape': 0x1b,
                      'return': 0x0d,
                      'space': 0x20,
                      'f1': 0x70,
                      'f2': 0x71,
                      'f3': 0x72,
                      'f4': 0x73,
                      'f5': 0x74,
                      'f6': 0x75,
                      'f7': 0x76,
                      'f8': 0x77,
                      'f9': 0x78,
                      'f10': 0x79,
                      'f11': 0x7a,
                      'f12': 0x7b,
                      'f13': 0x7c,
                      'f14': 0x7d,
                      'f15': 0x7e,
                      'f16': 0x7f,
                      'f17': 0x80,
                      'f18': 0x81,
                      'f19': 0x82,
                      'f20': 0x83,
                      'f21': 0x84,
                      'f22': 0x85,
                      'f23': 0x86,
                      'f24': 0x87,
                      }

def fusion_key_to_keyboard_key(key_sequence):
    keys = key_sequence.split('+')

    vk, shift_state = fusion_key_to_vk(keys[-1])

    # Get the scancode from the virtual key and then get the label from the scan code
    MAPVK_VK_TO_VSC_EX = 4
    keyname_buf = ctypes.create_unicode_buffer(32)
    input_locale = ctypes.windll.User32.GetKeyboardLayout(0)
    scan_code = ctypes.windll.User32.MapVirtualKeyExW(vk, MAPVK_VK_TO_VSC_EX, input_locale)    
    ret = ctypes.windll.User32.GetKeyNameTextW(scan_code << 16, keyname_buf, ctypes.sizeof(keyname_buf))
    if ret > 0:
        label = keyname_buf.value
    else:
        label = keys[-1]
        print(f"Failed to get name for vk 0x{vk:x}: {ctypes.GetLastError()}")
    
    keys[-1] = label
    return '+'.join(keys)

def fusion_key_to_vk(fusion_key):
    if len(fusion_key) == 1:
        char = ord(fusion_key)
    else:
        try:
            char = FUSION_VK_MAPPING[fusion_key]
        except KeyError:
            print(f"No keyboard mapping for \"{fusion_key}\". Ignoring.")
            return None, None
    ret = ctypes.windll.User32.VkKeyScanW(char)
    keycode = ret & 0xff
    shift_state = ret >> 8
    return keycode, shift_state
