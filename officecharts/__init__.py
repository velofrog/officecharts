import ctypes
import os
import re
import sysconfig
from .container import create_chart
from .clipboard import _set_os_clipboard


shared_library = None

if sysconfig.get_config_var("EXT_SUFFIX") is not None:
    shared_library = os.path.join(os.path.dirname(__file__), f"os_clipboard{sysconfig.get_config_var('EXT_SUFFIX')}")
else:
    for root, dirs, files in os.walk(os.path.dirname(__file__)):
        for file in files:
            if re.search("os_clipboard(.*)(dll|so|dylib)$", file):
                shared_library = os.path.join(root, file)
                break

if shared_library:
    os_clipboard = ctypes.cdll.LoadLibrary(shared_library)
    os_clipboard.send_officedrawing_to_clipboard.restype = None
    os_clipboard.send_officedrawing_to_clipboard.argtypes = [ctypes.c_size_t, ctypes.c_void_p]
    _set_os_clipboard(os_clipboard)

