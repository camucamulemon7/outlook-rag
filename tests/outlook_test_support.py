"""COM import substitutes for synthetic Outlook tests on non-Windows hosts."""
import sys
import os
from contextlib import nullcontext
from types import ModuleType
from unittest.mock import Mock, patch


def isolated_environment(values):
    """Keep OS runtime essentials while removing application configuration.

    Native Windows libraries need SystemRoot and a writable TEMP/TMP even in
    an isolated fixture; clearing them breaks Winsock and Lance spill files.
    """
    essential = {"SYSTEMROOT", "WINDIR", "TEMP", "TMP", "PATH", "COMSPEC", "PATHEXT"}
    return {**{key: value for key, value in os.environ.items() if key.upper() in essential}, **values}


def synthetic_com():
    if sys.platform == 'win32':
        return nullcontext()
    pythoncom = ModuleType('pythoncom')
    pythoncom.CoInitialize = Mock()
    pythoncom.CoUninitialize = Mock()
    pywintypes = ModuleType('pywintypes')
    pywintypes.com_error = type('com_error', (Exception,), {})
    win32com = ModuleType('win32com')
    win32com.__path__ = []
    client = ModuleType('win32com.client')
    # Every test must provide its synthetic Outlook object explicitly.
    client.Dispatch = Mock(side_effect=AssertionError('Synthetic Outlook fixture required'))
    win32com.client = client
    return patch.dict(sys.modules, {
        'pythoncom': pythoncom, 'pywintypes': pywintypes,
        'win32com': win32com, 'win32com.client': client,
    })
