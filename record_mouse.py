import ctypes
from ctypes import wintypes, Structure, c_ulong, byref
import time
import pickle
from datetime import datetime
import random
import os


CF_UNICODETEXT = 13

user32 = ctypes.WinDLL('user32', use_last_error=True)
kernel32 = ctypes.WinDLL('kernel32')

INPUT_MOUSE    = 0
INPUT_KEYBOARD = 1
INPUT_HARDWARE = 2
KEYEVENTF_EXTENDEDKEY = 0x0001
KEYEVENTF_KEYUP       = 0x0002
KEYEVENTF_UNICODE     = 0x0004
KEYEVENTF_SCANCODE    = 0x0008

MAPVK_VK_TO_VSC = 0

# msdn.microsoft.com/en-us/library/dd375731
VK_TAB  = 0x09
VK_MENU = 0x12

VK_NUMLOCK = 0x90
VK_SCROLLLOCK = 0x91
# C struct definitions

wintypes.ULONG_PTR = wintypes.WPARAM


OpenClipboard = user32.OpenClipboard
OpenClipboard.argtypes = wintypes.HWND,
OpenClipboard.restype = wintypes.BOOL
CloseClipboard = user32.CloseClipboard
CloseClipboard.restype = wintypes.BOOL
EmptyClipboard = user32.EmptyClipboard
EmptyClipboard.restype = wintypes.BOOL
GetClipboardData = user32.GetClipboardData
GetClipboardData.argtypes = wintypes.UINT,
GetClipboardData.restype = wintypes.HANDLE
SetClipboardData = user32.SetClipboardData
SetClipboardData.argtypes = (wintypes.UINT, wintypes.HANDLE)
SetClipboardData.restype = wintypes.HANDLE

GlobalLock = kernel32.GlobalLock
GlobalLock.argtypes = wintypes.HGLOBAL,
GlobalLock.restype = wintypes.LPVOID
GlobalUnlock = kernel32.GlobalUnlock
GlobalUnlock.argtypes = wintypes.HGLOBAL,
GlobalUnlock.restype = wintypes.BOOL
GlobalAlloc = kernel32.GlobalAlloc
GlobalAlloc.argtypes = (wintypes.UINT, ctypes.c_size_t)
GlobalAlloc.restype = wintypes.HGLOBAL
GlobalSize = kernel32.GlobalSize
GlobalSize.argtypes = wintypes.HGLOBAL,
GlobalSize.restype = ctypes.c_size_t
GMEM_MOVEABLE = 0x0002
GMEM_ZEROINIT = 0x0040

unicode_type = type(u'')

def clipboard_get():
    text = None
    OpenClipboard(None)
    handle = GetClipboardData(CF_UNICODETEXT)
    pcontents = GlobalLock(handle)
    size = GlobalSize(handle)
    if pcontents and size:
        raw_data = ctypes.create_string_buffer(size)
        ctypes.memmove(raw_data, pcontents, size)
        text = raw_data.raw.decode('utf-16le').rstrip(u'\0')
    GlobalUnlock(handle)
    CloseClipboard()
    return text

def clipboard_put(s):
    if not isinstance(s, unicode_type):
        s = s.decode('mbcs')
    data = s.encode('utf-16le')
    OpenClipboard(None)
    EmptyClipboard()
    handle = GlobalAlloc(GMEM_MOVEABLE | GMEM_ZEROINIT, len(data) + 2)
    pcontents = GlobalLock(handle)
    ctypes.memmove(pcontents, data, len(data))
    GlobalUnlock(handle)
    SetClipboardData(CF_UNICODETEXT, handle)
    CloseClipboard()

#'LeftMouseButton': '0x01', 'RightMouseButton': '0x02', 
KEY_CODES = {
            ';:':'0xba',
            '+=':'0xbb',
            ',<':'0xbc',
            '-_':'0xbd',
            '.>':'0xbe',
            '/?':'0xbf',
            '`~':'0xc0',
            '[{':'0xdb',
            '\|':'0xdc',
            ']}':'0xdd',
            'quote':'0xde',
            'ControlBreak': '0x03',
            'MiddleMouseButton': '0x04',
            'X1MouseButton': '0x05', 'X2MouseButton': '0x06', 'Backspace': '0x08', 'Tab': '0x09',
            'Clear': '0x0c',
            'Enter': '0x0d', 'Pause': '0x13', 'CapsLock': '0x14', 'Escape': '0x1B', 'Space': '0x20',
            'PageUp': '0x21',
            'PageDown': '0x22', 'End': '0x23', 'Home': '0x24', 'LeftArrow': '0x25',
            'UpArrow': '0x26',
            'RightArrow': '0x27',
            'DownArrow': '0x28', 'Select': '0x29', 'Print': '0x2a', 'Execute': '0x2b',
            'PrintScreen': '0x2c',
            'Ins': '0x2d',
            'Delete': '0x2e', 'Help': '0x2f', 'Key0': '0x30', 'Key1': '0x31', 'Key2': '0x32',
            'Key3': '0x33',
            'Key4': '0x34',
            'Key5': '0x35', 'Key6': '0x36', 'Key7': '0x37', 'Key8': '0x38', 'Key9': '0x39',
            'A': '0x41',
            'B': '0x42',
            'C': '0x43', 'D': '0x44', 'E': '0x45', 'F': '0x46', 'G': '0x47', 'H': '0x48',
            'I': '0x49',
            'J': '0x4a',
            'K': '0x4b', 'L': '0x4c', 'M': '0x4d', 'N': '0x4e', 'O': '0x4f', 'P': '0x50',
            'Q': '0x51',
            'R': '0x52',
            'S': '0x53', 'T': '0x54', 'U': '0x55', 'V': '0x56', 'W': '0x57', 'X': '0x58',
            'Y': '0x59', 'Z': '0x5a',
            'LeftWindowsKey': '0x5b', 'RightWindowsKey': '0x5c', 'Application': '0x5d',
            'Sleep': '0x5f',
            'NumpadKey0': '0x60',
            'NumpadKey1': '0x61', 'NumpadKey2': '0x62', 'NumpadKey3': '0x63', 'NumpadKey4': '0x64',
            'NumpadKey5': '0x65',
            'NumpadKey6': '0x66', 'NumpadKey7': '0x67', 'NumpadKey8': '0x68', 'NumpadKey9': '0x69',
            'Multiply': '0x6a',
            'Add': '0x6b', 'Separator': '0x6c', 'Subtract': '0x6d', 'Decimal': '0x6e',
            'Divide': '0x6f',
            'F1': '0x70',
            'F2': '0x71', 'F3': '0x72', 'F4': '0x73', 'F5': '0x74', 'F6': '0x75', 'F7': '0x76',
            'F8': '0x77',
            'F9': '0x78',
            'F10': '0x79', 'F11': '0x7a', 'F12': '0x7b', 'F13': '0x7c', 'F14': '0x7d',
            'F15': '0x7e',
            'F16': '0x7f',
            'F17': '0x80', 'F18': '0x81', 'F19': '0x82', 'F20': '0x83', 'F21': '0x84',
            'F22': '0x85',
            'F23': '0x86',
            'F24': '0x87', 'NumLock': '0x90', 'ScrollLock': '0x91', 'LeftShift': '0xa0',
            'RightShift': '0xa1',
            'LeftControl': '0xa2', 'RightControl': '0xa3', 'LeftMenu': '0xa4', 'RightMenu': '0xa5',
            'BrowserBack': '0xa6',
            'BrowserRefresh': '0xa8', 'BrowserStop': '0xa9', 'BrowserSearch': '0xaa',
            'BrowserFavorites': '0xab',
            'BrowserHome': '0xac', 'VolumeMute': '0xad', 'VolumeDown': '0xae', 'VolumeUp': '0xaf',
            'NextTrack': '0xb0',
            'PreviousTrack': '0xb1', 'StopMedia': '0xb2', 'PlayMedia': '0xb3',
            'StartMailKey': '0xb4',
            'SelectMedia': '0xb5',
            'StartApplication1': '0xb6', 'StartApplication2': '0xb7'}


class MOUSEINPUT(ctypes.Structure):
    _fields_ = (("dx",          wintypes.LONG),
                ("dy",          wintypes.LONG),
                ("mouseData",   wintypes.DWORD),
                ("dwFlags",     wintypes.DWORD),
                ("time",        wintypes.DWORD),
                ("dwExtraInfo", wintypes.ULONG_PTR))

class KEYBDINPUT(ctypes.Structure):
    _fields_ = (("wVk",         wintypes.WORD),
                ("wScan",       wintypes.WORD),
                ("dwFlags",     wintypes.DWORD),
                ("time",        wintypes.DWORD),
                ("dwExtraInfo", wintypes.ULONG_PTR))

    def __init__(self, *args, **kwds):
        super(KEYBDINPUT, self).__init__(*args, **kwds)
        # some programs use the scan code even if KEYEVENTF_SCANCODE
        # isn't set in dwFflags, so attempt to map the correct code.
        if not self.dwFlags & KEYEVENTF_UNICODE:
            self.wScan = user32.MapVirtualKeyExW(self.wVk,
                                                 MAPVK_VK_TO_VSC, 0)

class HARDWAREINPUT(ctypes.Structure):
    _fields_ = (("uMsg",    wintypes.DWORD),
                ("wParamL", wintypes.WORD),
                ("wParamH", wintypes.WORD))

class INPUT(ctypes.Structure):
    class _INPUT(ctypes.Union):
        _fields_ = (("ki", KEYBDINPUT),
                    ("mi", MOUSEINPUT),
                    ("hi", HARDWAREINPUT))
    _anonymous_ = ("_input",)
    _fields_ = (("type",   wintypes.DWORD),
                ("_input", _INPUT))

LPINPUT = ctypes.POINTER(INPUT)

def _check_count(result, func, args):
    if result == 0:
        raise ctypes.WinError(ctypes.get_last_error())
    return args

user32.SendInput.errcheck = _check_count
user32.SendInput.argtypes = (wintypes.UINT, # nInputs
                             LPINPUT,       # pInputs
                             ctypes.c_int)  # cbSize


def PressKey(hexKeyCode):
    x = INPUT(type=INPUT_KEYBOARD,
              ki=KEYBDINPUT(wVk=hexKeyCode))
    user32.SendInput(1, ctypes.byref(x), ctypes.sizeof(x))

def ReleaseKey(hexKeyCode):
    x = INPUT(type=INPUT_KEYBOARD,
              ki=KEYBDINPUT(wVk=hexKeyCode,
                            dwFlags=KEYEVENTF_KEYUP))
    user32.SendInput(1, ctypes.byref(x), ctypes.sizeof(x))

def AltTab():
    """Press Alt+Tab and hold Alt key for 2 seconds
    in order to see the overlay.
    """
    PressKey(VK_MENU)   # Alt
    PressKey(VK_TAB)    # Tab
    ReleaseKey(VK_TAB)  # Tab~
    time.sleep(2)
    ReleaseKey(VK_MENU) # Alt~

    
def NumL():
    PressKey(VK_NUMLOCK)
    ReleaseKey(VK_NUMLOCK)
    
    
def click():
    # Functions
    user32.SetCursorPos(600, 600)
    user32.mouse_event(2, 0, 0, 0,0) # left down
    user32.mouse_event(4, 0, 0, 0,0) # left up
    
def queryMousePosition():
    pt = POINT()
    user32.GetCursorPos(byref(pt))
    return { "x": pt.x, "y": pt.y}
    
class POINT(Structure):
	_fields_ = [("x", c_ulong), ("y", c_ulong)]
    
    

def write_data(data, name = None):
    now = datetime.now()
    dt_string = name if name else now.strftime("%d-%m-%Y--%H-%M")
    with open(r'{0}{1}.pkl'.format(HOMEDIR,dt_string), 'wb') as outp:
        pickle.dump(data, outp, pickle.HIGHEST_PROTOCOL)

        
        
def detect_click():
    l = 1 if user32.GetAsyncKeyState(0x01) not in [0, 1] else 0
    r = 1 if user32.GetAsyncKeyState(0x02) not in [0, 1] else 0
    return [l, r]

    
    
def keyboard_pressed():
    pressed = []
    
    for keyName, keyNum in KEY_CODES.items():
        keyN = int(keyNum, 16)
        if user32.GetAsyncKeyState(keyN) != 0:
            pressed.append(keyName)
    return pressed
    
def record_mouse(filename=None):
    release_all_keys()
    pack = []
    moves = []
    state_prev = [{},[0,0],[]]
    tt = 0
    while 1:
        coord = queryMousePosition()
        key = keyboard_pressed()
        clicks = detect_click()
        state_cur = [coord, clicks, key]

        if 'End' in state_cur[2]:
            for keyName in state_cur[2]:
                ReleaseKey(int(KEY_CODES[keyName],16))

            pack.append(moves)
            pack = [ll for ll in pack if ll]
            write_data(pack,filename)

            break
        
        
        if state_cur != state_prev:
            
            item = [int(time.time()*1000.0), *state_cur]
            moves.append(item)
            if key or (1 in clicks and 1 not in state_prev[1]) or (1 not in clicks and 1 in state_prev[1]):
                print(item)
                
            if 1 not in state_cur[1] and 1 in state_prev[1]:
                pack.append(moves)
                moves = [item]
                
            state_prev = list(state_cur)
             


def print_like_human(data):
    #проверить раскладку!
    for i in data:
        if i.isalpha():
            if i.isupper():
                PressKey(int(KEY_CODES['LeftShift'],16))
                time.sleep(random.randint(50, 250)*0.001)
                PressKey(int(KEY_CODES[i],16))
                time.sleep(random.randint(50, 300)*0.001)
                ReleaseKey(int(KEY_CODES['LeftShift'],16))
                time.sleep(random.randint(50, 250)*0.001)
                ReleaseKey(int(KEY_CODES[i],16))
                time.sleep(random.randint(50, 250)*0.001)
            else:
                PressKey(int(KEY_CODES[i.upper()],16))
                time.sleep(random.randint(50, 300)*0.001)
                ReleaseKey(int(KEY_CODES[i.upper()],16))
                time.sleep(random.randint(50, 250)*0.001)

        if i.isdigit():
            PressKey(int(KEY_CODES['Key'+i],16))
            time.sleep(random.randint(50, 300)*0.001)
            ReleaseKey(int(KEY_CODES['Key'+i],16))
            time.sleep(random.randint(50, 250)*0.001)

        if i in ';:+=,<-_.>/?`~[{\|]}':
            for key, value in KEY_CODES.items():
                if i in key:
                    if key.startswith(i):
                        PressKey(int(KEY_CODES[key],16))
                        time.sleep(random.randint(50, 300)*0.001)
                        ReleaseKey(int(KEY_CODES[key],16))
                        time.sleep(random.randint(50, 250)*0.001)
                    else:
                        PressKey(int(KEY_CODES['LeftShift'],16))
                        time.sleep(random.randint(50, 350)*0.001)
                        PressKey(int(KEY_CODES[key],16))
                        time.sleep(random.randint(50, 500)*0.001)
                        ReleaseKey(int(KEY_CODES['LeftShift'],16))
                        time.sleep(random.randint(50, 270)*0.001)
                        ReleaseKey(int(KEY_CODES[key],16))
                        time.sleep(random.randint(50, 270)*0.001)

        if i in ')!@#$%^&*(':
            ind = ')!@#$%^&*('.find(i)
            PressKey(int(KEY_CODES['LeftShift'],16))
            time.sleep(random.randint(50, 350)*0.001)
            PressKey(int(KEY_CODES['Key'+str(ind)],16))
            time.sleep(random.randint(50, 500)*0.001)
            ReleaseKey(int(KEY_CODES['LeftShift'],16))
            time.sleep(random.randint(50, 270)*0.001)
            ReleaseKey(int(KEY_CODES['Key'+str(ind)],16))
            time.sleep(random.randint(50, 270)*0.001)
            
        if i == ' ':
            PressKey(int(KEY_CODES['Space'],16))
            time.sleep(random.randint(50, 500)*0.001)
            ReleaseKey(int(KEY_CODES['Space'],16))
            time.sleep(random.randint(50, 270)*0.001)

                
                
def load_data(name):
    with open(r'{0}{1}.pkl'.format(HOMEDIR,name), 'rb') as f:
        return pickle.load(f)
    

def play_whole_data(segments):
    for segment in segments:
        play_data(segment)
    
    
def play_data(data):
    #print(data)
    time_t = data[0][0]
    state_prev = data[0]
    '''
    [
        321837129871,
        {'x':11,'y':11}, 
        [1,0],
        [A, B, Ctrl]
    ]
    
    '''
    state_prev[0] =state_prev[0]
    state_prev[3] =[]
    data.insert(0, state_prev)
    for action in data[1:]:
        #print(action)
        
        user32.SetCursorPos(int(action[1]['x']), int(action[1]['y']))
            
        # cliicks left
        if action[2][0] and not state_prev[2][0]:
            user32.mouse_event(2, 0, 0, 0,0)
            print(action)
        elif not action[2][0] and state_prev[2][0]:
            user32.mouse_event(4, 0, 0, 0,0)
            print(action)
        # cliicks right
        if action[2][1] and not state_prev[2][1]:
            user32.mouse_event(0x0008, 0, 0, 0,0)
            print(action)
        elif not action[2][1] and state_prev[2][1]:
            user32.mouse_event(0x0010, 0, 0, 0,0)
            print(action)
            

        # press new btn
        for key in action[3][::-1]:
            if key not in state_prev[3]:
                PressKey(int(KEY_CODES[key],16))
        
        # release btns
        for key in state_prev[3]:
            if key not in action[3]:
                ReleaseKey(int(KEY_CODES[key],16))
    
        state_prev = list(action)
        time.sleep((action[0]-time_t)*0.001)
        time_t = action[0]

def release_all_keys():
    key = keyboard_pressed()
    for keyName in key:
        ReleaseKey(int(KEY_CODES[keyName],16))
        
        
if __name__ == "__main__":
    '''
    запускаешь скрипт и:
        Home - начать запись
            во время записи:
                Home - отделить сегмент и записывать в новый
                End  - завершить запись
    
        PageDown - проиграть первый сегмент записи
        
    
    функции:
        clipboard_get() - возвращает значение из буфера обмена
        clipboard_put(s) - помещает строку s в буфер обмена
        print_like_human(s) - типа как человек набирает текст
    '''
    HOMEDIR = f'C://Users/{os.getlogin()}/Desktop/'
    version = 0.113
    

    while(1):
        keys = keyboard_pressed()
        if 'Home' in keys:
            while 'Home' in keys:
                keys = keyboard_pressed()
                pass
            record_mouse('1')
            release_all_keys()
            break
        
        if 'PageDown' in keys:
            while 'PageDown' in keys:
                keys = keyboard_pressed()
                pass
            data = load_data('1')
            play_whole_data(data)
            break


