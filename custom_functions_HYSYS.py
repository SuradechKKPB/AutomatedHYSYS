class cWindow:
    def __init__(self):
        self._hwnd = None

    def BringToTop(self):
        win32gui.BringWindowToTop(self._hwnd)

    def SetAsForegroundWindow(self):
        win32gui.SetForegroundWindow(self._hwnd)

    def Maximize(self):
        win32gui.ShowWindow(self._hwnd, win32con.SW_MAXIMIZE)

    def setActWin(self):
        win32gui.SetActiveWindow(self._hwnd)

    def _window_enum_callback(self, hwnd, wildcard):
        '''Pass to win32gui.EnumWindows() to check all the opened windows'''
        if re.match(wildcard, str(win32gui.GetWindowText(hwnd))) != None:
            self._hwnd = hwnd

    def find_window_wildcard(self, wildcard):
        self._hwnd = None
        win32gui.EnumWindows(self._window_enum_callback, wildcard)

def bringHYSYS():
    try:
        wildcard = ".*hsc*."
        cW = cWindow()
        cW.find_window_wildcard(wildcard)
        cW.BringToTop()
        cW.Maximize()
        cW.SetAsForegroundWindow()

    except:
        f = open("log.txt", "w")
        f.write(traceback.format_exc())
        print(traceback.format_exc())

def bringExcel():
    try:
        wildcard = ".*xlsx*."
        cW = cWindow()
        cW.find_window_wildcard(wildcard)
        cW.BringToTop()
        cW.Maximize()
        cW.SetAsForegroundWindow()

    except:
        f = open("log.txt", "w")
        f.write(traceback.format_exc())
        print(traceback.format_exc())

def copy_hysys():
    print('Found HYSYS OK')
    pg.hotkey('ctrl','w') #Obtain Value
    GUI.destroy()
    run_case()

def is_OK():
    print('Found HYSYS OK')
    pg.hotkey('ctrl','w')
    GUI.destroy()

def run_case():
    bringExcel()
    pg.hotkey('ctrl','q') #Run HYSYS case and adjust values
    time.sleep(10)
    bringHYSYS()