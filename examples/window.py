import win32gui
import win32con
import psutil
import win32process


class Window:
    """
    Window class
    :param window_text: Entire string or substring from titlebar of desired window
    :type window_text: str
    :param com_obj: COM Object
    :type com_obj: object
    """

    def __init__(self, window_text, com_obj):
        """
        :param window_text: Entire string or substring from titlebar of desired window
        :type window_text: str
        :param com_obj: COM Object
        :type com_obj: object
        """

        self.window_text = window_text
        self.com_obj = com_obj

    def is_foreground(self):
        """Returns True if window that contains desired string in titlebar is in foreground,
        else returns False

        :returns: True if window is in foreground, else False
        :rtype: bool
        """

        win_text = win32gui.GetWindowText(win32gui.GetForegroundWindow())
        if self.window_text in win_text:
            return True
        else:
            return False

    def close_callback(self, hwnd, lParam):
        """Closes window that contains desired string in titlebar
        by terminating associated process

        :param hwnd: Window handle
        :type hwnd: int
        :param lParam: second parameter passed in by win32gui.EnumWindows()
        :type lParam: object
        """
        if win32gui.IsWindowVisible(hwnd):
            if self.window_text in win32gui.GetWindowText(hwnd):
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                p = psutil.Process(pid)
                p.terminate()

    def close(self):
        """Enumerates all top-level windows on screen by passing handle to callback
        function that closes window
        """
        win32gui.EnumWindows(self.close_callback, None)

    def foreground_callback(self, hwnd, lParam):
        """Maximizes, activates, and brings window that contains desired string in titlebar
        to foreground

        :param hwnd: Window handle
        :type hwnd: int
        :param lParam: second parameter passed in by win32gui.EnumWindows()
        :type lParam: object
        """
        while True:
            if win32gui.IsWindowVisible(hwnd):
                if self.window_text in win32gui.GetWindowText(hwnd):
                    while not self.is_foreground():
                        _, pid = win32process.GetWindowThreadProcessId(hwnd)
                        self.com_obj.SendKeys('%')
                        self.com_obj.AppActivate(pid)
                        win32gui.SetForegroundWindow(hwnd)
                        win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)

            return

    def to_foreground(self):
        """Enumerates all top-level windows on screen by passing handle to callback
        function that maximizes, activates, and brings window to foreground
        """
        win32gui.EnumWindows(self.foreground_callback, None)

    def send_keys(self, keys):
        """Maximizes, activates, and brings window to foreground,
        then sends keystrokes to window

        :param keys: keystrokes to be sent to window
        :type keys: str
        """

        self.to_foreground()
        self.com_obj.SendKeys(keys)
