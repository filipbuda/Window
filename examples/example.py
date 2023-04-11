from window import Window
import win32com.client as com_client
import os
import time

ONE_SECOND = 1

if __name__ == "__main__":
    # open untitled notepad file
    os.startfile("notepad.exe")

    # 1 second delay
    time.sleep(ONE_SECOND)

    # create COM object
    com_obj = com_client.Dispatch("WScript.Shell")

    # create object for notepad window
    notepad_window = Window("Untitled", com_obj)

    # maximize, activate, and bring notepad window to foreground
    notepad_window.to_foreground()

    # 1 second delay
    time.sleep(ONE_SECOND)

    # send keystrokes to notepad window
    notepad_window.send_keys("foobar")

    # 1 second delay
    time.sleep(ONE_SECOND)

    # close notepad window
    notepad_window.close()
