import time
import win32con
import win32gui

from minitabexecute import numberget, quickdatanam, finishcheckstr, turnscaccheckstr, turnscachigeffcheckstr


class HandleAlertCla:
    def __init__(self, quickdatanam_, finishcheckstr_, turnscaccheckstr_, turnscachigeffcheckstr_):
        self.quickdatanam = quickdatanam_
        self.finishcheckstr = finishcheckstr_
        self.turnscaccheckstr = turnscaccheckstr_
        self.turnscachigeffcheckstr = turnscachigeffcheckstr_


    def handle_window(self, hwnd, extra):
        if win32gui.IsWindowVisible(hwnd):
            #print(win32gui.GetWindowText(hwnd))
            if 'Warning' in win32gui.GetWindowText(hwnd):
                win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
            if 'Error' in win32gui.GetWindowText(hwnd):
                win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
            if 'Motor-cad_13_1_10' in win32gui.GetWindowText(hwnd):
                win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)

    def handlealert(self):
        while 1:
            win32gui.EnumWindows(self.handle_window, None)
            time.sleep(10)
            finishcheck = self.numberget(self.quickdatanam, self.finishcheckstr)
            turnscaccheck = self.numberget(self.quickdatanam, self.turnscaccheckstr)
            turnscachigeffcheck = self.numberget(self.quickdatanam, self.turnscachigeffcheckstr)
            if turnscachigeffcheck > 0:
                if finishcheck == 1 and turnscachigeffcheck == 0:
                    break
            elif turnscaccheck > 0 and turnscachigeffcheck == 0:
                if finishcheck == 1 and turnscaccheck == 0:
                    break
            else:
                if finishcheck == 1:
                    break

    def numberget(self, quickdatanam_, finishcheckstr_):
        return numberget(quickdatanam_, finishcheckstr_)

    def handlealertonce(self):
        win32gui.EnumWindows(self.handle_window, None)


if __name__ == '__main__':
    HandleAlertCla(quickdatanam, finishcheckstr, turnscaccheckstr, turnscachigeffcheckstr).handlealert()
