import win32api, win32gui, win32print
import win32con


def get_real_resolution():
    """获取真实的分辨率"""
    hDC = win32gui.GetDC(0)
    w = win32print.GetDeviceCaps(hDC, win32con.DESKTOPHORZRES)
    h = win32print.GetDeviceCaps(hDC, win32con.DESKTOPVERTRES)
    return w, h


def get_screen_size():
    """获取缩放后的分辨率"""
    w = win32api.GetSystemMetrics(0)
    h = win32api.GetSystemMetrics(1)
    return w, h

def get_dpi():
    """ 获取缩放比 """
    real_width,real_height = get_real_resolution()
    virtual_width,virtual_height = get_screen_size()

    return real_width/virtual_width