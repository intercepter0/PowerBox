from pyautogui import size, position, FAILSAFE, hotkey, keyDown, keyUp, press
from os import system, path, popen, _exit
from pynput import mouse
from ctypes import cast, POINTER, windll
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
FAILSAFE = False


def get_width():
    return size().width


def get_height():
    return size().height


def mouse_on_right():
    return position().x == size().width - 1


def mouse_on_left():
    return position().x == 0


def mouse_on_top():
    return position().y <= 8


def mouse_on_bottom():
    return position().y >= size().height - 9


p1 = lambda: mouse_on_top() and mouse_on_left()
p2 = lambda: mouse_on_top()
p3 = lambda: mouse_on_top() and mouse_on_right()
p4 = lambda: mouse_on_left()
p5 = lambda: False
p6 = lambda: mouse_on_right()
p7 = lambda: mouse_on_bottom() and mouse_on_left()
p8 = lambda: mouse_on_bottom()
p9 = lambda: mouse_on_bottom() and mouse_on_right()

# Init audio utils
devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(
    IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))


def change_volume(dy):
    try:
        if dy < 0:
            volume.SetMasterVolumeLevel(volume.GetMasterVolumeLevel() - 1, None)
        else:
            volume.SetMasterVolumeLevel(volume.GetMasterVolumeLevel() + 1, None)
    except Exception as e:
        pass


def switch_desktop(side):
    hotkey('win', 'ctrl', 'right' if side == 1 else 'left')


def exit():
    _exit(0)


def open_powerline():
    powerline_path = path.realpath(__file__).replace("core.py", "") + "Powerline.exe"
    system(powerline_path)


def lock_session():
    windll.user32.LockWorkStation()


switch_apps_mode = False


def switch_apps(dy):
    global switch_apps_mode
    keyDown('alt')
    if dy < 0:
        if not switch_apps_mode:
            press('tab')
            press('left')
        press('left')
    else:
        press('tab')
    switch_apps_mode = True


def on_move(x, y):

    global switch_apps_mode
    if switch_apps_mode:
        if not mouse_on_top():
            keyUp('alt')
            switch_apps_mode = False


AVAILABLE_ACTIONS = {
    'change_volume': change_volume,
    'switch_desktop': switch_desktop,
    'start_powerline': open_powerline,
    'lock_session': lock_session,
    'switch_apps': switch_apps,
    'exit': exit,
    'none': lambda dy = 0: None
}


mouse_positions = [[[p1], [p2], [p3]],
                   [[p4], [p5], [p6]],
                   [[p7], [p8], [p9]]]


actions_on_scroll =      ['none',                 'switch_apps',                    'none',
                          'change_volume',        'none',                           'none',
                          'none',                 'switch_desktop',                 'none']


actions_on_wheel_press = ['none',                 'none',                    'exit',
                          'none',                 'none',                    'none',
                          'none',                 'lock_session',            'start_powerline']


def on_click(x, y, button, pressed):
    print('{0} at {1}'.format(
        ('Pressed' if pressed else 'Released')+" "+str(button),
        (x, y)))
    if pressed and str(button) == 'Button.middle':
        if p1():
            AVAILABLE_ACTIONS[actions_on_wheel_press[0]]()
        elif p3():
            AVAILABLE_ACTIONS[actions_on_wheel_press[2]]()
        elif p7():
            AVAILABLE_ACTIONS[actions_on_wheel_press[6]]()
        elif p9():
            AVAILABLE_ACTIONS[actions_on_wheel_press[8]]()
        elif p2():
            AVAILABLE_ACTIONS[actions_on_wheel_press[1]]()
        elif p4():
            AVAILABLE_ACTIONS[actions_on_wheel_press[3]]()
        elif p6():
            AVAILABLE_ACTIONS[actions_on_wheel_press[5]]()
        elif p8():
            AVAILABLE_ACTIONS[actions_on_wheel_press[7]]()


def on_scroll(x, y, dx, dy):
    #print('Scrolled {0} at {1}'.format(
    #    'down' if dy < 0 else 'up',
     #   (x, y)))
    if p1():
        AVAILABLE_ACTIONS[actions_on_scroll[0]](dy)
    elif p3():
        AVAILABLE_ACTIONS[actions_on_scroll[2]](dy)
    elif p7():
        AVAILABLE_ACTIONS[actions_on_scroll[6]](dy)
    elif p9():
        AVAILABLE_ACTIONS[actions_on_scroll[8]](dy)
    elif p2():
        AVAILABLE_ACTIONS[actions_on_scroll[1]](dy)
    elif p4():
        AVAILABLE_ACTIONS[actions_on_scroll[3]](dy)
    elif p6():
        AVAILABLE_ACTIONS[actions_on_scroll[5]](dy)
    elif p8():
        AVAILABLE_ACTIONS[actions_on_scroll[7]](dy)



# Collect events until released
with mouse.Listener(
    on_click=on_click,
    on_move=on_move,
    on_scroll=on_scroll) as listener:
    listener.join()

