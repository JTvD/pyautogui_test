import pyautogui as autopy
import win32gui
import win32con
import win32process
import win32api
import psutil
import time
import math
from pynput.mouse import Listener


class mouse_recorder():
    """Record the mouse clicks to be used in the automation of applications"""
    def __init__(self):
        print('\nMouse recorder initialized\n')
        self.record_mouse_movement()

    def print_click_coords(self, x, y, button, pressed):
        """Print the click coordinates"""
        if pressed:
            print(f'{button},{(x, y)}')

    def record_mouse_movement(self):
        """Record the mouse movement"""
        with Listener(on_click=self.print_click_coords) as listener:
            listener.join()


def is_program_running(program: str):
    """Help function to check if the program is already running"""
    for proc in psutil.process_iter(['name']):
        if proc.info['name'] == program:
            return True
    return False


def start_program(program: str):
    """Starts the program"""
    autopy.keyDown('win')
    autopy.keyUp('win')
    time.sleep(1)
    autopy.write('Paint')
    autopy.press('enter')
    time.sleep(2)


def bring_paint_to_foreground(program: str):
    """Bring the program to the foreground"""
    def window_enum_callback(hwnd, pid):
        if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
            _, found_pid = win32process.GetWindowThreadProcessId(hwnd)
            if found_pid == pid:
                # Simulate pressing the 'ALT' key
                win32api.keybd_event(win32con.VK_MENU, 0)
                try:
                    # Restore the window to its original size and position before it was minimized or maximized
                    # win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                    win32gui.SetForegroundWindow(hwnd)
                finally:
                    # Simulate releasing the 'ALT' key
                    win32api.keybd_event(win32con.VK_MENU, 0, win32con.KEYEVENTF_KEYUP)

    for proc in psutil.process_iter(['pid', 'name']):
        if proc.info['name'] == program:
            win32gui.EnumWindows(window_enum_callback, proc.info['pid'])
            break


def main():
    """Test implementation, startin paint and drawing a square and a circle"""
    # record_mouse_movement()
    if not is_program_running('mspaint.exe'):
        start_program('Paint')
    else:
        bring_paint_to_foreground('mspaint.exe')
    time.sleep(1)

    command_list = [{'Button.left': (287, 89)},
                    {'Button.left': (968, 73)},
                    {'Button.left': (211, 265)}]

    for command in command_list:
        for action, coords in command.items():
            if 'left' in action:
                autopy.click(coords)
            else:
                print('Action not implemented: %s', action)
            time.sleep(1)

    # Draw a square
    autopy.moveTo(400, 400)  # Starting point of the circle
    autopy.dragTo(500, 400, button='left')  # Draw the top side of the circle
    autopy.dragTo(500, 500, button='left')  # Draw the right side of the circle
    autopy.dragTo(400, 500, button='left')  # Draw the bottom side of the circle
    autopy.dragTo(400, 400, button='left')  # Draw the left side of the circle

    # Center of the circle
    x, y = 500, 500
    # Radius of the circle
    r = 100

    # Move the mouse to the starting point
    autopy.moveTo(x + r, y)

    # Draw the circle
    for i in range(360):
        # Calculate the angle in radians
        angle = math.radians(i)
        # Calculate the new x, y coordinates
        new_x = x + r * math.cos(angle)
        new_y = y + r * math.sin(angle)
        # Drag the mouse to the new coordinates
        autopy.dragTo(new_x, new_y, duration=0.1)

    # Release the mouse button
    autopy.mouseUp()


if __name__ == '__main__':
    from threading import Thread
    main()
    Thread(target=main).start()
