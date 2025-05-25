import ctypes
import math
import time

import cv2
import mediapipe as mp
import numpy as np
import pyautogui
import win32api
import win32com.client
import win32con
import win32gui
from PIL import Image, ImageDraw, ImageFont
from pywinauto import Desktop


###############################################
#               SYSTEM ACTIONS               #
###############################################

class Functions:
    @staticmethod
    def sleep():
        ctypes.windll.powrprof.SetSuspendState(False, True, False)

    @staticmethod
    def showdesktop():
        pyautogui.hotkey("win", "d")

    @staticmethod
    def playpause():
        win32api.keybd_event(win32con.VK_MEDIA_PLAY_PAUSE, 0, 0, 0)
        win32api.keybd_event(
            win32con.VK_MEDIA_PLAY_PAUSE, 0, win32con.KEYEVENTF_KEYUP, 0
        )

    @staticmethod
    def leftclick():
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)

    @staticmethod
    def rightclick():
        pyautogui.click(button="right")

    @staticmethod
    def mute_toggle():
        pyautogui.press("volumemute")

    @staticmethod
    def volume_up():
        win32api.keybd_event(win32con.VK_VOLUME_UP, 0, 0, 0)
        win32api.keybd_event(
            win32con.VK_VOLUME_UP, 0, win32con.KEYEVENTF_KEYUP, 0
        )

    @staticmethod
    def volume_down():
        win32api.keybd_event(win32con.VK_VOLUME_DOWN, 0, 0, 0)
        win32api.keybd_event(
            win32con.VK_VOLUME_DOWN, 0, win32con.KEYEVENTF_KEYUP, 0
        )

    @staticmethod
    def toggle_maximize():
        hwnd = win32gui.GetForegroundWindow()
        placement = win32gui.GetWindowPlacement(hwnd)
        win32gui.ShowWindow(
            hwnd,
            win32con.SW_RESTORE
            if placement[1] == win32con.SW_SHOWMAXIMIZED
            else win32con.SW_MAXIMIZE,
        )

    window_list = []
    current_index = -1

    @staticmethod
    def switchtab():
        try:
            windows = Desktop(backend="uia").windows()
            visible_windows = [w for w in windows if w.is_visible() and w.window_text().strip()]
            visible_windows.sort(key=lambda w: w.window_text().lower())
            if Functions.window_list != visible_windows:
                Functions.window_list = visible_windows
                Functions.current_index = 0
            else:
                Functions.current_index = (Functions.current_index + 1) % len(Functions.window_list)
            next_win = Functions.window_list[Functions.current_index]
            next_win.set_focus()
        except Exception:
            pass

    @staticmethod
    def exit():
        exit(0)


###############################################
#         DRAWING EMOJI DIRECTLY ON FRAME     #
###############################################

# Preferred emoji-capable font (Windows 10+)
_EMOJI_FONT_PATHS = [
    r"C:\\Windows\\Fonts\\seguiemj.ttf",  # main Segoe UI Emoji
    r"/usr/share/fonts/truetype/noto/NotoColorEmoji.ttf",  # Linux
]


def _load_emoji_font(size: int = 48):
    """Return the first emoji-capable font we can open, else default."""
    for p in _EMOJI_FONT_PATHS:
        try:
            return ImageFont.truetype(p, size)
        except (OSError, IOError):
            continue
    return ImageFont.load_default()


_EMOJI_FONT_CACHE = {}


def draw_emoji(frame: np.ndarray, emoji_char: str, x: int, y: int, size: int = 48):
    """Draw a single Unicode emoji onto `frame` (BGR) at pixel (x, y)."""
    if len(emoji_char) == 0:
        return
    key = size
    font = _EMOJI_FONT_CACHE.get(key)
    if font is None:
        font = _load_emoji_font(size)
        _EMOJI_FONT_CACHE[key] = font

    # Convert region of interest to PIL for text rendering
    img_pil = Image.fromarray(frame)
    draw = ImageDraw.Draw(img_pil)
    draw.text((x, y), emoji_char, font=font, embedded_color=True)
    frame[:] = np.array(img_pil)


EMOJI_MAP = {
    "Volume Up": "🔊",
    "Volume Down": "🔉",
    "Play/Pause": "⏯️",
    "Mute Toggle": "🔇",
    "Show Desktop": "🖥️",
    "Maximize/Restore": "🗖",
    "Switch Tabs": "🗂️",
    "Mouse Control": "🖱️"
}


###############################################
#           Windows Frame Settings            #
###############################################

def dist(a, b):
    return math.hypot(a.x - b.x, a.y - b.y)


def disable_maximize_button():
    hwnd = win32gui.FindWindow(None, window_name)
    style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
    style &= ~win32con.WS_MAXIMIZEBOX  # Disable maximize button
    style &= ~win32con.WS_THICKFRAME  # Disable resizing
    win32gui.SetWindowLong(hwnd, win32con.GWL_STYLE, style)
    win32gui.SetWindowPos(hwnd, None, 0, 0, 0, 0,
                          win32con.SWP_NOMOVE |
                          win32con.SWP_NOSIZE |
                          win32con.SWP_NOZORDER |
                          win32con.SWP_FRAMECHANGED)


def set_window_icon(icon_path):
    hwnd = win32gui.FindWindow(None, window_name)
    if hwnd == 0:
        print(f"Window '{window_name}' not found.")
        return

    hicon = win32gui.LoadImage(
        None,
        icon_path,
        win32con.IMAGE_ICON,
        0, 0,
        win32con.LR_LOADFROMFILE | win32con.LR_DEFAULTSIZE
    )

    if hicon == 0:
        print("Failed to load icon.")
        return

    win32gui.SendMessage(hwnd, win32con.WM_SETICON, win32con.ICON_SMALL, hicon)
    win32gui.SendMessage(hwnd, win32con.WM_SETICON, win32con.ICON_BIG, hicon)


###############################################
#               MAIN GESTURE LOOP             #
###############################################

def gesture_loop():
    if not run_in_background:
        cv2.namedWindow(window_name, cv2.WINDOW_NORMAL)
        disable_maximize_button()
        set_window_icon("icon.ico")
    while True:
        ret, frame = cap.read()
        if not ret:
            continue

        frame = cv2.flip(frame, 1)
        rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        result = hands.process(rgb)
        gesture_state = "No hand"
        function = None

        if result.multi_hand_landmarks:
            for hand_landmarks in result.multi_hand_landmarks:
                mp_draw.draw_landmarks(frame, hand_landmarks, mp_hands.HAND_CONNECTIONS)
                landmarks = hand_landmarks.landmark
                h, w, _ = frame.shape
                distance_thresh = dist(landmarks[0], landmarks[4]) / 9.0
                gesture_state = "Idle"

                # == Calculate finger states ==
                fingers = {
                    "Index": [5, 6, 8],
                    "Middle": [9, 10, 12],
                    "Ring": [13, 14, 16],
                    "Pinky": [17, 18, 20],
                }
                bent_fingers = {
                    name: landmarks[ids[2]].y > landmarks[ids[1]].y for name, ids in fingers.items()
                }

                # == Basic pose predicates ==
                y_values = [lm.y for i, lm in enumerate(landmarks) if i != 4]
                is_thumbs_up = landmarks[3].y <= min(y_values)
                is_thumbs_down = landmarks[3].y >= max(y_values)
                is_thumb_close = min(landmarks[8].x, landmarks[17].x) <= landmarks[4].x <= max(landmarks[8].x,
                                                                                               landmarks[17].x)
                is_palm = all(not bent for bent in bent_fingers.values())
                is_full_palm = is_palm and not is_thumb_close
                is_semi_palm = is_palm and is_thumb_close
                is_fist = all(bent_fingers.values()) and is_thumb_close and not is_thumbs_up and not is_thumbs_down
                is_peace = not bent_fingers["Index"] and not bent_fingers["Middle"] and bent_fingers["Ring"] and \
                           bent_fingers["Pinky"]
                is_index = not bent_fingers["Index"] and all(bent_fingers[f] for f in ["Middle", "Ring", "Pinky"])
                is_spiderman = not bent_fingers["Index"] and bent_fingers["Middle"] and bent_fingers["Ring"] and not \
                    bent_fingers["Pinky"]
                is_fuck = bent_fingers["Index"] and not bent_fingers["Middle"] and bent_fingers["Ring"] and \
                          bent_fingers["Pinky"]
                is_pinky = not bent_fingers["Pinky"] and all(bent_fingers[f] for f in ["Middle", "Ring", "Index"])

                # == Map gesture → function ==
                gesture_map = {
                    "Volume Up": (is_thumbs_up, Functions.volume_up),
                    "Volume Down": (is_thumbs_down, Functions.volume_down),
                    "Exit": (is_fuck, Functions.exit),
                    "Switch Tabs": (is_semi_palm, Functions.switchtab),
                    "Show Desktop": (is_fist, Functions.showdesktop),
                    "Maximize/Restore": (is_spiderman, Functions.toggle_maximize),
                    "Mute Toggle": (is_peace, Functions.mute_toggle),
                    "Play/Pause": (is_full_palm, Functions.playpause),
                    "Pinky": (is_pinky, None),
                }

                # Detect gesture and associated action
                for gesture, (condition, action) in gesture_map.items():
                    if condition:
                        gesture_state = gesture
                        function = action
                        break

                # Special case: Mouse Control
                if is_index and function is None:
                    gesture_state = "Mouse Control"
                    roi_ratio = 0.8
                    roi_x, roi_y = (1 - roi_ratio) / 2, (1 - roi_ratio) / 2
                    ix, iy = landmarks[8].x, landmarks[8].y

                    # Draw ROI
                    roi_px1 = int(roi_x * w)
                    roi_py1 = int(roi_y * h)
                    roi_px2 = int((1 - roi_x) * w)
                    roi_py2 = int((1 - roi_y) * h)
                    cv2.rectangle(frame, (roi_px1, roi_py1), (roi_px2, roi_py2), (0, 255, 0), 1)

                    cursor_dist_thumb = dist(landmarks[4], landmarks[8])
                    if cursor_dist_thumb < distance_thresh:
                        gesture_state = "Right Click!"
                        function = Functions.rightclick
                    elif dist(landmarks[4], landmarks[6]) < distance_thresh:
                        gesture_state = "Left Click!"
                        function = Functions.leftclick
                    else:
                        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)

                    # Mouse mapping
                    if roi_x <= ix <= 1 - roi_x and roi_y <= iy <= 1 - roi_y:
                        rel_x = (ix - roi_x) / roi_ratio
                        rel_y = (iy - roi_y) / roi_ratio
                        win32api.SetCursorPos((int(rel_x * screen_w), int(rel_y * screen_h)))
                        thickness = -1 if (gesture_state == "Right Click!" or gesture_state == "Left Click!") else 2
                        cv2.circle(frame, (int(ix * w), int(iy * h)), int(cursor_dist_thumb * 200.0), (0, 255, 255), thickness)
                else:
                    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)

                if function:
                    now = time.time()
                    last_time = gesture_timers.get(gesture_state, 0)
                    if now - last_time > delay:
                        gesture_timers[gesture_state] = now
                        function()

                if gesture_state in EMOJI_MAP:
                    wrist = landmarks[0]
                    wx, wy = int(wrist.x * w), int(wrist.y * h)
                    draw_emoji(frame, EMOJI_MAP[gesture_state], wx + 20, wy - 20, size=48)

        if not run_in_background:
            cv2.putText(frame, gesture_state, (10, 40), cv2.FONT_HERSHEY_SIMPLEX, 1.0, (0, 255, 0), 2)
            cv2.imshow(window_name, frame)
        if cv2.waitKey(1) == 27:
            break

    cap.release()
    cv2.destroyAllWindows()


###############################################
#                    MAIN                     #
###############################################

if __name__ == "__main__":
    window_name = "GestureOS"
    run_in_background = False
    pyautogui.FAILSAFE = False
    screen_w, screen_h = pyautogui.size()
    shell = win32com.client.Dispatch("Shell.Application")

    cap = cv2.VideoCapture(0)
    cap.set(cv2.CAP_PROP_FRAME_WIDTH, screen_w)
    cap.set(cv2.CAP_PROP_FRAME_HEIGHT, screen_h)
    cap.set(cv2.CAP_PROP_FPS, 60)

    delay = 1.0  # gesture cooldown in seconds
    gesture_timers = {}

    mp_hands = mp.solutions.hands
    hands = mp_hands.Hands(
        static_image_mode=False,
        max_num_hands=1,
        min_detection_confidence=0.8,
        min_tracking_confidence=0.8,
    )
    mp_draw = mp.solutions.drawing_utils

    gesture_loop()
