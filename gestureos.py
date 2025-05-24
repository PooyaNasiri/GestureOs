import ctypes
import math
import time
import cv2
import mediapipe as mp
import win32com.client
import win32api
import win32con
import win32gui
import pyautogui
from pywinauto import Desktop

# Setup
pyautogui.FAILSAFE = False
screen_w, screen_h = pyautogui.size()
run_in_background = False
shell = win32com.client.Dispatch("Shell.Application")
cap = cv2.VideoCapture(0)
cap.set(cv2.CAP_PROP_FRAME_WIDTH, screen_w)
cap.set(cv2.CAP_PROP_FRAME_HEIGHT, screen_h)
cap.set(cv2.CAP_PROP_FPS, 60)

bend_angle = 70
unbend_angle = 150
delay = 1.0
gesture_timers = {}

mp_hands = mp.solutions.hands
hands = mp_hands.Hands(static_image_mode=False, max_num_hands=1, min_detection_confidence=0.8,
                       min_tracking_confidence=0.8)
mp_draw = mp.solutions.drawing_utils


def calculate_angle(a, b, c):
    ba = (a[0] - b[0], a[1] - b[1])
    bc = (c[0] - b[0], c[1] - b[1])
    dot = ba[0] * bc[0] + ba[1] * bc[1]
    mag_ba = math.hypot(*ba)
    mag_bc = math.hypot(*bc)
    if mag_ba == 0 or mag_bc == 0:
        return 0
    cos_angle = max(-1, min(1, dot / (mag_ba * mag_bc)))
    return math.degrees(math.acos(cos_angle))

def dist(a, b):
    return math.hypot(a.x - b.x, a.y - b.y)


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
        win32api.keybd_event(win32con.VK_MEDIA_PLAY_PAUSE, 0, win32con.KEYEVENTF_KEYUP, 0)

    @staticmethod
    def leftclick():
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        # pyautogui.click(button="left")

    @staticmethod
    def rightclick():
        pyautogui.click(button="right")

    @staticmethod
    def mute_toggle():
        pyautogui.press("volumemute")

    @staticmethod
    def volume_up():
        win32api.keybd_event(win32con.VK_VOLUME_UP, 0, 0, 0)
        win32api.keybd_event(win32con.VK_VOLUME_UP, 0, win32con.KEYEVENTF_KEYUP, 0)

    @staticmethod
    def volume_down():
        win32api.keybd_event(win32con.VK_VOLUME_DOWN, 0, 0, 0)
        win32api.keybd_event(win32con.VK_VOLUME_DOWN, 0, win32con.KEYEVENTF_KEYUP, 0)

    @staticmethod
    def toggle_maximize():
        hwnd = win32gui.GetForegroundWindow()
        placement = win32gui.GetWindowPlacement(hwnd)
        win32gui.ShowWindow(hwnd,
                            win32con.SW_RESTORE if placement[1] == win32con.SW_SHOWMAXIMIZED else win32con.SW_MAXIMIZE)

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


# Gesture detection loop
while True:
    ret, frame = cap.read()
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
            distance = dist(landmarks[0], landmarks[4]) / 9.0
            gesture_state = "Idle"

            fingers = {
                "Index": [5, 6, 8],
                "Middle": [9, 10, 12],
                "Ring": [13, 14, 16],
                "Pinky": [17, 18, 20],
            }

            bent_fingers = {
                name: landmarks[ids[2]].y > landmarks[ids[1]].y
                for name, ids in fingers.items()
            }

            # Gesture definitions
            y_values = [lm.y for i, lm in enumerate(landmarks) if i != 4]
            is_thumbs_up = landmarks[3].y <= min(y_values)
            is_thumbs_down = landmarks[3].y >= max(y_values)

            is_thumb_close = (min(landmarks[8].x, landmarks[17].x)
                              <= landmarks[4].x <= max(landmarks[8].x, landmarks[17].x))
            is_palm = all(not bent for bent in bent_fingers.values())
            is_full_palm = is_palm and not is_thumb_close
            is_semi_palm = is_palm and is_thumb_close
            is_fist = all(bent_fingers.values()) and is_thumb_close and not is_thumbs_up and not is_thumbs_down
            is_peace = not bent_fingers["Index"] and not bent_fingers["Middle"] and bent_fingers["Ring"] and \
                       bent_fingers["Pinky"]
            is_index = not bent_fingers["Index"] and all(bent_fingers[f] for f in ["Middle", "Ring", "Pinky"])
            is_spiderman = not bent_fingers["Index"] and bent_fingers["Middle"] and bent_fingers["Ring"] and not \
                bent_fingers["Pinky"]
            is_fuck = bent_fingers["Index"] and not bent_fingers["Middle"] and bent_fingers["Ring"] and bent_fingers[
                "Pinky"]
            is_pinky = not bent_fingers["Pinky"] and all(bent_fingers[f] for f in ["Middle", "Ring", "Index"])

            # Gesture-to-function map
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

            # Detect gesture from map
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

                # Overlay ROI box
                roi_px1 = int(roi_x * w)
                roi_py1 = int(roi_y * h)
                roi_px2 = int((1 - roi_x) * w)
                roi_py2 = int((1 - roi_y) * h)
                cv2.rectangle(frame, (roi_px1, roi_py1), (roi_px2, roi_py2), (0, 255, 0), 1)

                if roi_x <= ix <= 1 - roi_x and roi_y <= iy <= 1 - roi_y:
                    rel_x = (ix - roi_x) / roi_ratio
                    rel_y = (iy - roi_y) / roi_ratio
                    win32api.SetCursorPos((int(rel_x * screen_w), int(rel_y * screen_h)))

                if dist(landmarks[4], landmarks[8]) < distance:
                    gesture_state = "Right Click!"
                    function = Functions.rightclick
                elif dist(landmarks[4], landmarks[6]) < distance:
                    gesture_state = "Left Click!"
                    function = Functions.leftclick
                else:
                    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
            else:
                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)

            # Call function with cooldown
            if function:
                now = time.time()
                last_time = gesture_timers.get(gesture_state, 0)
                if now - last_time > delay:
                    gesture_timers[gesture_state] = now
                    function()

    if not run_in_background:
        cv2.putText(frame, gesture_state, (10, 40), cv2.FONT_HERSHEY_SIMPLEX, 1.0, (0, 255, 0), 2)
        cv2.imshow("GestureOS", frame)
        if cv2.waitKey(1) == 27:
            break

cap.release()
cv2.destroyAllWindows()
