import cv2
import mediapipe as mp
import math
import win32api
import win32con
import win32gui
import win32com.client
import pyautogui
import time
import os

pyautogui.FAILSAFE = False
run_in_background = False
screen_w, screen_h = pyautogui.size()
shell = win32com.client.Dispatch("Shell.Application")
mp_hands = mp.solutions.hands
hands = mp_hands.Hands(
    static_image_mode=False,
    max_num_hands=1,
    min_detection_confidence=0.9,
    min_tracking_confidence=0.9,
)
mp_draw = mp.solutions.drawing_utils
gesture_timers = {}

cap = cv2.VideoCapture(0)
bend_angle = 70
unbend_angle = 150
delay = 1.0


def calculate_angle(a, b, c):
    ba = (a[0] - b[0], a[1] - b[1])
    bc = (c[0] - b[0], c[1] - b[1])
    dot_product = ba[0] * bc[0] + ba[1] * bc[1]
    magnitude_ba = math.hypot(ba[0], ba[1])
    magnitude_bc = math.hypot(bc[0], bc[1])
    if magnitude_ba == 0 or magnitude_bc == 0:
        return 0
    cos_angle = dot_product / (magnitude_ba * magnitude_bc)
    cos_angle = max(-1, min(1, cos_angle))
    return math.degrees(math.acos(cos_angle))


def dist(a, b):
    return ((a.x - b.x) ** 2 + (a.y - b.y) ** 2) ** 0.5

class Functions:
    @staticmethod
    def sleep():
        os.system(
            "powershell -command \"Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.Application]::SetSuspendState('Suspend', $false, $false)\""
        )

    @staticmethod
    def showdesktop():
        pyautogui.hotkey("win", "d")


    @staticmethod
    def switchtab():
        pyautogui.hotkey("alt", "tab")

    @staticmethod
    def playpause():
        VK_MEDIA_PLAY_PAUSE = 0xB3
        win32api.keybd_event(VK_MEDIA_PLAY_PAUSE, 0, 0, 0)
        time.sleep(0.05)
        win32api.keybd_event(VK_MEDIA_PLAY_PAUSE, 0, win32con.KEYEVENTF_KEYUP, 0)

    @staticmethod
    def leftclick():
        pyautogui.click(button="left")

    @staticmethod
    def rightclick():
        pyautogui.click(button="right")

    @staticmethod
    def mute_toggle():
        pyautogui.press("volumemute")

    @staticmethod
    def toggle_maximize():
        hwnd = win32gui.GetForegroundWindow()
        placement = win32gui.GetWindowPlacement(hwnd)
        is_maximized = placement[1] == win32con.SW_SHOWMAXIMIZED

        if is_maximized:
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        else:
            win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
    @staticmethod
    def exit():
        exit(0)


while True:
    ret, frame = cap.read()
    if not ret:
        break

    frame = cv2.flip(frame, 1)
    rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
    result = hands.process(rgb)
    gesture_state = ""
    if result.multi_hand_landmarks:
        for hand_landmarks in result.multi_hand_landmarks:
            mp_draw.draw_landmarks(frame, hand_landmarks, mp_hands.HAND_CONNECTIONS)
            landmarks = hand_landmarks.landmark
            distance = dist(landmarks[0], landmarks[4]) / 6.0
            h, w, _ = frame.shape
            fingers = {
                "Index": [5, 6, 8],
                "Middle": [9, 10, 12],
                "Ring": [13, 14, 16],
                "Pinky": [17, 18, 20],
            }
            bent_fingers = {}
            for name, ids in fingers.items():
                joint_y = landmarks[ids[1]].y
                tip_y = landmarks[ids[2]].y
                bent_fingers[name] = tip_y > joint_y
            is_thumb_close = min(landmarks[8].x, landmarks[17].x) <= landmarks[4].x <= max(landmarks[8].x,
                                                                                           landmarks[17].x)
            palm = all(not bent for bent in bent_fingers.values())
            is_full_palm = palm and not is_thumb_close
            is_semi_palm = palm and is_thumb_close
            is_fist = all(bent_fingers.values())
            is_peace = (
                    not bent_fingers["Index"] and
                    not bent_fingers["Middle"] and
                    bent_fingers["Ring"] and
                    bent_fingers["Pinky"]
            )
            is_index = (
                    not bent_fingers["Index"] and
                    bent_fingers["Middle"] and
                    bent_fingers["Ring"] and
                    bent_fingers["Pinky"]
            )
            is_spiderman = (
                    not bent_fingers["Index"] and
                    bent_fingers["Middle"] and
                    bent_fingers["Ring"] and
                    not bent_fingers["Pinky"]
            )
            is_fuck = (
                    bent_fingers["Index"] and
                    not bent_fingers["Middle"] and
                    bent_fingers["Ring"] and
                    bent_fingers["Pinky"]
            )

            function = None
            gesture_state = 'Idle'

            if is_fuck:
                gesture_state = 'Fuck'
                function = Functions.exit

            elif is_semi_palm:
                gesture_state = "Semi Palm"
                function = Functions.switchtab

            elif is_fist:
                gesture_state = "Fist"
                function = Functions.showdesktop

            elif is_spiderman:
                gesture_state = "SpiderMan"
                function = Functions.toggle_maximize

            elif is_peace:
                gesture_state = "Peace"
                function = Functions.mute_toggle

            elif is_full_palm:
                gesture_state = "Full Palm"
                function = Functions.playpause

            elif is_index:
                gesture_state = "Mouse Control"
                roi_ratio = 0.8
                roi_x, roi_y = (1 - roi_ratio) / 2, (1 - roi_ratio) / 2
                ix, iy = landmarks[8].x, landmarks[8].y

                if roi_x <= ix <= 1 - roi_x and roi_y <= iy <= 1 - roi_y:
                    rel_x = (ix - roi_x) / roi_ratio
                    rel_y = (iy - roi_y) / roi_ratio
                    pyautogui.moveTo(int(rel_x * screen_w), int(rel_y * screen_h))

                if dist(landmarks[4], landmarks[8]) < distance:
                    function = Functions.rightclick
                    gesture_state = "Right Click!"

                elif dist(landmarks[4], landmarks[6]) < distance:
                    function = Functions.leftclick
                    gesture_state = "Left Click!"

            if function:
                now = time.time()
                last_time = gesture_timers.get(gesture_state, 0)
                if now - last_time > delay:
                    gesture_timers[gesture_state] = now
                    function()

    if not run_in_background:
        cv2.putText(frame, gesture_state, (10, 50), cv2.FONT_HERSHEY_SIMPLEX, 1.0, (50, 200, 50), 2)
        cv2.imshow("Hand Detection", frame)
        if cv2.waitKey(1) == 27:
            break

cap.release()
cv2.destroyAllWindows()
