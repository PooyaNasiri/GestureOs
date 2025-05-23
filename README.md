# 🖐️ GestureOS: Touchless Gesture-Controlled Desktop

![Alt text](/GestureOS.jpg)

Welcome to **GestureOS**, a futuristic Python-based system that lets you control your Windows PC using nothing but your **hand gestures** via your webcam. Developed with passion by [Pooya Nasiri](https://github.com/pooyanasiri), this project combines computer vision, system control, and media automation in a single script.

---

## 🚀 Features

- 🖱️ **Mouse Control** — Move your index finger to control the mouse cursor
- 👊 **Fist** — Instantly show the desktop
- ✋ **Full Palm** — Play/Pause media playback
- ✌️ **Peace** — Toggle mute
- 🤟 **Spiderman Gesture** — Maximize or restore current window
- 🤬 **Middle Finger (Easter Egg)** — Exit the app 😈
- 🖐️ **Semi-Palm** — Switch tabs (Alt+Tab)
- 🖱️💥 **Pinch Gestures** — Right and left clicks
- 🎯 **Region-based cursor** — Mouse maps only inside a central 80% region of the screen

---
## 🧠 How It Works

GestureOS uses **MediaPipe** to detect hand landmarks and classify gestures based on finger positions and joint angles for both hands but a single hand at a time. When a known gesture is recognized, it triggers corresponding actions using `pyautogui`, `win32api`, and Windows shell commands.

> ⚙️ **Configuration Notice:**  
> The system includes a flag named `run_in_background` which controls whether the app runs with a visible UI.  
> - If `run_in_background` is set to `True`, **no UI window will appear**, and gesture control will run silently.  
> - ⚠️ **Warning:** In background mode, the only gesture to **safely close the app** is the **middle finger gesture**, or you must manually terminate the process via Task Manager.

> 💬 **Want More Gestures or Custom Actions?**  
> If you're looking to expand GestureOS with more gesture types, advanced system control, or app-specific shortcuts [**contact me**](#author).

---

## 🛠️ Setup Instructions

### 1. Clone the repository

```bash
git clone https://github.com/pooyanasiri/gestureos.git
cd gestureos
```

### 2. Install the required libraries

```bash
pip install -r requirements.txt
```

### 3. Run the script

```bash
python gestureos.py
```

Make sure your **webcam is enabled** and you have a **good lighting** condition.

---

## 💡 Gesture Guide

| Gesture         | Action                         |
|----------------|--------------------------------|
| Fist ✊           | Show Desktop                   |
| Full Palm ✋      | Play/Pause Media               |
| Peace ✌️        | Mute/Unmute                    |
| Spiderman 🤟     | Maximize / Restore Window      |
| Semi-Palm       | Switch Tabs (Alt+Tab)          |
| Middle Finger🖕   | Exit Script                    |
| Index Only ☝️     | Mouse Movement                 |
| Pinch Thumb + Index | Right Click               |
| Pinch Thumb + Middle | Left Click               |

---

## 🧩 How to Customize

Want to add your own gestures? Modify this section in `gestureos.py`:

```python
if is_fist:
    function = Functions.showdesktop
elif is_peace:
    function = Functions.mute_toggle
```

You can map gestures to any `pyautogui`, `os.system`, or `win32gui` commands.

---

##  Author


**Pooya Nasiri**  
E-mail: [pooya.nasiri75@gmail.com](mailto:pooya.nasiri75@gmail.com)  
GitHub: [@pooyanasiri](https://github.com/pooyanasiri)  
LinkedIn: [Pooya Nasiri](https://www.linkedin.com/in/pooyanasiri/)

---

## 📄 License

MIT License — free to use, modify, and share.