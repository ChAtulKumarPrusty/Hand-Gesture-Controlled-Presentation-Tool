# 🖐️ Hand Gesture Controlled Presentation Tool

This project allows you to **control PowerPoint presentations using hand gestures** in real time, using your webcam and OpenCV. The presentation slides are converted into images, and you can navigate them, draw annotations, and undo drawings using intuitive hand gestures.

## 📌 Features

- 🔄 Convert PowerPoint (`.pptx`) slides into images automatically.
- 📷 Use your webcam for real-time hand gesture detection.
- 👆 Navigate slides with:
  - 👈 Left swipe: Go to the previous slide.
  - 👉 Right swipe: Go to the next slide.
- ✍️ Draw on slides using a single finger.
- ➖ Undo drawings with a three-finger gesture.
- 🖱️ Smooth cursor movement with two fingers.
- 🖼️ Fullscreen display of presentation slides.

## 📁 Folder Structure

Hand-Gesture-Presentation/ │ ├── Presentation/ # Folder where slide images are stored ├── main.py # Main Python script └── README.md # Project documentation


## 🚀 Getting Started

### ✅ Prerequisites

Make sure you have the following installed:

- Python 3.7+
- [OpenCV](https://pypi.org/project/opencv-python/)
- [cvzone]
- [NumPy](https://pypi.org/project/numpy/)
- [comtypes](https://pypi.org/project/comtypes/)
- Microsoft PowerPoint (Installed on Windows)

### 📦 Installation

1. Clone the repository:

git clone https://github.com/your-username/Hand-Gesture-Controlled-Presentation-Tool.git
cd hand-gesture-presentation

Install the required packages:
pip install -r requirements.txt

Create a folder named Presentation (or use the GUI when you run the script).

🧠 How It Works
The script opens a GUI to let you select a .pptx file.

All slides are converted into .png images.
A webcam feed runs in parallel, detecting your hand using cvzone.HandTrackingModule.
Based on the number and position of fingers raised, different actions are triggered.

✋ Gesture Controls
Gesture	Action
👆 (1 finger up)	Draw on slide
✌️ (2 fingers up)	Move cursor
🤟 (3 fingers up)	Undo last annotation
👉 (4 fingers up)	Next slide
👈 (Thumb only)	Previous slide

⚠️ Notes
This only works on Windows systems due to comtypes and PowerPoint COM automation.
Make sure your hand is visible and well-lit for accurate gesture detection.

🙌 Acknowledgements
cvzone - For hand tracking.
OpenCV - For image processing.
Microsoft COM API - For PowerPoint automation.

👨‍💻 Author
Developed by Ch Atul Kumar Prusty

---

If you'd like me to also generate the `requirements.txt` file or a sample demo image/GIF banner, I can help with that too. Just let me know!
