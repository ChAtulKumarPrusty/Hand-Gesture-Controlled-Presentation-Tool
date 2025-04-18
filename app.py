import cv2
import os
import numpy as np
from cvzone.HandTrackingModule import HandDetector
from tkinter import Tk, filedialog
from comtypes.client import CreateObject

# Parameters
width, height = 1280, 720
gestureThreshold = 700  # Adjusted threshold for better detection
folderPath = os.path.abspath("Presentation")  # Absolute path to the folder


# Function to convert PPT to images
def ppt_to_images(ppt_path, folder_path):
    powerpoint = CreateObject("PowerPoint.Application")
    powerpoint.Visible = 1
    presentation = powerpoint.Presentations.Open(ppt_path)

    if not os.path.exists(folder_path):
        os.makedirs(folder_path)

    for i, slide in enumerate(presentation.Slides):
        img_path = os.path.join(folder_path, f"slide_{i + 1}.png")
        try:
            slide.Export(img_path, "PNG", width, height)
            print(f"Saved slide {i + 1} as image at {img_path}.")
        except Exception as e:
            print(f"Failed to save slide {i + 1}. Error: {e}")

    presentation.Close()
    powerpoint.Quit()


# GUI for choosing PPT file
Tk().withdraw()
ppt_file_path = filedialog.askopenfilename(title="Select PowerPoint File", filetypes=[("PPT Files", "*.pptx")])

if not ppt_file_path:
    print("No PowerPoint file selected.")
    exit(1)

# Convert PPT to images
ppt_to_images(ppt_file_path, folderPath)

# Get list of presentation images
pathImages = sorted(os.listdir(folderPath), key=len)
if not pathImages:
    print("No images were generated from the PowerPoint file.")
    exit(1)

# Camera Setup
cap = cv2.VideoCapture(0)
cap.set(3, width)
cap.set(4, height)

# Hand Detector
detectorHand = HandDetector(detectionCon=0.8, maxHands=1)

# Variables
imgList = []
delay = 30
buttonPressed = False
counter = 0
drawMode = False
imgNumber = 0
delayCounter = 0
annotations = [[]]
annotationNumber = -1
annotationStart = False

# Get list of presentation images
pathImages = sorted(os.listdir(folderPath), key=len)
print(pathImages)

# Initial cursor position
previousX, previousY = 0, 0

# Create a named window
cv2.namedWindow("Slides", cv2.WND_PROP_FULLSCREEN)
cv2.setWindowProperty("Slides", cv2.WND_PROP_FULLSCREEN, cv2.WINDOW_FULLSCREEN)

while True:
    # Get image frame
    success, img = cap.read()
    img = cv2.flip(img, 1)
    pathFullImage = os.path.join(folderPath, pathImages[imgNumber])
    imgCurrent = cv2.imread(pathFullImage)

    # Resize the slide to fit the full window
    imgCurrent = cv2.resize(imgCurrent, (width, height))

    # Find the hand and its landmarks
    hands, img = detectorHand.findHands(img)  # with draw

    if hands and buttonPressed is False:  # If hand is detected

        hand = hands[0]
        cx, cy = hand["center"]
        lmList = hand["lmList"]  # List of 21 Landmark points
        fingers = detectorHand.fingersUp(hand)  # List of which fingers are up

        # Constrain values for easier drawing
        xVal = int(np.interp(lmList[8][0], [0, width], [0, width]))
        yVal = int(np.interp(lmList[8][1], [0, height], [0, height]))
        indexFinger = xVal, yVal

        if cy <= gestureThreshold:  # If hand is at the height of the face
            if fingers == [1, 0, 0, 0, 0]:  # Left gesture
                print("Left")
                buttonPressed = True
                if imgNumber > 0:
                    imgNumber -= 1
                    annotations = [[]]
                    annotationNumber = -1
                    annotationStart = False
            if fingers == [0, 1, 1, 1, 1]:  # Right gesture
                print("Right")
                buttonPressed = True
                if imgNumber < len(pathImages) - 1:
                    imgNumber += 1
                    annotations = [[]]
                    annotationNumber = -1
                    annotationStart = False

        if fingers == [0, 1, 1, 0, 0]:  # Two-finger cursor movement

            # Smooth cursor movement
            smoothedX = int(previousX * 0.8 + xVal * 0.2)
            smoothedY = int(previousY * 0.8 + yVal * 0.2)
            indexFinger = smoothedX, smoothedY

            cv2.circle(imgCurrent, indexFinger, 12, (0, 0, 255), cv2.FILLED)

            previousX, previousY = smoothedX, smoothedY

        elif fingers == [0, 1, 0, 0, 0]:  # Single finger for drawing
            if annotationStart is False:
                annotationStart = True
                annotationNumber += 1
                annotations.append([])
            annotations[annotationNumber].append(indexFinger)
            cv2.circle(imgCurrent, indexFinger, 12, (0, 255, 255), cv2.FILLED)

        else:
            annotationStart = False

        if fingers == [0, 1, 1, 1, 0]:  # Undo gesture
            if annotations:
                annotations.pop(-1)
                annotationNumber -= 1
                buttonPressed = True

    else:
        annotationStart = False

    if buttonPressed:
        counter += 1
        if counter > delay:
            counter = 0
            buttonPressed = False

    for i, annotation in enumerate(annotations):
        for j in range(len(annotation)):
            if j != 0:
                cv2.line(imgCurrent, annotation[j - 1], annotation[j], (0, 0, 200), 12)

    cv2.imshow("Slides", imgCurrent)
    cv2.imshow("Image", img)

    key = cv2.waitKey(1)
    if key == ord('q'):
        break

cap.release()
cv2.destroyAllWindows()
