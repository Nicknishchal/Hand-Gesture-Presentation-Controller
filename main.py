# import for GUI project
import tkinter as tk
from tkinter import filedialog,messagebox

# import for model
from cvzone.HandTrackingModule import HandDetector
import cv2
import os
import numpy as np

# Function to convert PowerPoint file to PNG images
from spire.presentation.common import *
from spire.presentation import *

def convertTopng(file_name,output_name):
    # Create a Presentation object
    presentation = Presentation()

    # Load a PowerPoint presentation
    presentation.LoadFromFile(file_name)

    # Loop through the slides in the presentation for conversion
    for i, slide in enumerate(presentation.Slides):
        # Specify the output file name

        fileName =output_name+"/" + str(i+1) + ".png"
        # Save each slide as a PNG image
        image = slide.SaveAsImageByWH(960,720)
        image.Save(fileName)
        image.Dispose()

    label_area.config(text="Presentation Converted")
    presentation.Dispose()

# Function to handle button click event
def browse_and_convert():
    try:
        label_area.config(text="Converting...")
        file_path = filedialog.askopenfilename(filetypes=[("PowerPoint files", "*.pptx;*.ppt")])
        if file_path:
            # Specify the output folder where PNG images will be saved
            output_folder = filedialog.askdirectory(title="Select Output Folder")
            if output_folder:
                convertTopng(file_path, output_folder)
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while accessing file dialog: {e}")

# hand gesture controller functions start
def hand_gesture_controller():
    width, height = 1280, 720
    gestureThreshold = 300
    folderPath = "img src"

    # Camera Setup
    cap = cv2.VideoCapture(0)
    cap.set(3, height)
    cap.set(4, width)

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
    annotationNumber = 0
    annotationStart = False
    hs, ws = int(120 * 1), int(213 * 0.8)  # width and height of small image

    # Get list of presentation images
    pathImages = sorted(os.listdir(folderPath), key=len)
    print(pathImages)
    label_area.config(text="Starting the Controller...")

    while True:
        # Get image frame
        success, img = cap.read()
        img = cv2.flip(img, 1)
        pathFullImage = os.path.join(folderPath, pathImages[imgNumber])
        imgCurrent = cv2.imread(pathFullImage)
        # Find the hand and its landmarks
        hands, img = detectorHand.findHands(img)  # with draw

        # Draw Gesture Threshold line
        cv2.line(img, (0, gestureThreshold), (width, gestureThreshold), (0, 255, 0), 10)
        print(annotationNumber)
        if hands and buttonPressed is False:  # If hand is detected
            hand = hands[0]
            cx, cy = hand["center"]
            lmList = hand["lmList"]  # List of 21 Landmark points
            fingers = detectorHand.fingersUp(hand)  # List of which fingers are up
            # xVal=int(np.interp(lmList[8][0],[300,width-300],[0,width]))
            # yVal=int(np.interp(lmList[8][1],[250,height-250],[0,height]))
            # indexFinger =xVal,yVal
            indexFinger=lmList[8][0], lmList[8][1]

            if cy <= gestureThreshold:  # If hand is at the height of the face
                # gesture 1 move to previous slide
                if fingers == [1, 0, 0, 0, 0]:
                    annotationStart = False
                    print("Left")
                    buttonPressed = True
                    if imgNumber > 0:
                        imgNumber -= 1
                        annotations = [[]]
                        annotationNumber = 0
                # gesture 2 move to next slide
                if fingers == [0, 0, 0, 0, 1]:
                    annotationStart = False
                    print("Right")
                    buttonPressed = True
                    if imgNumber < len(pathImages) - 1:
                        imgNumber += 1
                        annotations = [[]]
                        annotationNumber = 0
                # gesture 3 move to last slide directly
                if fingers == [0, 0, 0, 1, 1]:
                    annotationStart = False
                    print("Right Most")
                    buttonPressed = True
                    imgNumber =len(pathImages)-1
                    if annotations:
                        annotations = [[]]
                        annotationNumber = 0
                #gesture 4 erase complete from slides
                if fingers == [0, 1, 1, 1, 1]:
                    if annotations:
                        annotations = [[]]
                        annotationNumber = 0
                        annotationStart = False
                # gesture 5 erase last opeartion        
                if fingers == [0, 1, 1, 1, 0]:
                    if annotations:
                        if annotationNumber>-1:
                            annotations.pop(-1)
                            annotationNumber -= 1
                            buttonPressed = True

            # gesture 6 show pointer 
            if fingers == [0, 1, 1, 0, 0]:
                cv2.circle(imgCurrent, indexFinger, 10, (0, 0, 255), cv2.FILLED)
            #  gesture 7 draw with pointer
            if fingers == [0, 1, 0, 0, 0]:
                if annotationStart is False:
                    annotationStart = True
                    annotationNumber += 1
                    annotations.append([])
                print(annotationNumber)
                cv2.circle(imgCurrent, indexFinger, 10, (0, 0, 255), cv2.FILLED)
                annotations[annotationNumber].append(indexFinger)
            else:
                annotationStart=False
        else:
            annotationStart = False

        if buttonPressed:
            counter += 1
            if counter > delay:
                counter = 0
                buttonPressed = False

        for i in range(len(annotations)):
            for j in range(len(annotations[i])):
                if j != 0:
                    cv2.line(imgCurrent, annotations[i][j - 1], annotations[i][j], (0, 0, 200), 12)

        imgSmall = cv2.resize(img, (ws, hs))
        h, w, _ = imgCurrent.shape
        imgCurrent[0:hs, w - ws: w] = imgSmall

        cv2.imshow("Slides", imgCurrent)
        cv2.imshow("Image", img)

        key = cv2.waitKey(1)
        if key == ord('q'):
            break

## Create window
root = tk.Tk()
root.minsize(width = 400, height = 420)
root.maxsize(width = 400, height = 420)

root.title("PRESENTATION CONTROLLER")
# Set the theme background to black and foreground to green
root.configure(bg="black")

# Create a label for heading
heading_label = tk.Label(root, text="HAND GESTURE PRESENTATION CONTROLLER", fg="yellow", bg="black", font=("Arial", 18, "bold"),wraplength=400)
heading_label.pack(pady=10)

# Create a label for project subheading
subheading_label = tk.Label(root, text="Final Year College Project", fg="white", bg="black", font=("Arial", 16, "bold"))
subheading_label.pack(pady=(5,20))

# Create button to browse and convert
convert_button1 = tk.Button(root, text="Presentation Conversion", fg="black", bg="White", font=("Arial", 16, "bold"), command=browse_and_convert)
convert_button1.pack(pady=(5,30))

convert_button2 = tk.Button(root, text="Presentation Controller", fg="black", bg="white", font=("Arial", 16, "bold"), command=hand_gesture_controller)
convert_button2.pack(pady=(5,20))

# text area
label_area = tk.Label(root, text="Result" ,fg="black", bg="white", font=("Arial", 16, "bold"))
label_area.pack()

root.mainloop()



