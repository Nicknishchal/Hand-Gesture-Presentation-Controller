# Hand-Gesture-Presentation-Controller
The Hand Gesture Presentation Controller project is focused on creating an innovative way to control PowerPoint presentations using hand gestures. Leveraging computer vision techniques, this project utilizes the cvzone.handTracking module along with cv2 (OpenCV) to detect and track hand movements, enabling seamless control of presentation slides without the need for traditional input devices. The project also integrates tkinter for creating a user-friendly interface and os and spire.Presentation libraries for converting PowerPoint slides into .png images for gesture-based navigation.

Objectives:

To develop a system that tracks and interprets hand gestures using a webcam to control the flow of a PowerPoint presentation.
To convert PowerPoint slides into images (.png format) using the spire.Presentation library, making them accessible for gesture-based control.
To create an intuitive graphical user interface (GUI) using tkinter for easy interaction and control of the application.
To ensure smooth and responsive slide transitions based on specific hand gestures, such as swiping left or right to change slides.
To test and validate the system in various lighting conditions and with different users to ensure robustness and accuracy.

Key Features:

Hand Gesture Detection and Tracking:

Utilized cvzone.handTracking and cv2 (OpenCV) to detect and track hand gestures in real-time.
Implemented gesture recognition logic to map specific hand movements to presentation control actions (e.g., swiping to change slides).

Hand Gestures Used: 
Move to Next Slide: Thumb to the CAMERA [1,0,0,0,0]
Move to Prev Slide: Pnky Finger to the CAMERA [0,0,0,0,1]
Move to last Slide: Both Ring Finger and Pinky Finger to the CAMERA [0,0,0,1,1]
Control Pointer: Both Index Finger and Middle Finger to the CAMERA [0,1,1,0,0]
Access Pointer: Index Finger to the CAMERA [0,1,0,0,0]
Clear last action: Combination of Index, Middle, Ring Finger to the CAMERA [0,1,1,1,0]
Clear Screen: All Fingers (EXCEPT THUMB) to CAMERA [0,1,1,1,1]


Slide Conversion and Management:

Used spire.Presentation to convert PowerPoint (.pptx) slides into .png image format, enabling them to be processed and displayed by the application.
Managed slide navigation and display using os to ensure seamless slide transitions.

User Interface:

Developed a user-friendly GUI using tkinter to allow users to start, pause, and stop the presentation control system easily.
Provided options to select the PowerPoint file, adjust settings, and view real-time gesture detection feedback.

Tools and Technologies:

Programming Language: Python

Libraries: cvzone, cv2 (OpenCV), os, tkinter, spire.Presentation

Environment: VS Code
