# Hand Gesture Presentation Controller

## Overview
This project enables users to control PowerPoint presentations using hand gestures detected through a webcam. It uses computer vision techniques to track hand movements and map them to presentation controls, eliminating the need for traditional input devices such as a mouse or keyboard.

The system converts presentation slides into images and allows users to navigate through them using predefined gestures, providing an intuitive and touch-free interaction experience.

---

## Objectives

- Develop a system to detect and interpret hand gestures in real time  
- Enable gesture-based navigation of presentation slides  
- Convert PowerPoint slides into image format for processing  
- Build a user-friendly interface for controlling the application  
- Ensure smooth and responsive gesture recognition  

---

## Features

### Hand Gesture Detection and Control
- Real-time hand tracking using webcam input  
- Gesture recognition mapped to presentation actions  
- Smooth slide transitions using gesture commands  

### Supported Gestures

- Next Slide: [1, 0, 0, 0, 0]  
- Previous Slide: [0, 0, 0, 0, 1]  
- Last Slide: [0, 0, 0, 1, 1]  
- Pointer Control: [0, 1, 1, 0, 0]  
- Activate Pointer: [0, 1, 0, 0, 0]  
- Clear Last Action: [0, 1, 1, 1, 0]  
- Clear Screen: [0, 1, 1, 1, 1]  

---

### Slide Processing
- Converts PowerPoint (.pptx) slides into image format (.png)  
- Enables efficient rendering and navigation of slides  

---

### User Interface
- GUI built using tkinter  
- Options to start, pause, and stop the system  
- File selection for PowerPoint input  
- Real-time gesture feedback  

---

## Tech Stack

- Python  
- OpenCV (cv2)  
- cvzone  
- tkinter  
- spire.Presentation  
- os  

---

## Project Workflow

1. Load PowerPoint file  
2. Convert slides into image format  
3. Capture webcam input  
4. Detect and track hand gestures  
5. Map gestures to presentation controls  
6. Display slides and update based on gestures  

---

## Use Cases

- Touchless presentation control  
- Smart classrooms and presentations  
- Assistive technology for hands-free interaction  
- Computer vision-based application development  

---

## How to Run

1. Clone the repository  
```
git clone https://github.com/Nicknishchal/Hand-Gesture-Presentation-Controller.git
```

2. Navigate to the project folder  
```
cd Hand-Gesture-Presentation-Controller
```

3. Install required libraries  
```
pip install -r requirements.txt
```

4. Run the main Python file  

---

## Contact

Email: saxenanishchal275@gmail.com  
LinkedIn: https://www.linkedin.com/in/nishchal-saxena-49551a271/

---

## Contribution

Contributions are welcome. Feel free to open issues or submit pull requests.
