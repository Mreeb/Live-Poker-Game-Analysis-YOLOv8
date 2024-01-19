# Live-Poker-Game-Analysis-YOLOv8
Poker Symbol Recognition with YOLOv8 and Python

Overview:
This Python program captures frames from a live poker game on a specified website, runs a YOLOv8 model for symbol recognition, and saves the respective symbol colors in an Excel file. The script utilizes libraries such as pyautogui, webbrowser, tkinter, ultralytics, and openpyxl to achieve the desired functionality.

Features:
Live Frame Capture: The program uses pyautogui to capture frames from a specified region of the screen, focusing on the live poker game.

YOLOv8 Symbol Recognition: The captured frames are processed through a YOLOv8 model (Detector.pt) for symbol recognition. The detected symbols are then classified into color categories.

Excel File Output: The script generates an Excel file (New_Results2.xlsx) to store the recognized symbols' respective colors. Each row represents a captured frame, and columns represent different users and their symbol colors.

Usage:
Install the required libraries using: pip install pyautogui webbrowser tkinter ultralytics openpyxl.

Set the url variable to the specific website where the live poker game is hosted.

Run the script, and it will open the specified website in the default browser.

The program continuously captures frames, runs the YOLOv8 model, and updates the Excel file with symbol colors.

Customization:
Modify the region variable to adjust the screen capture region based on the poker game's layout.

Customize the YOLOv8 model or use a different model by replacing the Detector.pt file.

Dependencies:
pyautogui
webbrowser
tkinter
ultralytics
openpyxl
