# This file serves as the root of the program for pyinstaller.
# It creates the main window and runs the application.
# To build the executable with pyinstaller, 
# run the following command (after installing it): pyinstaller --onefile --windowed Program/Anomaly_Detector.py
# The executable will be in the dist folder.

from gui import AnomalyDetectorApp
import tkinter as tk

if __name__ == "__main__":
    root = tk.Tk()
    app = AnomalyDetectorApp(root)
    root.mainloop() 