# This file is responsible for creating the GUI for the Anomaly Detector application.
# It uses the ttkbootstrap library for styling and tkinter for the GUI components.

import ttkbootstrap as ttk
from ttkbootstrap import Style
from tkinter import filedialog, messagebox
import tkinter as tk
from controller import handle_anomaly_detection
import threading
import time

class AnomalyDetectorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Anomaly Detector")
        self.center_window(1200, 1000)
        self.style = Style(theme="darkly")
        self.current_theme = "darkly"

        # Main Frame
        main_frame = ttk.Frame(root, padding=40)
        main_frame.pack(expand=True, fill='both')

        # CSV File Selection (No fixed theme/bootstyle so text inherits theme colors)
        self.csv_label = ttk.Label(main_frame, text="Select CSV File To Analyze:", font=("Helvetica", 18))
        self.csv_label.pack(pady=(20, 5))
        self.csv_entry = ttk.Entry(main_frame, width=100, bootstyle="primary", font=("Helvetica", 16))
        self.csv_entry.pack(pady=5)
        self.browse_button = ttk.Button(main_frame, text="Browse", command=self.browse_file, bootstyle="info", width=30)
        self.browse_button.pack(pady=10)

        # Threshold Selection
        self.threshold_label = ttk.Label(main_frame, text="Z-Score Threshold for Anomaly:", font=("Helvetica", 18))
        self.threshold_label.pack(pady=(20, 5))
        self.threshold_entry = ttk.Entry(main_frame, width=20, bootstyle="primary", font=("Helvetica", 16))
        self.threshold_entry.insert(0, "5.0")
        self.threshold_entry.pack(pady=5)

        # Target Year
        self.year_label = ttk.Label(main_frame, text="Year to Analyze:", font=("Helvetica", 18))
        self.year_label.pack(pady=(20, 5))
        self.year_entry = ttk.Entry(main_frame, width=20, bootstyle="primary", font=("Helvetica", 16))
        self.year_entry.insert(0, "0000")
        self.year_entry.pack(pady=5)

        # Dividing Sheets Option
        self.split_label = ttk.Label(main_frame, text="Divide Sheets By:", font=("Helvetica", 18))
        self.split_label.pack(pady=(20, 5))
        self.split_by = tk.StringVar(value="none")  # default is "none"
        frame_radio = ttk.Frame(main_frame)
        frame_radio.pack(pady=5)
        self.none_rb = ttk.Radiobutton(frame_radio, text="None", variable=self.split_by, value="none", bootstyle="primary")
        self.none_rb.pack(side="left", padx=10)
        self.category_rb = ttk.Radiobutton(frame_radio, text="Category", variable=self.split_by, value="category", bootstyle="primary")
        self.category_rb.pack(side="left", padx=10)
        self.muni_rb = ttk.Radiobutton(frame_radio, text="Municipality", variable=self.split_by, value="municipality", bootstyle="primary")
        self.muni_rb.pack(side="left", padx=10)

        # Save As Button
        self.save_button = ttk.Button(main_frame, text="Save As", command=self.save_as, bootstyle="info", width=30)
        self.save_button.pack(pady=10)

        # Run Button
        self.run_button = ttk.Button(main_frame, text="Run Anomaly Detection", command=self.run_detection, bootstyle="success", width=30)
        self.run_button.pack(pady=10)

        # Progress Bar (This doesn't actually show progress, just an indeterminate loading bar to show it is thinking)
        self.progress = ttk.Progressbar(main_frame, orient="horizontal", length=1000, mode="indeterminate")
        self.progress.pack(pady=20, fill='x')

        # Toggle Theme Button
        self.theme_button = ttk.Button(main_frame, text="Toggle Light/Dark Mode", command=self.toggle_theme, bootstyle="info", width=30)
        self.theme_button.pack(pady=10)

        # Store output path
        self.output_path = None

    # Center window on screen
    def center_window(self, width, height):
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    # Browse file system for CSV file
    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if file_path:
            self.csv_entry.delete(0, "end")
            self.csv_entry.insert(0, file_path)

    # Save output file to file system
    def save_as(self):
        self.output_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if self.output_path:
            messagebox.showinfo("Output Path", f"Output will be saved as: {self.output_path}")

    # Gathers information from GUI fields and validates them
    # Also starts progress bar
    def run_detection(self):
        csv_path = self.csv_entry.get()
        threshold = self.threshold_entry.get()
        target_year = self.year_entry.get()
        split_option = self.split_by.get()  # "none", "category", or "municipality"
        if not csv_path:
            messagebox.showerror("Error", "Please select a CSV file.")
            return
        if not self.output_path:
            messagebox.showerror("Error", "Please select an output location using 'Save As'.")
            return
        try:
            threshold = float(threshold)
            target_year = int(target_year)
        except ValueError:
            messagebox.showerror("Error", "Please enter valid numbers for threshold and target year.")
            return

        self.progress.start()
        threading.Thread(target=self.run_detection_thread, args=(csv_path, threshold, target_year, split_option)).start()

    # Run detection in separate thread
    # Also stops progress bar and shows result message
    def run_detection_thread(self, csv_path, threshold, target_year, split_option):
        time.sleep(2)
        result = handle_anomaly_detection(csv_path, threshold, self.output_path, split_option, target_year)
        self.progress.stop()
        messagebox.showinfo("Success", result)

    # Light/Dark Mode
    def toggle_theme(self):
        if self.current_theme == "darkly":
            self.style.theme_use("flatly")
            self.current_theme = "flatly"
        else:
            self.style.theme_use("darkly")
            self.current_theme = "darkly"
