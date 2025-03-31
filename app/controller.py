# This file is responsible for interacting between GUI and Logic. 
# It calls the logic function and returns the result to the GUI.
# It also handles any exceptions that may occur during the process and displays them.
from logic import run_anomaly_detection

def handle_anomaly_detection(csv_path, threshold, output_path, split_by, target_year):
    try:
        run_anomaly_detection(csv_path, threshold, output_path, split_by, target_year)
        return f"✅ Anomaly detection complete! Output saved to: {output_path}"
    except Exception as e:
        return f"❌ Error during anomaly detection: {str(e)}"
