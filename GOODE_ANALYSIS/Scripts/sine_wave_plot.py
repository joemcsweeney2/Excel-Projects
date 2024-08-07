import matplotlib.pyplot as plt
import numpy as np
import xlwings as xw
import os
import io
import tempfile

def log_message(message, log_num=1):
    log_dir = 'C:\\Dev\\GOODE_ANALYSIS\\Logs'
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
    log_file = os.path.join(log_dir, f'debug_log_{log_num:02}.txt')
    with open(log_file, 'a') as log:
        log.write(message + '\n')

def generate_sine_wave_plot():
    try:
        log_message("Generating sine wave data", 1)
        # Generate sine wave data
        x = np.linspace(0, 2 * np.pi, 100)
        y = np.sin(x)

        # Create plot
        plt.figure()
        plt.plot(x, y)
        plt.title("Sine Wave")
        
        log_message("Sine wave data generated", 1)

        # Save plot to a temporary file
        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as temp_file:
            plt.savefig(temp_file.name)
            temp_file_path = temp_file.name

        log_message(f"Plot saved to temporary file: {temp_file_path}", 1)

        # Insert plot into Excel
        wb = xw.Book.caller()
        sheet = wb.sheets['Sheet1']
        sheet.pictures.add(temp_file_path, name='SineWavePlot', update=True)

        log_message("Plot inserted into Excel", 1)

    except Exception as e:
        log_message(f"Error occurred: {e}", 1)
        raise

if __name__ == "__main__":
    xw.Book.caller()
    generate_sine_wave_plot()
