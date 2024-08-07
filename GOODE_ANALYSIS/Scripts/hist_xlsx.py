import os
import io
import matplotlib.pyplot as plt
import xlwings as xw
import numpy as np
from scipy.interpolate import make_interp_spline

def log_message(message, log_num):
    log_dir = 'C:\\Dev\\GOODE_ANALYSIS\\Logs'
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
    log_file = os.path.join(log_dir, f'hist_xlsx_log_{log_num:02}.txt')
    with open(log_file, 'a') as log:
        log.write(message + '\n')

def get_next_log_number():
    log_dir = 'C:\\Dev\\GOODE_ANALYSIS\\Logs'
    if not os.path.exists(log_dir):
        return 1
    log_files = [f for f in os.listdir(log_dir) if f.startswith('hist_xlsx_log_') and f.endswith('.txt')]
    log_numbers = [int(f.split('_')[-1].split('.')[0]) for f in log_files]
    return max(log_numbers) + 1 if log_numbers else 1

def create_histogram():
    log_num = get_next_log_number()
    try:
        log_message("Starting histogram generation.", log_num)
        wb = xw.Book.caller()
        sheet = wb.sheets['SELECTED DATA']
        log_message(f"Sheet '{sheet.name}' selected.", log_num)
        
        named_ranges = [name.name for name in wb.names]
        log_message(f"Named ranges in workbook: {named_ranges}", log_num)

        try:
            filtered_frequency = sheet.range('FilteredFrequency').value
            bin_labels = sheet.range('BinLabels').value
            bin_midpoints = sheet.range('BinArrayMid').value
            bin_midpoints = [float(i) for i in bin_midpoints]  # Convert to float
            mean_value = float(sheet.range('B8').value)  # Ensure mean value is float
            filtered_normal = sheet.range('FilteredNormal').value
            log_message(f"FilteredFrequency: {filtered_frequency}", log_num)
            log_message(f"BinLabels: {bin_labels}", log_num)
            log_message(f"BinMidpoints: {bin_midpoints}", log_num)
            log_message(f"Mean Value: {mean_value}", log_num)
            log_message(f"FilteredNormal: {filtered_normal}", log_num)
        except Exception as e:
            log_message(f"Error accessing named ranges: {e}", log_num)
            return

        if filtered_frequency and bin_labels and bin_midpoints and filtered_normal:
            log_message("All named ranges are available.", log_num)

            # Plot histogram without borders and bars touching
            plt.figure(figsize=(10, 6))
            plt.bar(bin_labels, filtered_frequency, alpha=0.7, edgecolor='none', width=1, label='Histogram')
            plt.xlabel('Bins')
            plt.ylabel('Frequency')
            plt.title('Histogram')
            plt.xticks(rotation=45, ha='right')

            # Secondary horizontal axis
            secax = plt.gca().secondary_xaxis('top')
            secax.set_xlabel('Secondary Axis')
            secax.set_xticks(np.arange(len(bin_midpoints)))
            secax.set_xticklabels(np.round(bin_midpoints, 2), rotation=45, ha='right')
            secax.set_xlim(-0.5, len(bin_midpoints) - 0.5)

            # Smoothing the lines
            x_smooth = np.linspace(0, len(bin_labels) - 1, 300)
            spline_actual = make_interp_spline(np.arange(len(filtered_frequency)), filtered_frequency, k=3)
            spline_normal = make_interp_spline(np.arange(len(filtered_normal)), filtered_normal, k=3)

            plt.plot(x_smooth, spline_actual(x_smooth), color='blue', label='Actual Distribution')
            plt.plot(x_smooth, spline_normal(x_smooth), color='orange', label='Expected Distribution')

            # Mean line
            mean_idx = (np.abs(np.array(bin_midpoints) - mean_value)).argmin()  # Find closest bin midpoint to mean
            plt.axvline(x=mean_idx, color='red', linestyle='--', label='Mean')

            # Adjust layout to prevent cutoff
            plt.tight_layout()

            # Adding legend
            plt.legend()

            buf = io.BytesIO()
            plt.savefig(buf, format='png')
            buf.seek(0)
            log_message("Histogram image saved to buffer.", log_num)
            
            img_num = f'{log_num:02}'
            img_path = os.path.join('C:\\Dev\\GOODE_ANALYSIS\\Logs', f'temp_histogram_{img_num}.png')
            with open(img_path, 'wb') as f:
                f.write(buf.getbuffer())
            log_message(f"Histogram image saved to {img_path}.", log_num)
            
            # Insert image into Excel
            sheet.pictures.add(img_path, name='Histogram', update=True, left=sheet.range('AD1').left, top=sheet.range('AD1').top)
            log_message("Histogram inserted into Excel.", log_num)
            
    except Exception as e:
        log_message(f"Error occurred: {e}", log_num)
        raise

# For testing directly in Python
if __name__ == "__main__":
    create_histogram()
