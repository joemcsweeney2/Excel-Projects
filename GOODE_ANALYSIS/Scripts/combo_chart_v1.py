import xlwings as xw
import matplotlib.pyplot as plt
import os

def create_combo_chart():
    def get_unique_filename(base, ext):
        counter = 1
        while True:
            filename = f"{base}_{counter:02d}.{ext}"
            if not os.path.exists(filename):
                return filename
            counter += 1

    log_file = get_unique_filename('C:\\Dev\\GOODE_ANALYSIS\\Scripts\\debug_log', 'txt')
    chart_path = get_unique_filename('C:\\Dev\\GOODE_ANALYSIS\\Scripts\\combo_chart', 'png')

    def log_message(message):
        with open(log_file, 'a') as f:
            f.write(message + '\n')
        print(message)

    log_message("Script started.")
    wb = xw.Book.caller()
    selected_data_sheet = wb.sheets['SELECTED DATA']
    
    log_message(f"Workbook accessed: {wb.name}")
    log_message(f"Sheet accessed: {selected_data_sheet.name}")
    
    log_message("Listing all named ranges:")
    for name in wb.names:
        log_message(f"Named range: {name.name}, Refers to: {name.refers_to}")

    filtered_frequency = None
    bin_labels = None
    filtered_normal = None

    try:
        filtered_frequency = selected_data_sheet.range('FilteredFrequency').value
        log_message(f"FilteredFrequency: {filtered_frequency}")
    except Exception as e:
        log_message(f"Error accessing 'FilteredFrequency': {e}")

    try:
        bin_labels = selected_data_sheet.range('BinLabels').value
        log_message(f"BinLabels: {bin_labels}")
    except Exception as e:
        log_message(f"Error accessing 'BinLabels': {e}")
    
    try:
        filtered_normal = selected_data_sheet.range('FilteredNormal').value
        log_message(f"FilteredNormal: {filtered_normal}")
    except Exception as e:
        log_message(f"Error accessing 'FilteredNormal': {e}")

    mean_value = selected_data_sheet.range('B8').value

    if filtered_frequency and bin_labels and filtered_normal:
        log_message("Data retrieved successfully. Plotting combo chart...")
        fig, ax1 = plt.subplots(figsize=(10, 6))

        # Plot histogram
        ax1.bar(bin_labels, filtered_frequency, color='blue', alpha=0.3, width=1.0, label='Actual Distribution')
        ax1.set_xlabel('Bins')
        ax1.set_ylabel('Frequency')
        ax1.set_title('Combo Chart: Histogram and Line Charts')
        ax1.tick_params(axis='x', rotation=90)

        # Plot lines
        ax2 = ax1.twinx()
        ax2.plot(bin_labels, filtered_frequency, color='green', marker='o', linestyle='-', label='Actual Distribution Line')
        ax2.plot(bin_labels, filtered_normal, color='orange', marker='o', linestyle='-', label='Expected Distribution Line')
        ax2.set_ylabel('Normalized Frequency')

        # Plot mean line
        ax1.axhline(y=mean_value, color='red', linestyle='--', label='Mean Value')

        fig.tight_layout()
        fig.legend(loc='upper right')
        fig.savefig(chart_path)
        plt.close(fig)
        log_message(f"Combo chart saved as '{chart_path}'.")

        # Insert the chart into the 'SELECTED DATA' sheet at AD1
        selected_data_sheet.pictures.add(chart_path, name='ComboChart', update=True, left=selected_data_sheet.range('AD1').left, top=selected_data_sheet.range('AD1').top)
        log_message("Combo chart inserted into the SELECTED DATA sheet at AD1.")
    else:
        log_message("Data retrieval failed.")

if __name__ == "__main__":
    xw.Book.caller().sheets[0].name
    create_combo_chart()
