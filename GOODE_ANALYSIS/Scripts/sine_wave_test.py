import matplotlib.pyplot as plt
import numpy as np
import xlwings as xw
import io

def generate_sine_wave_plot():
    # Generate sine wave data
    x = np.linspace(0, 2 * np.pi, 100)
    y = np.sin(x)

    # Create plot
    plt.figure()
    plt.plot(x, y)
    plt.title("Sine Wave")

    # Save plot to a bytes buffer
    buf = io.BytesIO()
    plt.savefig(buf, format='png')
    buf.seek(0)

    # Insert plot into Excel
    wb = xw.Book.caller()
    sheet = wb.sheets['Sheet1']
    sheet.pictures.add(buf, name='SineWavePlot', update=True)

if __name__ == "__main__":
    xw.Book.caller()
    generate_sine_wave_plot()
