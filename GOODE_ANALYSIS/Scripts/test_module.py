import xlwings as xw

def test_function():
    wb = xw.Book.caller()
    sheet = wb.sheets['SELECTED DATA']  # Replace 'Sheet1' with the actual sheet name
    sheet.range('AA3').value = "Test function executed."

if __name__ == "__main__":
    test_function()
