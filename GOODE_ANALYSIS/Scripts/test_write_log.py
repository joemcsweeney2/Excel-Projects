import os

def test_write_log():
    log_dir = 'C:\\Dev\\GOODE_ANALYSIS\\Logs'
    log_file = os.path.join(log_dir, 'test_log.txt')
    try:
        with open(log_file, 'w') as log:
            log.write('This is a test log entry.\n')
        print("Log file created successfully.")
    except Exception as e:
        print(f"Error occurred: {e}")

test_write_log()
