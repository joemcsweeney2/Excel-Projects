# simple_test.py
import os

log_dir = 'C:\\Dev\\GOODE_ANALYSIS\\Logs'
log_file = os.path.join(log_dir, 'test_log.txt')

try:
    with open(log_file, 'w') as f:
        f.write("This is a test log file.\n")
    print(f"Successfully wrote to {log_file}")
except Exception as e:
    print(f"Failed to write to {log_file}: {e}")
