import os

def log_message(message):
    log_dir = r'C:\Dev\GOODE_ANALYSIS\Scripts'
    logs = [f for f in os.listdir(log_dir) if f.startswith("minimal_test_log")]
    new_log_file = os.path.join(log_dir, f"minimal_test_log_{len(logs) + 1}.txt")

    with open(new_log_file, 'w') as log_file:
        log_file.write(message + '\n')

if __name__ == "__main__":
    log_message("Minimal test script executed successfully.")
    print("Minimal test script executed successfully.")
