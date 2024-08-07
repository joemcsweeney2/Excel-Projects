import subprocess
import os

def run_test_script():
    script_path = r"C:\Dev\GOODE_ANALYSIS\Scripts\xlwings_test.py"
    debug_info = []

    try:
        debug_info.append(f"Running script: {script_path}")
        result = subprocess.run(["python", script_path], check=True, capture_output=True, text=True)
        debug_info.append(f"Script output: {result.stdout}")
        debug_info.append(f"Script error output: {result.stderr}")
    except subprocess.CalledProcessError as e:
        debug_info.append(f"CalledProcessError: {e}")
        debug_info.append(f"Output: {e.output}")
    except FileNotFoundError as e:
        debug_info.append(f"FileNotFoundError: {e}")
    except Exception as e:
        debug_info.append(f"Unexpected error: {e}")

    log_debug_info(debug_info)

def log_debug_info(debug_info):
    log_dir = r'C:\Dev\GOODE_ANALYSIS\Scripts'
    logs = [f for f in os.listdir(log_dir) if f.startswith("debug_report")]
    new_log_file = os.path.join(log_dir, f"debug_report_{len(logs) + 1}.txt")

    with open(new_log_file, 'w') as log_file:
        for line in debug_info:
            log_file.write(line + '\n')

if __name__ == "__main__":
    debug_info = ["Starting run_test_script"]
    try:
        run_test_script()
        debug_info.append("run_test_script executed successfully.")
    except Exception as e:
        debug_info.append(f"Error in run_test_script: {e}")
    log_debug_info(debug_info)
