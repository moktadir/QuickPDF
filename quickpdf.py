import sys
import os
import subprocess
import time
import math
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed

cancel_event = threading.Event()

def get_executable_dir():
    return os.path.dirname(os.path.abspath(sys.argv[0]))

def print_intro():
    """Displays the author credit and project links via ASCII banner."""
    banner = """
  ____        _      _     ____  ____  _____ 
 / __ \\      (_)    | |   |  _ \\|  _ \\|  ___|
| |  | |_   _ _  ___| | __| |_) | | | | |_   
| |  | | | | | |/ __| |/ /|  __/| | | |  _|  
| |__| | |_| | | (__|   < | |   | |_| | |    
 \\___\\_\\\\__,_|_|\\___|_|\\_\\|_|   |____/|_|    
                                               
======================================================
 Author: Moktadir Shourav @ Akkel.IT
 GitHub: https://github.com/moktadir
======================================================
"""
    print(banner)

def wait_for_any_key():
    print("\nPress any key to exit...")
    if os.name == 'nt':
        import msvcrt
        while msvcrt.kbhit(): 
            msvcrt.getch()
        msvcrt.getch()
    else:
        import tty
        import termios
        fd = sys.stdin.fileno()
        if sys.stdin.isatty():
            old_settings = termios.tcgetattr(fd)
            try:
                tty.setraw(sys.stdin.fileno())
                sys.stdin.read(1)
            finally:
                termios.tcsetattr(fd, termios.TCSADRAIN, old_settings)

def keyboard_listener():
    if os.name == 'nt':
        import msvcrt
        while not cancel_event.is_set():
            if msvcrt.kbhit():
                if msvcrt.getch() == b'\x1b':
                    cancel_event.set()
                    break
            time.sleep(0.1)
    else:
        import sys, select
        if not sys.stdin.isatty():
            return 
        import tty, termios
        fd = sys.stdin.fileno()
        old_settings = termios.tcgetattr(fd)
        try:
            tty.setcbreak(fd)
            while not cancel_event.is_set():
                if sys.stdin in select.select([sys.stdin], [], [], 0.1)[0]:
                    if sys.stdin.read(1) == '\x1b':
                        cancel_event.set()
                        break
        finally:
            termios.tcsetattr(fd, termios.TCSADRAIN, old_settings)

def is_file_locked(filepath):
    if not os.path.exists(filepath):
        return False
    try:
        os.rename(filepath, filepath)
        return False
    except OSError:
        return True

def format_time(seconds):
    mins, secs = divmod(int(seconds), 60)
    return f"{mins:02d}:{secs:02d}"

def print_progress(completed, total, start_time):
    if cancel_event.is_set():
        return
        
    percent = (completed / total) * 100
    elapsed = time.time() - start_time
    if completed > 0:
        time_per_item = elapsed / completed
        eta_seconds = time_per_item * (total - completed)
        eta_str = format_time(eta_seconds)
    else:
        eta_str = "--:--"
        
    filled_length = int(50 * completed // total)
    bar = '█' * filled_length + '-' * (50 - filled_length)
    
    sys.stdout.write(f"\rProgress: |{bar}| {percent:.1f}% | ETA: {eta_str} ")
    sys.stdout.flush()

def process_windows_chunk(files, exe_dir, is_word):
    import pythoncom 
    import win32com.client

    pythoncom.CoInitialize() 
    results = []
    app = None
    
    try:
        app_name = "Word.Application" if is_word else "Powerpoint.Application"
        app = win32com.client.DispatchEx(app_name)
        
        if is_word:
            app.Visible = False

        for file_path in files:
            if cancel_event.is_set():
                results.append((False, file_path, "Cancelled by user (Esc)."))
                continue

            base_name = os.path.splitext(os.path.basename(file_path))[0]
            output_pdf = os.path.join(exe_dir, f"{base_name}.pdf")
            
            try:
                if is_word:
                    doc = app.Documents.Open(file_path, ReadOnly=True)
                    doc.SaveAs(output_pdf, FileFormat=17) 
                    doc.Close()
                else:
                    pres = app.Presentations.Open(file_path, ReadOnly=True, WithWindow=False)
                    pres.SaveAs(output_pdf, 32) 
                    pres.Close()
                results.append((True, file_path, ""))
            except Exception as e:
                results.append((False, file_path, str(e)))
                
    except Exception as e:
        for f in files:
            results.append((False, f, f"COM Engine Failure: {e}"))
    finally:
        if app:
            try:
                app.Quit()
            except:
                pass
        pythoncom.CoUninitialize()
        
    return results

def process_macos_chunk(files, exe_dir, is_word):
    results = []
    for file_path in files:
        if cancel_event.is_set():
            results.append((False, file_path, "Cancelled by user (Esc)."))
            continue

        base_name = os.path.splitext(os.path.basename(file_path))[0]
        output_pdf = os.path.join(exe_dir, f"{base_name}.pdf")
        
        word_script = f'''
        tell application "Microsoft Word"
            open POSIX file "{file_path}" read only true
            set theDoc to active document
            save as theDoc file name POSIX file "{output_pdf}" file format format PDF
            close theDoc saving no
        end tell
        '''
        
        ppt_script = f'''
        tell application "Microsoft PowerPoint"
            open POSIX file "{file_path}" read only true
            set thePres to active presentation
            save thePres in POSIX file "{output_pdf}" as save as PDF
            close thePres saving no
        end tell
        '''
        
        script = word_script if is_word else ppt_script
        
        try:
            subprocess.run(['osascript', '-e', script], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            results.append((True, file_path, ""))
        except subprocess.CalledProcessError as e:
            results.append((False, file_path, str(e)))
            
    return results

def process_chunk_router(files, exe_dir, is_word):
    if os.name == 'nt':
        return process_windows_chunk(files, exe_dir, is_word)
    elif sys.platform == 'darwin':
        return process_macos_chunk(files, exe_dir, is_word)
    else:
        return [(False, f, "Unsupported OS") for f in files]

def chunk_list(lst, n):
    if not lst: return []
    chunk_size = math.ceil(len(lst) / n)
    return [lst[i:i + chunk_size] for i in range(0, len(lst), chunk_size)]

if __name__ == "__main__":
    # Display the branding splash immediately upon execution
    print_intro()

    if len(sys.argv) < 2:
        print("Please drag and drop .doc, .docx, .ppt, or .pptx files onto the executable.")
        wait_for_any_key()
        sys.exit(1)

    exe_dir = get_executable_dir()
    word_exts = {'.doc', '.docx'}
    ppt_exts = {'.ppt', '.pptx'}
    
    raw_word_files = [f for f in sys.argv[1:] if os.path.isfile(f) and os.path.splitext(f)[1].lower() in word_exts]
    raw_ppt_files = [f for f in sys.argv[1:] if os.path.isfile(f) and os.path.splitext(f)[1].lower() in ppt_exts]

    raw_total = len(raw_word_files) + len(raw_ppt_files)
    if raw_total == 0:
        print("No supported files found in drop payload.")
        wait_for_any_key()
        sys.exit(1)

    print(f"Running pre-flight lock checks on {raw_total} file(s)...\n")
    
    ready_word_files = []
    ready_ppt_files = []
    global_skip = False
    global_wait = False

    for collection, ready_collection in [(raw_word_files, ready_word_files), (raw_ppt_files, ready_ppt_files)]:
        for file_path in collection:
            filename = os.path.basename(file_path)
            while True:
                if not is_file_locked(file_path):
                    ready_collection.append(file_path)
                    break
                
                if global_skip:
                    print(f"Skipping locked file: {filename}")
                    break
                if global_wait:
                    print(f"Waiting on {filename} to unlock...")
                    time.sleep(2)
                    continue

                print(f"WARNING: '{filename}' is currently locked by another process.")
                print("  [W] Wait for this file | [S] Skip this file")
                print("  [A] Wait for ALL       | [X] Skip ALL")
                choice = input("Select an option (W/S/A/X): ").strip().upper()

                if choice == 'W':
                    print("Waiting... (Will automatically proceed when unlocked)")
                    time.sleep(2)
                elif choice == 'A':
                    print("Global Wait enabled. Polling for unlock...")
                    global_wait = True
                    time.sleep(2)
                elif choice == 'X':
                    print("Global Skip enabled. Skipping...")
                    global_skip = True
                    break
                elif choice == 'S':
                    print(f"Skipping {filename}...")
                    break
                else:
                    print("Invalid input. Please try again.")

    total_ready = len(ready_word_files) + len(ready_ppt_files)
    if total_ready == 0:
        print("\nNo files available to process.")
        wait_for_any_key()
        sys.exit(0)

    print(f"\nInitializing hyper-threaded batch conversion for {total_ready} file(s)...")
    print("(Press 'Esc' at any time to safely cancel the remaining queue)")
    
    word_chunks = chunk_list(ready_word_files, min(4, len(ready_word_files))) if ready_word_files else []
    ppt_chunks = chunk_list(ready_ppt_files, min(4, len(ready_ppt_files))) if ready_ppt_files else []
    all_chunks = [(chunk, True) for chunk in word_chunks] + [(chunk, False) for chunk in ppt_chunks]
    
    success_count = 0
    completed_files = 0
    error_log = []
    is_cancelling_printed = False
    
    listener_thread = threading.Thread(target=keyboard_listener, daemon=True)
    listener_thread.start()

    start_time = time.time()
    print_progress(0, total_ready, start_time)
    
    with ThreadPoolExecutor(max_workers=min(4, len(all_chunks))) as executor:
        futures = [executor.submit(process_chunk_router, chunk, exe_dir, is_word) for chunk, is_word in all_chunks]
        
        for future in as_completed(futures):
            chunk_results = future.result()
            for success, file_path, error_msg in chunk_results:
                completed_files += 1
                if success:
                    success_count += 1
                else:
                    filename = os.path.basename(file_path)
                    error_log.append(f"Failed {filename}: {error_msg}")
            
            if cancel_event.is_set():
                if not is_cancelling_printed:
                    sys.stdout.write("\n\n[!] ESC PRESSED: Cancelling remaining queue and cleaning up COM threads...\n")
                    is_cancelling_printed = True
            else:
                print_progress(completed_files, total_ready, start_time)

    cancel_event.set()

    print() 

    elapsed = round(time.time() - start_time, 2)
    if is_cancelling_printed:
        print(f"\nBatch Aborted! Converted {success_count}/{total_ready} files in {elapsed} seconds before cancellation.")
    else:
        print(f"\nBatch complete! Converted {success_count}/{total_ready} files in {elapsed} seconds.")
    
    if total_ready < raw_total:
        print(f"Note: {raw_total - total_ready} file(s) were skipped due to locks.")

    if error_log:
        print("\n--- Error Log ---")
        for err in error_log:
            print(err)

    wait_for_any_key()