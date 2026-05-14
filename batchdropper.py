import sys
import os
import subprocess
import time
from concurrent.futures import ThreadPoolExecutor, as_completed

def get_executable_dir():
    return os.path.dirname(os.path.abspath(sys.argv[0]))

def convert_windows(input_path, output_path):
    """Uses Windows COM in an isolated thread to convert files."""
    # We must import within the function or OS-guard it to prevent macOS crashes
    import pythoncom 
    import win32com.client

    # Initialize COM library for this specific background thread
    pythoncom.CoInitialize() 
    try:
        ext = os.path.splitext(input_path)[1].lower()
        
        if ext in ['.doc', '.docx']:
            # DispatchEx forces a new, separate instance of Word for parallel processing
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            doc = word.Documents.Open(input_path)
            doc.SaveAs(output_path, FileFormat=17) # 17 is PDF
            doc.Close()
            word.Quit()
            
        elif ext in ['.ppt', '.pptx']:
            ppt = win32com.client.DispatchEx("Powerpoint.Application")
            pres = ppt.Presentations.Open(input_path, WithWindow=False)
            pres.SaveAs(output_path, 32) # 32 is PDF
            pres.Close()
            ppt.Quit()
            
        return True, f"Success: {os.path.basename(output_path)}"
    except Exception as e:
        return False, f"Error on {os.path.basename(input_path)}: {e}"
    finally:
        # Release the COM resources for this thread
        pythoncom.CoUninitialize()

def convert_macos(input_path, output_path):
    """Uses AppleScript to convert files on macOS."""
    ext = os.path.splitext(input_path)[1].lower()
    
    word_script = f'''
    tell application "Microsoft Word"
        open POSIX file "{input_path}"
        set theDoc to active document
        save as theDoc file name POSIX file "{output_path}" file format format PDF
        close theDoc saving no
    end tell
    '''
    
    ppt_script = f'''
    tell application "Microsoft PowerPoint"
        open POSIX file "{input_path}"
        set thePres to active presentation
        save thePres in POSIX file "{output_path}" as save as PDF
        close thePres saving no
    end tell
    '''
    
    script = word_script if ext in ['.doc', '.docx'] else ppt_script
    
    try:
        subprocess.run(['osascript', '-e', script], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        return True, f"Success: {os.path.basename(output_path)}"
    except subprocess.CalledProcessError as e:
        return False, f"Error on {os.path.basename(input_path)}: {e}"

def process_file(file_path, exe_dir):
    """Router function for the thread pool."""
    base_name = os.path.splitext(os.path.basename(file_path))[0]
    output_pdf = os.path.join(exe_dir, f"{base_name}.pdf")

    if os.name == 'nt':
        return convert_windows(file_path, output_pdf)
    elif sys.platform == 'darwin':
        return convert_macos(file_path, output_pdf)
    else:
        return False, "Unsupported OS for native Office automation."

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Please drag and drop .doc, .docx, .ppt, or .pptx files onto the executable.")
        input("Press Enter to exit...")
        sys.exit(1)

    start_time = time.time()
    exe_dir = get_executable_dir()
    supported_extensions = {'.doc', '.docx', '.ppt', '.pptx'}
    
    # Filter out unsupported drops
    files_to_process = [
        f for f in sys.argv[1:] 
        if os.path.isfile(f) and os.path.splitext(f)[1].lower() in supported_extensions
    ]

    if not files_to_process:
        print("No supported files found in drop payload.")
        input("Press Enter to exit...")
        sys.exit(1)

    print(f"Initializing batch conversion for {len(files_to_process)} file(s)...")
    
    # Cap workers at 4. Opening more than 4 headless MS Office instances 
    # simultaneously causes diminishing returns and memory spikes.
    max_workers = min(4, len(files_to_process))
    
    success_count = 0
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        # Map futures to their file paths
        futures = {executor.submit(process_file, fp, exe_dir): fp for fp in files_to_process}
        
        for future in as_completed(futures):
            success, message = future.result()
            print(message)
            if success:
                success_count += 1

    elapsed = round(time.time() - start_time, 2)
    print(f"\nBatch complete! Converted {success_count}/{len(files_to_process)} files in {elapsed} seconds.")
    input("Press Enter to exit...")