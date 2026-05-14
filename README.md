# QuickPDF 📄⚡

![OS Compatibility](https://img.shields.io/badge/OS-Windows%20|%20macOS-blue)
![Python](https://img.shields.io/badge/Python-3.x-yellow)

QuickPDF is a hyper-threaded, zero-dependency command-line utility that converts Microsoft Word and PowerPoint documents into PDFs. By leveraging native OS automation (Windows COM and macOS AppleEvents), it requires no external layout engines or software installations for the end user—if Microsoft Office is installed on the host machine, QuickPDF just works.

Developed by **Moktadir Shourav** at **Akkel.IT**.

---

## ✨ Features

* **Zero Dependency Output:** End-users do not need to install LibreOffice, Ghostscript, or any other layout engines. It uses the native MS Office architecture already on the machine.
* **Hyper-Threaded Chunking:** Bypasses COM initialization bottlenecks by grouping files into chunks and processing them concurrently across a 4-worker thread pool, dramatically reducing conversion times.
* **Smart File-Lock Negotiation:** Includes a pre-flight lock detector that checks if files are held hostage by OneDrive, SharePoint, or Antivirus software, allowing the user to seamlessly wait, skip, or auto-poll.
* **Graceful Cancellation:** Runs a background daemon thread listening for the `Esc` key to safely terminate processes and cleanly kill headless COM engines without memory leaks.
* **Drag-and-Drop Simplicity:** Compiled as a standalone `.exe` or `.app`. Users simply drag their `.doc`, `.docx`, `.ppt`, or `.pptx` files directly onto the executable.
* **Live ETA:** Features a clean, in-place CLI progress bar with dynamic time estimations.

---

## 🚀 Usage (End User)

1. Highlight one or multiple supported Office documents.
2. Drag and drop them directly onto `QuickPDF.exe` (Windows) or `QuickPDF.app` (macOS).
3. The console will appear, run pre-flight lock checks, and generate the PDFs in the exact same directory as the executable.
4. Press `Esc` at any time to safely* abort the batch queue. 

*Note: Due to constraints of using the COM engine, the first file cannot be skipped once the process is started and may take ~30-40 seconds before the job is properly aborted.

---

## 🛠️ Building from Source (Developers)

If you want to compile the Python script into a standalone executable, follow these steps.

### Prerequisites
* Python 3.x
* `pyinstaller` (`pip install pyinstaller`)
* `pywin32` (`pip install pywin32` - *Required for Windows compilation only*)

### Compiling for Windows (.exe)
Run the following command in your terminal from the project directory:

```bash
pyinstaller --onefile --console --name "QuickPDF" quickpdf.py
```

### Compiling for macOS (.app)

Because macOS handles drag-and-drop file routing differently, the script is wrapped using Automator.

1. Compile the core binary first:
```Bash
pyinstaller --onefile --name "QuickPDF_Core" quickpdf.py
```

2. Open Automator on your Mac and create a new Application.

3. Search for and drag in the Run Shell Script action.

4. Set "Pass input" to as arguments.

5. Paste the path to your compiled binary, followed by "$@":

```Bash
"/path/to/your/dist/QuickPDF_Core" "$@"
```

6. Save the Automator file as QuickPDF.app.

## 🧠 Under the Hood: Architecture Notes

Dealing with GUI applications in a headless, automated batch environment presents unique challenges. This project solves them through:

- Isolated COM Apartments: On Windows, pythoncom.CoInitialize() is used to spin up dedicated COM apartments for each thread, preventing thread collisions when communicating with MS Word.

- Read-Only Overrides: Office applications aggressively lock files. QuickPDF passes ReadOnly=True flags directly to the COM/AppleScript dispatchers to read heavily guarded files.

- Chunking vs. Iteration: Instead of booting a headless instance of Word for every single file (which takes ~40 seconds per file), the thread pool groups files and feeds them to persistent background instances of Office, reducing conversion time to mere seconds.

## 👨‍💻 Author & Links

### Moktadir Shourav

### Systems & Data Engineering @ Akkel.IT

- GitHub: github.com/moktadir