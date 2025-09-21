import time
import os
import subprocess
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

class DocxChangeHandler(FileSystemEventHandler):
    def __init__(self, docx_path, script_path):
        self.docx_path = os.path.abspath(docx_path)
        self.script_path = os.path.abspath(script_path)

    def on_modified(self, event):
        if os.path.abspath(event.src_path) == self.docx_path:
            print(f"{self.docx_path} has been modified. Running conversion script...")
            result = subprocess.run(['python', self.script_path], capture_output=True, text=True)
            print(result.stdout)
            if result.stderr:
                print("Errors:", result.stderr)

if __name__ == "__main__":
    docx_file = "PASSWORDS.docx"
    conversion_script = "convert_docx_to_pdf.py"

    event_handler = DocxChangeHandler(docx_file, conversion_script)
    observer = Observer()
    observer.schedule(event_handler, path=os.path.dirname(os.path.abspath(docx_file)), recursive=False)
    observer.start()
    print(f"Watching for changes in {docx_file}... Press Ctrl+C to stop.")

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()
