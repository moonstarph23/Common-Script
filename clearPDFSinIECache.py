import os
import glob
import shutil
import subprocess

# Get the current user's environment username
username = os.environ.get('USERNAME') or os.environ.get('USER')

# Path to the IE cache directory
ie_cache_dir = os.path.join(
    os.environ.get('LOCALAPPDATA') or os.path.join(
        os.environ.get('USERPROFILE', ''), 'AppData', 'Local'),
    r'Microsoft\Windows\INetCache\Low\IE'
)

# Function to clear all PDF files in the directory
def clear_pdf_cache(directory):
    if os.path.exists(directory):
        try:
            # Find all PDF files recursively
            pdf_files = glob.glob(os.path.join(directory, '**', '*.pdf'), recursive=True)
            for pdf in pdf_files:
                os.remove(pdf)
                print(f"Deleted PDF: {pdf}")
            print("All PDF files in IE cache have been cleared.")
        except Exception as e:
            print(f"Error clearing PDF cache: {e}")
    else:
        print("IE cache directory not found.")

# Clear the PDF cache
clear_pdf_cache(ie_cache_dir)

# Optionally, use RunDll32 to clear IE cache (uncomment if preferred)
try:
    result = subprocess.run(['RunDll32.exe', 'InetCpl.cpl,ClearMyTracksByProcess', '8'], capture_output=True, text=True)
    if result.returncode == 0:
        print("Internet Explorer cache cleared successfully.")
    else:
        print(f"Error clearing IE cache: {result.stderr}")
except Exception as e:
    print(f"Failed to clear IE cache: {e}")