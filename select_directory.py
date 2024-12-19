
import tkinter as tk
from tkinter import filedialog
import sys

def select_directory():
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    directory = filedialog.askdirectory(title="Select Directory to Install")
    if not directory:
        print("No directory selected!")
        sys.exit(1)
    print(directory)

if __name__ == "__main__":
    select_directory()