import os
import tkinter as tk
from gui import ConverterApp
from logger import logger

def setup_folders():
    """Create essential folders if they don't exist yet."""
    folders = ["input", "output", "logs"]
    for folder in folders:
        if not os.path.exists(folder):
            os.makedirs(folder)
            logger.info(f"Created system folder: {folder}")

def main():
    logger.info("Starting CSV_Excel_Converter application...")
    setup_folders()
    
    root = tk.Tk()
    app = ConverterApp(root)
    root.mainloop()
    
    logger.info("Application closed successfully.")

if __name__ == "__main__":
    main()