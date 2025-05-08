#!/usr/bin/env python3
"""
Helper for manual download of Alberta Cannabis order form.
Since the website now requires CAPTCHA verification, automated login isn't feasible.
This script provides instructions for manual download and then processes the file.
"""
import os
import sys
import tkinter as tk
from tkinter import filedialog
import shutil

def guide_manual_download():
    """
    Provide instructions for manual download and then process the file
    """
    print("\n=== Alberta Cannabis Order Form Manual Download Helper ===\n")
    print("The website now requires CAPTCHA verification, which prevents automated login.")
    print("Follow these steps to manually download and process the file:\n")
    
    print("1. Open a web browser and go to: https://retail.albertacannabis.org/login")
    print("2. Log in with your retailer credentials")
    print("3. Navigate to the Quick Order page")
    print("4. Download the 'CannabisRetailersManualOrderForm.xlsm' file")
    print("5. When prompted, select the downloaded file to continue processing\n")
    
    input("Press Enter when you've downloaded the file and are ready to continue...")
    
    # Create a root window but hide it
    root = tk.Tk()
    root.withdraw()
    
    # Prompt user to select the downloaded file
    print("\nPlease select the downloaded CannabisRetailersManualOrderForm.xlsm file...")
    file_path = filedialog.askopenfilename(
        title="Select the downloaded Excel file",
        filetypes=[("Excel Files", "*.xlsm"), ("All Files", "*.*")]
    )
    
    if not file_path:
        print("No file selected. Exiting.")
        return False
    
    try:
        # Copy the file to the expected location
        target_path = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 
            "CannabisRetailersManualOrderForm.xlsm"
        )
        shutil.copy2(file_path, target_path)
        print(f"\nSuccess! File copied to: {target_path}")
        return True
    except Exception as e:
        print(f"Error copying file: {e}")
        return False

if __name__ == "__main__":
    success = guide_manual_download()
    sys.exit(0 if success else 1)