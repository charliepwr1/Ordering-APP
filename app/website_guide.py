#!/usr/bin/env python3
"""
Detailed instructions for downloading the Cannabis Retailers Manual Order Form.
This script provides specific steps to find the download form and diagnose issues.
"""
import os
import sys

def display_website_guide():
    """Display detailed instructions for navigating the Alberta Cannabis website."""
    print("\n=== DETAILED GUIDE: Finding and Downloading the Order Form ===\n")
    
    print("IMPORTANT: The website structure or login process may have changed!")
    print("Follow these detailed steps to locate and download the correct file:\n")
    
    print("Step 1: Go to https://retail.albertacannabis.org/login")
    print("  - Complete the login form with your retailer credentials")
    print("  - Solve the CAPTCHA if presented (this is why automated download fails)\n")
    
    print("Step 2: After login, try these navigation paths:")
    print("  - Look for 'Quick Order' or 'Order Form' in the main navigation")
    print("  - Or check 'Resources', 'Forms', or 'Downloads' sections")
    print("  - Or go directly to: https://retail.albertacannabis.org/quick-order (after login)\n")
    
    print("Step 3: Check browser developer tools to find the correct download URL:")
    print("  - Right-click on the download link and select 'Inspect' or 'Inspect Element'")
    print("  - Look for the href attribute or download URL")
    print("  - Note this URL for updating the automation script later\n")
    
    print("Step 4: Save the downloaded Excel file")
    print("  - Ensure you're downloading an .xlsm file (Excel with macros)")
    print("  - The file should be several hundred KB in size, not just a few KB")
    print("  - Save it in this location:")
    print(f"    {os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'CannabisRetailersManualOrderForm.xlsm'))}\n")
    
    print("Step 5: If you find the correct download:")
    print("  - Take note of the URL, file name, and any parameters")
    print("  - This information can help update the automation script\n")
    
    print("Step 6: Troubleshooting tips:")
    print("  - If you see HTML when opening the downloaded file, you got the webpage, not the Excel file")
    print("  - Make sure you're fully logged in before attempting download")
    print("  - Try disabling any browser extensions that might interfere with downloads")
    print("  - Try a different browser if one doesn't work\n")
    
    print("Step 7: Update the app with your findings:")
    print("  - If the download URL has changed, update it in app/download_order_form.py")
    print("  - If the file structure or parameters have changed, note these details\n")
    
    print("=== DETAILED GUIDE COMPLETE ===\n")
    
    should_try = input("Would you like to try manual download now? (y/n): ").strip().lower()
    if should_try == 'y':
        print("\nStart the manual download process now. After downloading, place the file in the correct location.")
        print("Then restart the app to use the downloaded file.")

if __name__ == "__main__":
    display_website_guide()