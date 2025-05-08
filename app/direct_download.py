#!/usr/bin/env python3
"""
Alternative direct download approach for Alberta Cannabis order form.
If the main method isn't working, this takes a more simplified approach.
"""
import requests
import os
from pathlib import Path
import sys

def direct_download():
    """
    Try to download the file directly (skipping API approaches)
    and save to the working directory
    """
    # Set up a persistent session with headers that look like a browser
    session = requests.Session()
    session.headers.update({
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/135.0.0.0 Safari/537.36"
        ),
        "Accept": "*/*",
        "Accept-Language": "en-US,en;q=0.9",
        "Connection": "keep-alive"
    })
    
    # Direct download URL (this should be the actual file URL)
    # This URL might be different - check the network tab in browser dev tools
    # when downloading the file manually to get the correct URL
    try:
        # Try method 1: direct file URL
        url = "https://retail.albertacannabis.org/media/default/order-form/CannabisRetailersManualOrderForm.xlsm"
        print(f"Attempting direct download from {url}")
        
        response = session.get(url, stream=True)
        if response.status_code == 200 and response.headers.get('Content-Type') in [
            'application/vnd.ms-excel.sheet.macroEnabled.12',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'application/octet-stream'
        ]:
            output_file = "CannabisRetailersManualOrderForm_direct.xlsm"
            with open(output_file, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)
            
            file_size = Path(output_file).stat().st_size
            print(f"Successfully downloaded file to {output_file} ({file_size} bytes)")
            return True
        else:
            print(f"Method 1 failed - Status: {response.status_code}, Content-Type: {response.headers.get('Content-Type')}")
    
    except Exception as e:
        print(f"Method 1 error: {e}")
    
    # If method 1 failed, try method 2: download from media library
    try:
        url = "https://retail.albertacannabis.org/api/cxa/QuickOrder/DownloadQuickOrderForm"
        params = {
            "downloadFileName": "CannabisRetailersManualOrderForm.xlsm",
            "mediaLibraryGuid": "51f2bf35-856d-484e-b84b-f6e66710b54b"
        }
        
        print(f"Attempting download from {url} with params {params}")
        response = session.get(url, params=params, stream=True)
        
        # Save whatever we get, even if it's not an Excel file
        output_file = "CannabisRetailersManualOrderForm_method2.xlsm"
        with open(output_file, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                f.write(chunk)
        
        file_size = Path(output_file).stat().st_size
        print(f"Downloaded file to {output_file} ({file_size} bytes)")
        
        # Save headers for debugging
        with open("download_headers.txt", "w") as f:
            for k, v in response.headers.items():
                f.write(f"{k}: {v}\n")
        
        if file_size > 0:
            with open(output_file, 'rb') as f:
                header = f.read(4)
                if header.startswith(b"PK"):
                    print("File appears to be a valid ZIP-based Office document")
                    return True
                else:
                    print(f"File doesn't appear to be a valid Office file (first 4 bytes: {header})")
        
        return False
        
    except Exception as e:
        print(f"Method 2 error: {e}")
        return False

if __name__ == "__main__":
    success = direct_download()
    sys.exit(0 if success else 1)