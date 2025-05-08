# test_download.py
import sys
import requests
import os
import json
from bs4 import BeautifulSoup
import traceback

# Get credentials from environment or prompt
def get_credentials():
    try:
        from dotenv import load_dotenv
        load_dotenv()  # expects RETAIL_USER and RETAIL_PASS in .env
        USERNAME = os.getenv("RETAIL_USER") or input("Retailer Username: ")
        PASSWORD = os.getenv("RETAIL_PASS") or input("Retailer Password: ")
        return USERNAME, PASSWORD
    except ImportError:
        return input("Retailer Username: "), input("Retailer Password: ")

def debug_download():
    try:
        USERNAME, PASSWORD = get_credentials()
        
        # Create a session
        session = requests.Session()
        session.headers.update({
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/135.0.0.0 Safari/537.36"
            )
        })
        
        # 1. Get login page for CSRF token
        print("Step 1: Getting login page...")
        login_page = session.get("https://retail.albertacannabis.org/login")
        login_page.raise_for_status()
        soup = BeautifulSoup(login_page.text, "html.parser")
        token_input = soup.find("input", {"name": "__RequestVerificationToken"})
        if not token_input:
            print("ERROR: Could not find CSRF token on login page")
            with open("login_page_response.html", "w") as f:
                f.write(login_page.text)
            print("Saved login page response to login_page_response.html")
            return
            
        csrf_token = token_input["value"]
        print(f"Found CSRF token: {csrf_token[:10]}...")
        
        # 2. Login
        print("Step 2: Attempting login...")
        login_api = "https://retail.albertacannabis.org/api/cxa/AglcLogin/AglcLogin"
        login_payload = {
            "__RequestVerificationToken": csrf_token,
            "returnUrl": "/",
            "UserName": USERNAME,
            "Password": PASSWORD,
            "g-recaptcha-response": "",
            "g-recaptcha-currentItemId": ""
        }
        login_headers = {
            "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
            "X-Requested-With": "XMLHttpRequest",
            "Referer": "https://retail.albertacannabis.org/login"
        }
        
        resp = session.post(login_api, data=login_payload, headers=login_headers)
        resp.raise_for_status()
        result = resp.json()
        
        # Save login response for debugging
        with open("login_response.json", "w") as f:
            json.dump(result, f, indent=2)
            
        print(f"Login response status: {resp.status_code}")
        print(f"Login response JSON: {result}")
        
        if not result.get("success", True) and not result.get("user"):
            print("ERROR: Login failed")
            return
            
        print("Login appears successful")
        
        # 3. Try to download the file
        print("Step 3: Attempting to download file...")
        download_url = "https://retail.albertacannabis.org/api/cxa/QuickOrder/DownloadQuickOrderForm"
        download_params = {
            "downloadFileName": "CannabisRetailersManualOrderForm.xlsm",
            "mediaLibraryGuid": "51f2bf35-856d-484e-b84b-f6e66710b54b"
        }
        download_headers = {
            "Accept": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "Referer": "https://retail.albertacannabis.org/quick-order",
        }
        
        dl = session.get(download_url, params=download_params, headers=download_headers)
        
        print(f"Download response status: {dl.status_code}")
        print(f"Download content type: {dl.headers.get('Content-Type', 'unknown')}")
        print(f"Download content length: {len(dl.content)} bytes")
        
        # Save raw response and headers
        with open("download_response_headers.json", "w") as f:
            json.dump(dict(dl.headers), f, indent=2)
            
        with open("download_response.raw", "wb") as f:
            f.write(dl.content)
            
        if dl.content.startswith(b"PK"):
            print("SUCCESS: Response starts with PK signature (valid Office file)")
            # Save as Excel if valid
            with open("test_OrderForm.xlsm", "wb") as f:
                f.write(dl.content)
            print("Saved to test_OrderForm.xlsm")
        else:
            print("ERROR: Response does not start with PK signature")
            snippet = dl.content[:200].decode("utf-8", errors="replace")
            print(f"First 200 bytes: {snippet}")
            
            # Try to save as text for inspection if it's not binary
            try:
                with open("download_response.txt", "w") as f:
                    f.write(dl.content.decode("utf-8", errors="replace"))
                print("Saved text response to download_response.txt")
            except Exception as e:
                print(f"Could not save as text: {e}")
    
    except Exception as e:
        print(f"ERROR: {e}")
        traceback.print_exc()
        
if __name__ == "__main__":
    debug_download()
