import os
import requests
import sys
from bs4 import BeautifulSoup
from dotenv import load_dotenv

load_dotenv()  # expects RETAIL_USER and RETAIL_PASS in .env

def download_order_form():
    """
    Attempts to download the Cannabis Retailers Manual Order Form.
    Returns the file content as bytes if successful.
    Raises RuntimeError with detailed message if download fails.
    """
    USERNAME = os.getenv("RETAIL_USER") or input("Retailer Username: ")
    PASSWORD = os.getenv("RETAIL_PASS") or input("Retailer Password: ")

    session = requests.Session()
    session.headers.update({
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/135.0.0.0 Safari/537.36"
        )
    })

    try:
        # 1) GET login page → extract CSRF token & initial cookies
        login_page = session.get("https://retail.albertacannabis.org/login")
        login_page.raise_for_status()
        soup = BeautifulSoup(login_page.text, "html.parser")
        token_input = soup.find("input", {"name": "__RequestVerificationToken"})
        if not token_input:
            raise RuntimeError("Could not find CSRF token on login page")
        csrf_token = token_input["value"]

        # 2) POST credentials + token
        login_api = "https://retail.albertacannabis.org/api/cxa/AglcLogin/AglcLogin"
        login_payload = {
            "__RequestVerificationToken": csrf_token,
            "returnUrl": "/api/cxa/quickorder/downloadquickorderform?downloadFileName=CannabisRetailersManualOrderForm.xlsm&mediaLibraryGuid=51f2bf35-856d-484e-b84b-f6e66710b54b",
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
        if not result.get("success", True) or result.get("HasErrors", False):
            # Check for CAPTCHA error specifically
            if any('captcha' in str(err).lower() for err in result.get('Errors', [])):
                raise RuntimeError(
                    "Login failed: CAPTCHA verification required. "
                    "The website now requires manual CAPTCHA verification. "
                    "Please use the manual download option instead."
                )
            else:
                raise RuntimeError("Login failed: " + str(result))

        # Verify login success by checking a protected page
        dashboard = session.get("https://retail.albertacannabis.org/dashboard")
        if "Log out" not in dashboard.text and "Sign out" not in dashboard.text:
            raise RuntimeError("Login succeeded but session was not established properly.")

        # 3) Download the order form
        download_url = (
            "https://retail.albertacannabis.org/api/cxa/QuickOrder/DownloadQuickOrderForm"
        )
        # The mediaLibraryGuid might have changed. If this doesn't work,
        # you'll need to manually check the site to get the updated GUID.
        download_params = {
            "downloadFileName": "CannabisRetailersManualOrderForm.xlsm",
            "mediaLibraryGuid": "51f2bf35-856d-484e-b84b-f6e66710b54b"
        }
        
        download_headers = {
            "Accept": "application/vnd.ms-excel.sheet.macroEnabled.12,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/octet-stream,*/*",
            "Referer": "https://retail.albertacannabis.org/quick-order",
        }
        
        dl = session.get(download_url, params=download_params, headers=download_headers)
        dl.raise_for_status()

        content = dl.content
        content_type = dl.headers.get('Content-Type', 'unknown')
        print(f"Content-Type received: {content_type}")
        print(f"Response status code: {dl.status_code}")
        print(f"Content length: {len(content)} bytes")
        
        # Check if we received an Excel file by content type and signature
        if not content.startswith(b"PK") or (content_type.startswith('text/html') and len(content) < 1000000):
            # We got HTML instead of Excel
            snippet = content[:200].decode("utf-8", errors="replace")
            print(f"Response URL: {dl.url}")
            print(f"Full headers: {dict(dl.headers)}")
            
            # Save the HTML response for debugging
            debug_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'download_response.html')
            with open(debug_path, 'wb') as f:
                f.write(content)
            
            # Check if it's a login page (likely session expired)
            if 'login' in snippet.lower() or 'sign in' in snippet.lower():
                raise RuntimeError(
                    "Download returned a login page instead of an Excel file. "
                    "The session may have expired or the site's security has changed. "
                    "Please download the file manually."
                )
            
            raise RuntimeError(
                "Download did not return a valid Office file. "
                f"Content type was {content_type} (expected Excel). "
                f"Debug file saved to {debug_path}. "
                "First 200 bytes:\n" + snippet
            )

        return content
        
    except Exception as e:
        # Catch all exceptions and provide helpful error message
        print(f"Download failed: {str(e)}")
        print("\nThe website may require manual login with CAPTCHA verification.")
        print("Please try downloading the file manually:")
        print("1. Open https://retail.albertacannabis.org/ in your web browser")
        print("2. Log in with your retailer credentials")
        print("3. Navigate to Forms & Resources or Quick Order")
        print("4. Download the Cannabis Retailers Manual Order Form")
        print("5. Save it to your project directory")
        
        # Re-raise the exception so the caller can handle it
        raise

def check_local_file(file_path):
    """Check if a local copy of the file exists and is valid."""
    if not os.path.exists(file_path):
        return False
        
    try:
        with open(file_path, 'rb') as f:
            header = f.read(8)
            if not header.startswith(b"PK"):
                return False
        return True
    except Exception:
        return False

if __name__ == "__main__":
    # When run directly, download and save the file
    output_path = os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
        "CannabisRetailersManualOrderForm.xlsm"
    )
    
    # Check if we already have a valid local file
    if check_local_file(output_path):
        print(f"✅ Valid local order form already exists at {output_path}")
        print(f"   Size: {os.path.getsize(output_path):,} bytes")
        sys.exit(0)
    
    try:
        print("Attempting to download order form...")
        content = download_order_form()
        
        with open(output_path, "wb") as f:
            f.write(content)
        
        print(f"✅ Downloaded order form to {output_path}")
        print(f"   Size: {len(content):,} bytes")
        
        # Check if it looks like a valid Excel file
        if not content.startswith(b"PK"):
            print("⚠️ Warning: Downloaded file doesn't appear to be a valid Excel file.")
            print("   First few bytes:", content[:20])
            sys.exit(1)
        else:
            print("✓ File appears to be a valid Excel document.")
            sys.exit(0)
            
    except Exception as e:
        print(f"❌ Error downloading order form: {str(e)}")
        print("\nPlease download the form manually:")
        print("1. Open https://retail.albertacannabis.org/ in your browser")
        print("2. Log in with your retailer credentials")
        print("3. Navigate to Forms & Resources")
        print("4. Download the 'CannabisRetailersManualOrderForm.xlsm' file")
        print(f"5. Save it to: {output_path}")
        
        # Optionally launch the manual helper
        try:
            from manual_download_helper import guide_manual_download
            if input("\nWould you like to launch the manual download helper? (y/n): ").lower() == 'y':
                guide_manual_download()
        except ImportError:
            # Manual helper not available, that's fine
            pass
            
        sys.exit(1)