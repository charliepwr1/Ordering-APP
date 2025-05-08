import os
import requests
from bs4 import BeautifulSoup
from dotenv import load_dotenv

load_dotenv()  # expects RETAIL_USER and RETAIL_PASS in .env

def download_order_form():
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
        "Accept": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Referer": "https://retail.albertacannabis.org/quick-order",
    }
    
    dl = session.get(download_url, params=download_params, headers=download_headers)
    dl.raise_for_status()

    content = dl.content
    content_type = dl.headers.get('Content-Type', 'unknown')
    print(f"Content-Type received: {content_type}")
    print(f"Response status code: {dl.status_code}")
    print(f"Content length: {len(content)} bytes")
    
    # Check if we received an Excel file by content type
    valid_excel_types = [
        'application/vnd.ms-excel.sheet.macroEnabled.12',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'application/octet-stream',
        'application/vnd.ms-excel'
    ]
    
    # sanity‐check first two bytes are 'PK' (signature for Office files)
    if not content.startswith(b"PK") or (content_type.startswith('text/html') and len(content) < 1000000):
        # We got HTML instead of Excel
        snippet = content[:200].decode("utf-8", errors="replace")
        print(f"Response URL: {dl.url}")
        print(f"Full headers: {dict(dl.headers)}")
        
        # Save the HTML response for debugging
        debug_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'debug_download_response.html')
        with open(debug_path, 'wb') as f:
            f.write(content)
        
        raise RuntimeError(
            "Download did not return a valid Office file. "
            f"Content type was {content_type} (expected Excel). "
            f"Debug file saved to {debug_path}. "
            "First 200 bytes:\n" + snippet
        )

    return content
