"""
Keep-alive ping for a Streamlit Community Cloud app.

Streamlit Cloud puts free-tier apps to sleep after ~12 hours with no real
visitor traffic. A plain HTTP request (e.g. curl) doesn't count, because
Streamlit tracks live page sessions over WebSocket, not simple HTTP hits.
This script opens the app in a real (headless) browser, waits for it to
finish loading, and closes -- which counts as a genuine visit.

Run on a schedule via the GitHub Actions workflow in
.github/workflows/keep-alive.yml.
"""

import os
import sys
import time

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Set this as a repository variable/secret named STREAMLIT_APP_URL,
# or edit the default below.
APP_URL = os.environ.get("STREAMLIT_APP_URL", "https://your-app-name.streamlit.app")


def main():
    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1280,1696")

    # Selenium 4.6+ auto-detects and downloads the matching chromedriver,
    # no separate driver management needed.
    driver = webdriver.Chrome(options=options)

    try:
        print(f"Visiting {APP_URL} ...")
        driver.get(APP_URL)

        # Give the app time to wake up if it was asleep (cold start can take
        # a little while) and wait for Streamlit's main content to appear.
        WebDriverWait(driver, 90).until(
            EC.presence_of_element_located((By.TAG_NAME, "body"))
        )
        # Small extra pause so the WebSocket session is clearly established
        # before we disconnect.
        time.sleep(10)

        title = driver.title
        print(f"Loaded page. Title: '{title}'")
        print("Keep-alive visit complete.")
    except Exception as exc:
        print(f"Keep-alive visit failed: {exc}")
        sys.exit(1)
    finally:
        driver.quit()


if __name__ == "__main__":
    main()
