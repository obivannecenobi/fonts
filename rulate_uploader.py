from __future__ import annotations

"""Utilities for uploading chapters to the rulate.ru translation platform.

The implementation relies on Selenium WebDriver and tries to keep the logic
simple enough to work in automated environments.  It accepts a list of files and
uploads each one sequentially, optionally configuring a few settings supported
by the site such as marking a chapter as deferred or requiring a subscription.

The function returns a mapping of file paths to a boolean flag indicating
whether the chapter appeared on the page after uploading.
"""

from typing import Iterable, Dict
import os

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager


def upload_chapters(
    book_url: str,
    files: Iterable[str],
    *,
    deferred: bool = False,
    subscription: bool = False,
    volume: int | None = None,
    publish_at: str | None = None,
    headless: bool = True,
) -> Dict[str, bool]:
    """Upload multiple chapter files to a book on Rulate.

    Parameters
    ----------
    book_url:
        Base URL to the book on ``rulate.ru``.  The function will navigate to the
        ``/chapter/new`` endpoint of this URL when uploading.
    files:
        Iterable of paths to chapter documents that should be uploaded.
    deferred:
        If ``True`` the chapters will be marked as deferred (draft mode).
    subscription:
        If ``True`` the chapters will require a subscription to read.
    volume:
        Optional volume number for the uploaded chapters.
    publish_at:
        Optional datetime string for scheduling publication.
    headless:
        When ``True`` (default) the browser runs in headless mode.

    Returns
    -------
    Dict[str, bool]
        Mapping of file path to a boolean indicating whether the chapter was
        detected on the resulting page after upload.
    """

    options = Options()
    if headless:
        options.add_argument("--headless")

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    wait = WebDriverWait(driver, 20)

    results: Dict[str, bool] = {}

    try:
        for file_path in files:
            # Navigate to the upload page for a new chapter
            driver.get(os.path.join(book_url, "chapter", "new"))

            # Upload the file
            file_input = wait.until(
                EC.presence_of_element_located((By.NAME, "file"))
            )
            file_input.send_keys(os.path.abspath(file_path))

            # Optional parameters
            if volume is not None:
                try:
                    volume_input = driver.find_element(By.NAME, "volume")
                    volume_input.clear()
                    volume_input.send_keys(str(volume))
                except Exception:
                    pass

            if deferred:
                try:
                    driver.find_element(By.NAME, "deferred").click()
                except Exception:
                    pass

            if subscription:
                try:
                    driver.find_element(By.NAME, "subscription").click()
                except Exception:
                    pass

            if publish_at:
                try:
                    publish_input = driver.find_element(By.NAME, "publish_at")
                    publish_input.clear()
                    publish_input.send_keys(publish_at)
                except Exception:
                    pass

            # Submit the form
            submit_btn = driver.find_element(By.CSS_SELECTOR, "button[type='submit']")
            submit_btn.click()

            chapter_name = os.path.basename(file_path)
            try:
                wait.until(
                    EC.presence_of_element_located(
                        (By.XPATH, f"//a[contains(@href, '{chapter_name}')]")
                    )
                )
                results[file_path] = True
            except TimeoutException:
                results[file_path] = False

        return results
    finally:
        driver.quit()
