import os
import time
from pathlib import Path

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains

# =======================
# CONFIG
# =======================
USERNAME = os.environ["ENOTIF_USERNAME"]
PASSWORD = os.environ["ENOTIF_PASSWORD"]

download_dir = os.environ.get("DOWNLOAD_DIR", str(Path.cwd() / "downloads"))
download_dir_path = Path(download_dir)
download_dir_path.mkdir(parents=True, exist_ok=True)

# =======================
# EDGE OPTIONS
# =======================
edge_opts = webdriver.EdgeOptions()
edge_opts.add_experimental_option("prefs", {
    "download.default_directory": str(download_dir_path),
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    # try to reduce blocking (may still depend on Edge security settings/policies)
    "safebrowsing.enabled": True,
})

# GitHub Actions friendly
edge_opts.add_argument("--headless=new")
edge_opts.add_argument("--no-sandbox")
edge_opts.add_argument("--disable-dev-shm-usage")
edge_opts.add_argument("--window-size=1920,1080")

driver = webdriver.Edge(options=edge_opts)
wait = WebDriverWait(driver, 30)

def wait_for_new_xls(download_folder: Path, start_time: float, timeout=180):
    """
    Wait for a NEW .xls file created after start_time.
    Also ensures there is no .crdownload still in progress for that file.
    """
    end = time.time() + timeout
    newest = None

    while time.time() < end:
        # if any .crdownload exists and is recent, keep waiting
        crs = list(download_folder.glob("*.crdownload"))
        if crs:
            time.sleep(1)
            continue

        # find .xls created after start_time
        candidates = []
        for f in download_folder.glob("*.xls"):
            try:
                if f.stat().st_mtime >= start_time:
                    candidates.append(f)
            except FileNotFoundError:
                pass

        if candidates:
            newest = max(candidates, key=lambda x: x.stat().st_mtime)
            return newest

        time.sleep(1)

    raise TimeoutError("No new .xls download completed (still blocked or still downloading).")

try:
    # 1) Login page
    driver.get("http://enotifikasi.moh.gov.my/Login.aspx")

    # username
    wait.until(EC.presence_of_element_located((By.ID, "txtUsrCd"))).send_keys(USERNAME)
    # password
    wait.until(EC.presence_of_element_located((By.ID, "txtUsrPwd"))).send_keys(PASSWORD)
    # login
    wait.until(EC.element_to_be_clickable((By.ID, "btnLogin"))).click()

    # 2) Hover "Muat Turun" then click "Muat Turun Fail"
    muat_turun_menu = wait.until(EC.visibility_of_element_located((By.LINK_TEXT, "Muat Turun")))
    ActionChains(driver).move_to_element(muat_turun_menu).perform()

    muat_turun_fail = wait.until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "a[href='/UserInterface/Download/Download.aspx']"))
    )
    muat_turun_fail.click()

    # 3) If page uses iframe, switch in
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "iframe")))
    iframe = driver.find_element(By.TAG_NAME, "iframe")
    driver.switch_to.frame(iframe)

    # 4) Tick "Semua Medan"
    semua_medan = wait.until(EC.element_to_be_clickable((By.ID, "ctl00_ContentPlaceHolder1_boolCheckAll")))
    semua_medan.click()

    # 5) Click "Muat Turun" button
    start_time = time.time()  # mark time just before triggering download
    muat_turun_btn = wait.until(EC.element_to_be_clickable((By.ID, "ctl00_ContentPlaceHolder1_btnSearch")))
    muat_turun_btn.click()

    # 6) Wait for the actual .xls to fully appear (not .crdownload)
    xls_file = wait_for_new_xls(download_dir_path, start_time, timeout=240)
    print("Download finished:", xls_file)

finally:
    # Keep browser open if detach=True
    pass
