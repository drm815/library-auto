import sys
import logging
import os
import time
from datetime import datetime

# ============================================================
# ★★★ 여기만 본인 정보로 수정하세요 ★★★
# ============================================================
아이디 = ""
비밀번호 = ""
FOLDER_ID = ""
SHEET_ID = ""
# ============================================================

# 파일 경로 설정
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OAUTH_FILE = os.path.join(os.path.expanduser("~"), "Desktop", "oauth_credentials.json")
TOKEN_FILE = os.path.join(os.path.expanduser("~"), "Desktop", "token.json")
DOWNLOAD_DIR = os.path.join(os.path.expanduser("~"), "Downloads")
SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets"
]

# 로그 설정
LOG_FILE = os.path.join(SCRIPT_DIR, "library_auto_log.txt")
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s %(message)s",
    encoding="utf-8"
)
def log(msg):
    print(msg)
    logging.info(msg)

log("=" * 50)
log("📚 학교도서관 자동화 시스템 시작!")
log("=" * 50)

try:
    from selenium import webdriver
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    from webdriver_manager.chrome import ChromeDriverManager
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaFileUpload
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request
    import gspread
    import pandas as pd

    # Google 인증
    log("\n🔐 Google 인증 중...")
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(OAUTH_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, "w") as token:
            token.write(creds.to_json())
    log("✅ Google 인증 완료!")

    # 크롬 드라이버 준비
    log("\n🌐 크롬 드라이버 준비 중...")
    options = webdriver.ChromeOptions()
    prefs = {"download.default_directory": DOWNLOAD_DIR, "download.prompt_for_download": False}
    options.add_experimental_option("prefs", prefs)
    options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1920,1080")
    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options
    )
    log("✅ 크롬 드라이버 준비 완료!")

    try:
        # 독서로 접속
        log("\n🌐 독서로 접속 중...")
        driver.get("https://dls2.edunet.net/DLS/loginMain")
        time.sleep(3)

        # 로그인
        driver.execute_script("document.getElementById('selectArea').style.display='block'")
        time.sleep(1)
        driver.execute_script("""
            var select = document.getElementById('prov_code');
            select.value = 'D10';
            var event = new Event('change');
            select.dispatchEvent(event);
        """)
        time.sleep(2)
        driver.find_element(By.ID, "lgID").send_keys(아이디)
        time.sleep(1)
        driver.find_element(By.ID, "lgPW").send_keys(비밀번호)
        time.sleep(1)
        driver.find_element(By.ID, "loginBtn").click()
        time.sleep(8)
        log(f"✅ 로그인 완료! 현재 URL: {driver.current_url}")

        # 소장자료관리 이동
        time.sleep(3)
        driver.execute_script("move('CO010')")
        time.sleep(8)
        log(f"✅ 자료관리 이동 완료! 현재 URL: {driver.current_url}")

        # 검색 버튼 클릭
        buttons = driver.find_elements(By.TAG_NAME, "button")
        clicked = False
        for btn in buttons:
            if btn.text.strip() == "검색":
                btn.click()
                clicked = True
                break
        log(f"✅ 검색 버튼: {'성공' if clicked else '실패'}")
        time.sleep(10)

        # 페이지 사이즈 변경
        try:
            page_input = driver.find_element(By.CLASS_NAME, "pageNum")
        except:
            page_input = driver.find_element(By.CSS_SELECTOR, "input.pageNum")
        page_input.clear()
        page_input.send_keys("18291")
        page_input.send_keys(Keys.ENTER)
        time.sleep(10)
        log("✅ 페이지 사이즈 변경 완료!")

        # 반출 버튼 클릭
        buttons = driver.find_elements(By.TAG_NAME, "button")
        clicked = False
        for btn in buttons:
            if btn.text.strip() == "반출":
                btn.click()
                clicked = True
                break
        log(f"✅ 반출 버튼: {'성공' if clicked else '실패'}")
        time.sleep(20)
        log("✅ 다운로드 완료!")

    except Exception as e:
        log(f"❌ 브라우저 단계 오류: {e}")
        import traceback
        logging.error(traceback.format_exc())
        raise
    finally:
        driver.quit()
        log("🔒 브라우저 종료")

    # 다운로드 파일 찾기 (최근 5분 내)
    today = datetime.now().strftime("%Y%m%d_%H%M")
    downloaded_file = None
    for f in sorted(os.listdir(DOWNLOAD_DIR),
                    key=lambda x: os.path.getmtime(os.path.join(DOWNLOAD_DIR, x)),
                    reverse=True):
        if f.endswith(".xlsx"):
            full_path = os.path.join(DOWNLOAD_DIR, f)
            if (datetime.now().timestamp() - os.path.getmtime(full_path)) < 300:
                downloaded_file = full_path
                break

    if not downloaded_file:
        log("❌ 다운로드 파일 없음! (최근 5분 내 xlsx 없음)")
        sys.exit(1)

    log(f"✅ 다운로드 파일 확인: {downloaded_file}")

    # 파일 이름 변경
    new_name = f"도서목록_{today}.xlsx"
    new_path = os.path.join(DOWNLOAD_DIR, new_name)
    os.rename(downloaded_file, new_path)
    log(f"✅ 파일 이름 변경: {new_name}")

    # Google Drive 업로드
    log("\n📤 Google Drive 업로드 중...")
    drive_service = build("drive", "v3", credentials=creds)
    results = drive_service.files().list(
        q=f"name='도서목록_최신.xlsx' and '{FOLDER_ID}' in parents",
        fields="files(id, name)"
    ).execute()
    for f in results.get("files", []):
        drive_service.files().delete(fileId=f["id"]).execute()
        log("🗑️ 기존 파일 삭제!")

    file_metadata = {"name": "도서목록_최신.xlsx", "parents": [FOLDER_ID]}
    media = MediaFileUpload(
        new_path,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    drive_service.files().create(body=file_metadata, media_body=media, fields="id").execute()
    log("✅ Google Drive 업로드 완료!")

    # Google Sheets 업데이트
    log("\n📊 Google Sheets 업데이트 중...")
    df = pd.read_excel(new_path)
    df = df.fillna("")

    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SHEET_ID)
    worksheet = sh.worksheet("도서목록")

    worksheet.clear()
    data = [df.columns.tolist()] + df.values.tolist()
    data = [[str(cell) for cell in row] for row in data]
    worksheet.update("A1", data)

    log(f"✅ Google Sheets 업데이트 완료! 총 {len(df)}행")
    log("=" * 50)
    log(f"🎉 모든 작업 완료! ({today})")
    log("=" * 50)

except Exception as e:
    log(f"\n❌ 치명적 오류: {e}")
    import traceback
    logging.error(traceback.format_exc())
    sys.exit(1)
print("=" * 50)
