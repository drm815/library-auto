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
import time
import os
from datetime import datetime

# ============================================================
# ★★★ 여기만 본인 정보로 수정하세요 ★★★
# ============================================================
아이디 = "여기에독서로아이디입력"
비밀번호 = "여기에독서로비밀번호입력"
FOLDER_ID = "여기에구글드라이브폴더ID입력"
SHEET_ID = "여기에구글시트ID입력"
# ============================================================

# 파일 경로 설정 (수정하지 마세요)
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OAUTH_FILE = os.path.join(SCRIPT_DIR, "oauth_credentials.json")
TOKEN_FILE = os.path.join(SCRIPT_DIR, "token.json")
DOWNLOAD_DIR = os.path.join(os.path.expanduser("~"), "Downloads")
SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets"
]

print("=" * 50)
print("📚 학교도서관 자동화 시스템 시작!")
print("=" * 50)

# Google 인증
print("\n🔐 Google 인증 중...")
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
print("✅ Google 인증 완료!")

# 크롬 설정
options = webdriver.ChromeOptions()
prefs = {"download.prompt_for_download": False}
options.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install()),
    options=options
)

try:
    # 독서로 접속
    print("\n🌐 독서로 접속 중...")
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
    print("✅ 로그인 완료!")

    # 소장자료관리 → 자료관리 이동
    time.sleep(3)
    driver.execute_script("move('CO010')")
    time.sleep(8)
    print("✅ 자료관리 페이지 이동 완료!")

    # 검색 버튼 클릭
    buttons = driver.find_elements(By.TAG_NAME, "button")
    for btn in buttons:
        if btn.text.strip() == "검색":
            btn.click()
            break
    time.sleep(10)
    print("✅ 검색 완료!")

    # 페이지 사이즈 변경 (전체 권수)
    try:
        page_input = driver.find_element(By.CLASS_NAME, "pageNum")
    except:
        page_input = driver.find_element(By.CSS_SELECTOR, "input.pageNum")
    page_input.clear()
    page_input.send_keys("99999")
    page_input.send_keys(Keys.ENTER)
    time.sleep(10)
    print("✅ 페이지 사이즈 변경 완료!")

    # 반출 버튼 클릭
    buttons = driver.find_elements(By.TAG_NAME, "button")
    for btn in buttons:
        if btn.text.strip() == "반출":
            btn.click()
            break
    time.sleep(15)
    print("✅ 다운로드 완료!")

finally:
    driver.quit()

# 다운로드된 파일 찾기
today = datetime.now().strftime("%Y%m%d_%H%M")
downloaded_file = None
for f in sorted(os.listdir(DOWNLOAD_DIR),
                key=lambda x: os.path.getmtime(os.path.join(DOWNLOAD_DIR, x)),
                reverse=True):
    if f.endswith(".xlsx"):
        downloaded_file = os.path.join(DOWNLOAD_DIR, f)
        break

if not downloaded_file:
    print("❌ 다운로드된 파일을 찾을 수 없어요!")
    exit()

# 파일 이름 변경
new_name = f"도서목록_{today}.xlsx"
new_path = os.path.join(DOWNLOAD_DIR, new_name)
os.rename(downloaded_file, new_path)
print(f"✅ 파일 이름 변경: {new_name}")

# Google Drive 업로드
print("\n📤 Google Drive 업로드 중...")
drive_service = build("drive", "v3", credentials=creds)

# 기존 파일 삭제
results = drive_service.files().list(
    q=f"name='도서목록_최신.xlsx' and '{FOLDER_ID}' in parents",
    fields="files(id, name)"
).execute()
for f in results.get("files", []):
    drive_service.files().delete(fileId=f["id"]).execute()
    print("🗑️ 기존 파일 삭제 완료!")

# 새 파일 업로드
file_metadata = {
    "name": "도서목록_최신.xlsx",
    "parents": [FOLDER_ID]
}
media = MediaFileUpload(
    new_path,
    mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
uploaded = drive_service.files().create(
    body=file_metadata,
    media_body=media,
    fields="id"
).execute()
print(f"✅ Google Drive 업로드 완료!")

# Google Sheets 업데이트
print("\n📊 Google Sheets 업데이트 중...")
df = pd.read_excel(new_path)
df = df.fillna("")

gc = gspread.authorize(creds)
sh = gc.open_by_key(SHEET_ID)
worksheet = sh.get_worksheet(0)

worksheet.clear()
data = [df.columns.tolist()] + df.values.tolist()
data = [[str(cell) for cell in row] for row in data]
worksheet.update("A1", data)

print(f"✅ Google Sheets 업데이트 완료! 총 {len(df)}행")
print("\n" + "=" * 50)
print(f"🎉 모든 작업 완료! ({today})")
print("=" * 50)
