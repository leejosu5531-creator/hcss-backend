import os
import datetime as dt
import logging
import traceback
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

# 1. 디버깅 로깅 설정
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("allocation_debug")

app = FastAPI(title="화원교회 찬양대 좌석 배치 API")

# 2. CORS 미들웨어 설정 (Vercel 프론트엔드 통신 허용)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# 3. 데이터 모델 정의
class RequestModel(BaseModel):
    date_str: str
    total_seats: int = 67

# 4. 구글 드라이브 설정
SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]
DRIVE_FILE_NAME = "할렐루야 출석부.xlsx"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SERVICE_ACCOUNT_FILE = os.path.join(BASE_DIR, "service_account.json")

def google_drive_service():
    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    return build("drive", "v3", credentials=creds)

def find_drive_file(service):
    q = f"name = '{DRIVE_FILE_NAME.replace('\'', '\\\'')}' and trashed = false"
    r = service.files().list(q=q, spaces="drive", fields="files(id,name)").execute()
    fs = r.get("files", [])
    if not fs:
        raise FileNotFoundError(f"출석부 파일('{DRIVE_FILE_NAME}')을 찾지 못했습니다.")
    return fs[0]

@app.get("/")
def read_root():
    return {"message": "화원교회 찬양대 좌석 배치 API가 정상 동작 중입니다."}

@app.post("/api/allocate")
def run_allocation(req: RequestModel):
    logger.info("=== [1/6] 좌석 배치 요청 수신 ===")
    logger.info(f"요청 데이터: date_str={req.date_str}, total_seats={req.total_seats}")

    try:
        # [Step 1] 날짜 변환
        logger.info("=== [2/6] 날짜 데이터 변환 ===")
        target = dt.date.fromisoformat(req.date_str)

        # [Step 2] 구글 인증 파일 존재 여부 확인 및 서비스 생성
        logger.info("=== [3/6] Google Drive 서비스 인증 ===")
        logger.info(f"인증 파일 경로: {SERVICE_ACCOUNT_FILE}")
        if not os.path.exists(SERVICE_ACCOUNT_FILE):
            raise FileNotFoundError(f"service_account.json 파일이 경로에 없습니다: {SERVICE_ACCOUNT_FILE}")
        
        service = google_drive_service()
        logger.info("Google Drive 서비스 생성 성공")

        # [Step 3] 파일 검색
        logger.info("=== [4/6] 출석부 파일 검색 ===")
        file_info = find_drive_file(service)
        logger.info(f"파일 찾음: ID={file_info.get('id')}, Name={file_info.get('name')}")

        # 테스트용 임시 성공 응답
        return {
            "status": "success",
            "date": str(target),
            "attending_count": 0,
            "message": "구글 드라이브 파일 검색까지 정상 통과되었습니다!"
        }

    except Exception as e:
        logger.error("❌ 처리 중 에러 발생!")
        logger.error(traceback.format_exc())
        raise HTTPException(status_code=400, detail=str(e))
