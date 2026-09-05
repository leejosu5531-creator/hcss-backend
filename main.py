import os
import logging
import traceback
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("google_test")

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

class RequestModel(BaseModel):
    date_str: str
    total_seats: int = 67

SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]
DRIVE_FILE_NAME = "할렐루야 출석부.xlsx"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SERVICE_ACCOUNT_FILE = os.path.join(BASE_DIR, "service_account.json")

@app.post("/api/allocate")
def test_google_connection(req: RequestModel):
    logger.info("========== [구글 클라우드 연결 진단 시작] ==========")
    
    # [1단계] 파일 존재 여부 점검
    logger.info(f"🔍 [1단계] 인증 파일 존재 확인: {SERVICE_ACCOUNT_FILE}")
    if not os.path.exists(SERVICE_ACCOUNT_FILE):
        logger.error("❌ 1단계 실패: service_account.json 파일이 존재하지 않습니다.")
        raise HTTPException(status_code=400, detail="[1단계 실패] service_account.json 파일이 없음")
    logger.info("✅ 1단계 성공: 인증 파일 존재 확인됨")

    # [2단계] 토큰 생성 및 JWT 서명 검증
    try:
        logger.info("🔍 [2단계] OAuth2 Credentials 생성 및 토큰 리프레시 시도...")
        creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
        import google.auth.transport.requests
        request = google.auth.transport.requests.Request()
        creds.refresh(request) # 여기서 JWT 서명을 실제로 구글 서버에 검증합니다.
        logger.info("✅ 2단계 성공: 구글 클라우드 JWT 인증 및 토큰 발급 성공!")
    except Exception as e:
        logger.error(f"❌ 2단계 실패 (JWT/키 오류): {str(e)}")
        logger.error(traceback.format_exc())
        raise HTTPException(status_code=400, detail=f"[2단계 실패 - 키 손상] {str(e)}")

    # [3단계] Drive API 서비스 생성
    try:
        logger.info("🔍 [3단계] Google Drive API 서비스 객체 생성...")
        service = build("drive", "v3", credentials=creds)
        logger.info("✅ 3단계 성공: Drive API 서비스 생성 완료")
    except Exception as e:
        logger.error(f"❌ 3단계 실패: {str(e)}")
        raise HTTPException(status_code=400, detail=f"[3단계 실패] {str(e)}")

    # [4단계] 구글 드라이브 파일 검색 및 권한 검증
    try:
        logger.info(f"🔍 [4단계] 드라이브 내 '{DRIVE_FILE_NAME}' 파일 검색...")
        q = f"name = '{DRIVE_FILE_NAME.replace('\'', '\\\'')}' and trashed = false"
        r = service.files().list(q=q, spaces="drive", fields="files(id, name)").execute()
        files = r.get("files", [])
        
        if not files:
            logger.error("❌ 4단계 실패: 파일을 찾을 수 없습니다. (서비스 계정에 파일 공유 권한이 없는 경우 포함)")
            raise FileNotFoundError(f"'{DRIVE_FILE_NAME}' 파일을 찾을 수 없습니다.")
        
        file_info = files[0]
        logger.info(f"✅ 4단계 성공! 찾은 파일 ID: {file_info['id']}, 파일명: {file_info['name']}")
        
        return {
            "status": "success",
            "message": "구글 클라우드 모든 연동 프로세스가 정상입니다!",
            "file_id": file_info['id']
        }
    except Exception as e:
        logger.error(f"❌ 4단계 실패: {str(e)}")
        logger.error(traceback.format_exc())
        raise HTTPException(status_code=400, detail=f"[4단계 실패 - 파일/권한 오류] {str(e)}")