import os
import json
import base64
import logging
import traceback
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

# 디버그 로그 설정
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("google_test")

app = FastAPI()

# CORS 설정 (Vercel 프론트엔드와 통신 허용)
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

def get_google_credentials():
    # Render의 Environment Variables에서 SERVICE_ACCOUNT_BASE64 값을 가져옵니다.
    b64_str = os.environ.get("SERVICE_ACCOUNT_BASE64", "").strip()
    
    if not b64_str:
        logger.error("❌ SERVICE_ACCOUNT_BASE64 환경변수를 찾을 수 없습니다!")
        raise ValueError("SERVICE_ACCOUNT_BASE64 환경변수가 Render 설정에 등록되지 않았습니다.")
    
    logger.info(f"✅ SERVICE_ACCOUNT_BASE64 환경변수 로드 성공 (데이터 길이: {len(b64_str)})")
    
    # Base64 디코딩 후 JSON 객체로 변환
    decoded_json = base64.b64decode(b64_str).decode("utf-8")
    info = json.loads(decoded_json)
    
    logger.info(f"🔑 읽어온 서비스 계정 이메일: {info.get('client_email')}")
    return Credentials.from_service_account_info(info, scopes=SCOPES)

@app.post("/api/allocate")
def test_google_connection(req: RequestModel):
    logger.info("========== [구글 클라우드 연결 진단 시작] ==========")
    
    # [1단계] SERVICE_ACCOUNT_BASE64 환경변수를 이용한 인증 토큰 생성
    try:
        creds = get_google_credentials()
        import google.auth.transport.requests
        request = google.auth.transport.requests.Request()
        creds.refresh(request) # 구글 인증 서버와 실제로 서명을 주고받습니다.
        logger.info("✅ 1단계 성공: 구글 클라우드 JWT 인증에 완벽히 성공했습니다!")
    except Exception as e:
        logger.error(f"❌ 1단계 인증 실패: {str(e)}")
        logger.error(traceback.format_exc())
        raise HTTPException(status_code=400, detail=f"[인증 실패] {str(e)}")

    # [2단계] Google Drive API 연결 및 출석부 엑셀 파일 검색
    try:
        service = build("drive", "v3", credentials=creds)
        q = f"name = '{DRIVE_FILE_NAME.replace('\'', '\\\'')}' and trashed = false"
        r = service.files().list(q=q, spaces="drive", fields="files(id, name)").execute()
        files = r.get("files", [])
        
        if not files:
            logger.error(f"❌ 파일 찾기 실패: '{DRIVE_FILE_NAME}' 파일을 찾을 수 없습니다.")
            raise FileNotFoundError(f"'{DRIVE_FILE_NAME}' 파일이 구글 드라이브에 존재하지 않거나, 서비스 계정에 공유 권한이 없습니다.")
            
        file_info = files[0]
        logger.info(f"✅ 2단계 성공! 찾은 파일 ID: {file_info['id']}")
        
        return {
            "status": "success",
            "message": "구글 드라이브 연결 및 인증에 완전히 성공했습니다!",
            "file_id": file_info['id']
        }
    except Exception as e:
        logger.error(traceback.format_exc())
        raise HTTPException(status_code=400, detail=f"[파일 검색 실패] {str(e)}")