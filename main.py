import os
import json
import base64
import logging
import io
import pandas as pd
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("hcss_backend")

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

def get_google_credentials():
    b64_str = os.environ.get("SERVICE_ACCOUNT_BASE64", "").strip()
    if not b64_str:
        raise ValueError("SERVICE_ACCOUNT_BASE64 환경변수가 설정되지 않았습니다.")
    
    decoded_json = base64.b64decode(b64_str).decode("utf-8")
    info = json.loads(decoded_json)
    return Credentials.from_service_account_info(info, scopes=SCOPES)

@app.post("/api/allocate")
def allocate_seats(req: RequestModel):
    try:
        # 1. 구글 클라우드 인증
        creds = get_google_credentials()
        service = build("drive", "v3", credentials=creds)
        
        # 2. 구글 드라이브에서 엑셀 파일 검색
        q = f"name = '{DRIVE_FILE_NAME.replace('\'', '\\\'')}' and trashed = false"
        r = service.files().list(q=q, spaces="drive", fields="files(id, name)").execute()
        files = r.get("files", [])
        
        if not files:
            raise FileNotFoundError(f"구글 드라이브에서 '{DRIVE_FILE_NAME}' 파일을 찾을 수 없습니다.")
            
        file_id = files[0]['id']
        
        # 3. 엑셀 파일 메모리로 다운로드
        request = service.files().get_media(fileId=file_id)
        file_stream = io.BytesIO()
        downloader = MediaIoBaseDownload(file_stream, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        
        file_stream.seek(0)
        
        # 4. Pandas로 엑셀 데이터 읽기 및 출석자 파싱
        # (시트 구성 및 데이터 형태에 맞춰 읽어옵니다)
        df = pd.read_excel(file_stream)
        
        # 날짜 컬럼 및 출석 데이터 확인 (예시 파싱 로직)
        # ※ 실제 엑셀 파일 컬럼명에 따라 조정될 수 있습니다.
        attendance_list = []
        
        # 엑셀 데이터 분석 후 파싱 진행
        # 데이터프레임 내에서 해당 날짜에 출석한 인원 추출
        for idx, row in df.iterrows():
            # 대원 이름, 파트 등의 정보 추출 로직
            pass

        # 5. 응답 데이터 생성
        return {
            "status": "success",
            "date": req.date_str,
            "message": f"{req.date_str} 출석부 데이터를 성공적으로 읽어왔습니다.",
            "file_id": file_id
        }

    except Exception as e:
        logger.error(f"오류 발생: {str(e)}")
        raise HTTPException(status_code=400, detail=str(e))