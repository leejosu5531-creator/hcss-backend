# -*- coding: utf-8 -*-
import os
import io
import json
import base64
import logging
import datetime as dt
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from openpyxl import load_workbook

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("hcss_backend")

app = FastAPI(title="화원교회 찬양대 좌석 배치 API")

# 스마트폰/웹 브라우저 접속 허용 (CORS)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]
DRIVE_FILE_NAME = "할렐루야 출석부.xlsx"

def get_google_credentials():
    b64_str = os.environ.get("SERVICE_ACCOUNT_BASE64", "").strip()
    if not b64_str:
        logger.error("❌ SERVICE_ACCOUNT_BASE64 환경변수가 설정되지 않았습니다.")
        raise ValueError("SERVICE_ACCOUNT_BASE64 환경변수가 설정되지 않았습니다.")
    
    decoded_json = base64.b64decode(b64_str).decode("utf-8")
    info = json.loads(decoded_json)
    return Credentials.from_service_account_info(info, scopes=SCOPES)

def google_drive_service():
    creds = get_google_credentials()
    return build("drive", "v3", credentials=creds)

def find_drive_file(service):
    q = f"name = '{DRIVE_FILE_NAME.replace('\'', '\\\'')}' and trashed = false"
    r = service.files().list(q=q, spaces="drive", fields="files(id,name)").execute()
    fs = r.get("files", [])
    if not fs:
        raise FileNotFoundError(f"구글 드라이브에서 '{DRIVE_FILE_NAME}' 출석부 파일을 찾지 못했습니다.")
    return fs[0]

def download_excel_to_stream(service, file_id):
    request = service.files().get_media(fileId=file_id)
    file_stream = io.BytesIO()
    downloader = MediaIoBaseDownload(file_stream, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    file_stream.seek(0)
    return file_stream

# --- HCSS 핵심 데이터 처리 로직 ---

def as_date(val):
    if isinstance(val, dt.datetime):
        return val.date()
    if isinstance(val, dt.date):
        return val
    if isinstance(val, str):
        val = val.strip()
        for fmt in ("%Y-%m-%d", "%Y.%m.%d", "%Y/%m/%d"):
            try:
                return dt.datetime.strptime(val, fmt).date()
            except ValueError:
                pass
    return None

def read_attendance(file_stream, target_date):
    wb = load_workbook(file_stream, data_only=True)
    ws = wb.active

    date_col = None
    for col in range(5, ws.max_column + 1):
        v = ws.cell(row=3, column=col).value
        d = as_date(v)
        if d and d == target_date:
            date_col = col
            break

    if not date_col:
        raise ValueError(f"해당 날짜({target_date})를 출석부 3행에서 찾을 수 없습니다.")

    people = {"S": [], "A": [], "T": [], "B": []}

    for row in range(4, ws.max_row + 1):
        p_val = ws.cell(row=row, column=1).value
        n_val = ws.cell(row=row, column=2).value
        c_val = ws.cell(row=row, column=date_col).value

        part = str(p_val).strip().upper() if p_val else ""
        name = str(n_val).strip() if n_val else ""

        if part in people and name:
            is_present = False
            if c_val is True:
                is_present = True
            elif isinstance(c_val, str):
                v = c_val.strip().upper()
                if v in ("O", "1", "TRUE", "Y", "○", "●"):
                    is_present = True
            elif isinstance(c_val, (int, float)) and c_val == 1:
                is_present = True

            if is_present:
                people[part].append(name)

    return people

def row_targets(N):
    if N == 67: return [15, 17, 17, 18]
    if N == 68: return [15, 17, 18, 18]
    if N == 69: return [15, 18, 18, 18]
    if N == 70: return [16, 18, 18, 18]
    
    base = N // 4
    rem = N % 4
    t = [base]*4
    for i in range(rem):
        t[3 - i] += 1
    return t

def allocate(people, total_seats):
    s_lst, a_lst = people["S"], people["A"]
    t_lst, b_lst = people["T"], people["B"]

    rt = row_targets(total_seats)
    organ_row = 1  # 2열 오르간

    a_req_org = max(0, 3 - len(s_lst))
    a_in_org = a_lst[:a_req_org]
    a_leftover = a_lst[a_req_org:]

    org_seats = []
    org_seats.extend([("S", n) for n in s_lst])
    org_seats.extend([("A", n) for n in a_in_org])

    rows = [[], [], [], []]
    leftovers = {"A": []}

    # 1열 배치
    cap1 = rt[0]
    take_t1 = min(len(t_lst), cap1)
    rows[0].extend([("T", n) for n in t_lst[:take_t1]])
    t_rem = t_lst[take_t1:]

    rem1 = cap1 - len(rows[0])
    take_b1 = min(len(b_lst), rem1)
    rows[0].extend([("B", n) for n in b_lst[:take_b1]])
    b_rem = b_lst[take_b1:]

    # 2열 배치
    cap2 = rt[1]
    rows[1].extend(org_seats)
    rem2 = cap2 - len(rows[1])
    if rem2 > 0:
        take_b2 = min(len(b_rem), rem2)
        rows[1].extend([("B", n) for n in b_rem[:take_b2]])
        b_rem = b_rem[take_b2:]

    # 3열, 4열 배치
    for r_idx in [2, 3]:
        cap = rt[r_idx]
        if t_rem:
            take_t = min(len(t_rem), cap - len(rows[r_idx]))
            rows[r_idx].extend([("T", n) for n in t_rem[:take_t]])
            t_rem = t_rem[take_t:]
        
        rem = cap - len(rows[r_idx])
        if rem > 0 and b_rem:
            take_b = min(len(b_rem), rem)
            rows[r_idx].extend([("B", n) for n in b_rem[:take_b]])
            b_rem = b_rem[take_b:]

    # 남은 알토 처리
    for r_idx in [2, 3]:
        rem = rt[r_idx] - len(rows[r_idx])
        if rem > 0 and a_leftover:
            take_a = min(len(a_leftover), rem)
            rows[r_idx].extend([("A", n) for n in a_leftover[:take_a]])
            a_leftover = a_leftover[take_a:]

    leftovers["A"] = a_leftover
    return rows, leftovers, rt, organ_row

# --- API 엔드포인트 ---

class RequestModel(BaseModel):
    date_str: str
    total_seats: int = 67

@app.post("/api/allocate")
def run_allocation(req: RequestModel):
    try:
        target = dt.date.fromisoformat(req.date_str)
        service = google_drive_service()
        file_info = find_drive_file(service)
        
        file_stream = download_excel_to_stream(service, file_info["id"])
        people = read_attendance(file_stream, target)
        rows, leftover, rt, org = allocate(people, req.total_seats)
        
        return {
            "status": "success",
            "date": str(target),
            "attending_count": sum(len(people[p]) for p in "SATB"),
            "row_targets": rt,
            "organ_row": org + 1,
            "rows": rows,
            "leftover": leftover
        }
    except Exception as e:
        logger.error(f"오류 발생: {str(e)}")
        raise HTTPException(status_code=400, detail=str(e))