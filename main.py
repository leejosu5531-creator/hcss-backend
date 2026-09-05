# -*- coding: utf-8 -*-
import os
import io
import re
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

def parse_cell_date(val, target_year):
    if val is None:
        return None
    
    if isinstance(val, dt.datetime):
        return val.date()
    if isinstance(val, dt.date):
        return val

    if isinstance(val, (int, float)):
        try:
            return (dt.datetime(1899, 12, 30) + dt.timedelta(days=val)).date()
        except Exception:
            pass

    val_str = str(val).strip()
    if not val_str:
        return None

    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%Y.%m.%d", "%Y/%m/%d"):
        try:
            return dt.datetime.strptime(val_str, fmt).date()
        except ValueError:
            pass

    match = re.search(r"(\d{1,2})[\/\.\-\s월]+(\d{1,2})", val_str)
    if match:
        try:
            month = int(match.group(1))
            day = int(match.group(2))
            return dt.date(target_year, month, day)
        except ValueError:
            pass

    return None

def read_attendance(file_stream, target_date):
    wb = load_workbook(file_stream, data_only=True)
    
    sheet_part_map = {
        "소프라노": "S",
        "알토": "A",
        "테너": "T",
        "베이스": "B"
    }

    people = {"S": [], "A": [], "T": [], "B": []}
    target_year = target_date.year

    for sheet_name, part_code in sheet_part_map.items():
        if sheet_name not in wb.sheetnames:
            continue
        
        ws = wb[sheet_name]
        
        date_col = None
        for r_idx in range(1, 4):
            for col in range(1, ws.max_column + 1):
                cell_val = ws.cell(row=r_idx, column=col).value
                parsed_d = parse_cell_date(cell_val, target_year)
                if parsed_d and parsed_d.month == target_date.month and parsed_d.day == target_date.day:
                    date_col = col
                    break
            if date_col:
                break

        if not date_col:
            logger.warning(f"'{sheet_name}' 시트에서 날짜({target_date})열을 찾지 못했습니다.")
            continue

        for row in range(3, ws.max_row + 1):
            name_val = ws.cell(row=row, column=2).value
            if not name_val:
                continue
            
            name_str = str(name_val).strip()
            if name_str in ("구분", "월통계", "총인원") or "통계" in name_str:
                break
                
            c_val = ws.cell(row=row, column=date_col).value
            
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
                people[part_code].append(name_str)

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

# --- 요청사항 적용 좌석 배치 알고리즘 ---
def allocate(people, total_seats):
    s_lst = list(people["S"])
    a_lst = list(people["A"])
    t_lst = list(people["T"])
    b_lst = list(people["B"])

    rt = row_targets(total_seats)
    organ_row = 2  # 3열(인덱스 2)에 오르간 위치 지정

    rows = [[], [], [], []]
    leftovers = {"S": [], "A": [], "T": [], "B": []}

    # 1열: 알토 -> 소프라노
    cap1 = rt[0]
    take_a1 = min(len(a_lst), cap1)
    rows[0].extend([("A", n) for n in a_lst[:take_a1]])
    a_lst = a_lst[take_a1:]

    rem1 = cap1 - len(rows[0])
    if rem1 > 0:
        take_s1 = min(len(s_lst), rem1)
        rows[0].extend([("S", n) for n in s_lst[:take_s1]])
        s_lst = s_lst[take_s1:]

    # 2열: 알토 -> 소프라노
    cap2 = rt[1]
    take_a2 = min(len(a_lst), cap2)
    rows[1].extend([("A", n) for n in a_lst[:take_a2]])
    a_lst = a_lst[take_a2:]

    rem2 = cap2 - len(rows[1])
    if rem2 > 0:
        take_s2 = min(len(s_lst), rem2)
        rows[1].extend([("S", n) for n in s_lst[:take_s2]])
        s_lst = s_lst[take_s2:]

    # 3열: 베이스 -> 테너 -> 소프라노 (오르간 열)
    cap3 = rt[2]
    take_b3 = min(len(b_lst), cap3)
    rows[2].extend([("B", n) for n in b_lst[:take_b3]])
    b_lst = b_lst[take_b3:]

    rem3 = cap3 - len(rows[2])
    if rem3 > 0:
        take_t3 = min(len(t_lst), rem3)
        rows[2].extend([("T", n) for n in t_lst[:take_t3]])
        t_lst = t_lst[take_t3:]

    rem3 = cap3 - len(rows[2])
    if rem3 > 0:
        take_s3 = min(len(s_lst), rem3)
        rows[2].extend([("S", n) for n in s_lst[:take_s3]])
        s_lst = s_lst[take_s3:]

    # 4열: 베이스 -> 테너 -> 소프라노
    cap4 = rt[3]
    take_b4 = min(len(b_lst), cap4)
    rows[3].extend([("B", n) for n in b_lst[:take_b4]])
    b_lst = b_lst[take_b4:]

    rem4 = cap4 - len(rows[3])
    if rem4 > 0:
        take_t4 = min(len(t_lst), cap4 - len(rows[3]))
        rows[3].extend([("T", n) for n in t_lst[:take_t4]])
        t_lst = t_lst[take_t4:]

    rem4 = cap4 - len(rows[3])
    if rem4 > 0:
        take_s4 = min(len(s_lst), cap4 - len(rows[3]))
        rows[3].extend([("S", n) for n in s_lst[:take_s4]])
        s_lst = s_lst[take_s4:]

    leftovers["S"] = s_lst
    leftovers["A"] = a_lst
    leftovers["T"] = t_lst
    leftovers["B"] = b_lst

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