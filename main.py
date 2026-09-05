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

app = FastAPI(title="화원교회 찬양대 좌석 배치 API (HCSS 3.17)")

# CORS 설정
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]
DRIVE_FILE_NAME = "26년 할렐루야 출석부.xlsx"

# 기본 좌석표
BASE_SEAT_ORDER = {
    1: ["남현숙", "김경미", "백시원", "박지현", "이서연", "이은진", "고태옥", "황진서", "장영자", "김은희", "황재연", "박유림",
        "임은애", "김혜진", "노인숙", "유미영", "박소윤", "오정민", "서정영", "박혜숙", "심윤정"],
    2: ["김은영", "정진순", "조한주", "손국희", "김서영", "김은희", "박진희", "황주경", "김은현", "박도희", "윤선미",
        "장형미", "조경화", "권순예", "김현경", "박미선", "강미화", "김준희", "방민주", "김지영", "우다연"],
    3: ["류은채", "노희령", "김지은", "강종훈", "권재훈", "하효동", "권영훈", "김기형", 
        "이입교", "권상대b", "성준호", "전병대", "목현성", "정동근", "이재관"],
    4: ["시진규", "남명호", "이병륜", "김경남", "김동근", "문지민", "박상훈", "권상대a", 
        "이종성", "박동민", "박선일", "장현재", "최한열", "이의헌", "김기훈", "김대성"]
}

# 알토 성도석 이동 우선순위 목록
ALTO_OUT_PRIORITY = [
    "박소윤", "오정민", "박혜숙", "심윤정", "박미선", "강미화", "유미영", "배수련"
]

# 3열 고정 소프라노 우선순위 목록
SOPRANO_ROW3_PRIORITY = [
    "류은채", "노희령", "김지은"
]

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
        # 혹시 몰라 기존 구 파일명으로도 한 번 더 재시도
        q_old = "name = '할렐루야 출석부.xlsx' and trashed = false"
        r_old = service.files().list(q=q_old, spaces="drive", fields="files(id,name)").execute()
        fs_old = r_old.get("files", [])
        if not fs_old:
            raise FileNotFoundError(f"구글 드라이브에서 '{DRIVE_FILE_NAME}' 파일을 찾지 못했습니다.")
        return fs_old[0]
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

    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%Y.%m.%d", "%Y/%m/%d", "%Y%m%d"):
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
    sheet_map = {"소프라노": "S", "알토": "A", "테너": "T", "베이스": "B"}
    people = {p: [] for p in "SATB"}
    target_year = target_date.year

    for sheet_name, part_code in sheet_map.items():
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        
        date_col = None
        for col in range(1, ws.max_column + 1):
            cell_val = ws.cell(row=2, column=col).value or ws.cell(row=1, column=col).value
            parsed_d = parse_cell_date(cell_val, target_year)
            if parsed_d and parsed_d == target_date:
                date_col = col
                break

        if not date_col:
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
                if v in ("TRUE", "O", "○", "ㅇ", "출석", "참석", "Y", "YES", "1", "●"):
                    is_present = True
            elif isinstance(c_val, (int, float)) and c_val == 1:
                is_present = True

            if is_present:
                people[part_code].append(name_str)

    return people

def row_targets(total):
    caps = [18, 18, 17, 16]
    target_total = min(max(1, int(total)), sum(caps))
    
    candidates = []
    for r_equal in range(0, min(caps[0], caps[1]) + 1):
        r1 = r2 = r_equal
        for r3 in range(0, caps[2] + 1):
            if r2 < r3: continue
            r4 = target_total - r1 - r2 - r3
            if 0 <= r4 <= caps[3]:
                if r3 < r4: continue
                penalty = (r2 - r3) * 100 + (r3 - r4) * 50
                candidates.append((penalty, [r1, r2, r3, r4]))

    if not candidates:
        for r_equal in range(min(caps[0], caps[1]), -1, -1):
            r1 = r2 = r_equal
            rem = target_total - r1 - r2
            if rem >= 0:
                r3 = min(rem, caps[2])
                r4 = rem - r3
                if r4 <= caps[3]:
                    candidates.append((0, [r1, r2, r3, r4]))
                    break

    return min(candidates, key=lambda x: x[0])[1]

# --- 자동 배치 알고리즘 (HCSS 3.17 스펙 완벽 반영) ---
def allocate(people_dict, total_seats):
    people = {p: list(people_dict[p]) for p in ("S", "A", "T", "B")}
    rt = row_targets(total_seats)
    
    male_count = len(people["T"]) + len(people["B"])
    if male_count < rt[3] or (rt[2] < rt[3]):
        organ_row = 3  # 4열로 오르간 이동
    else:
        organ_row = 2  # 기본 3열 배치
    
    caps = rt[:]
    if caps[organ_row] > 0:
        caps[organ_row] -= 1

    leftover = {"S": [], "A": [], "T": [], "B": []}
    
    # 1. 성도석 이동 (오직 알토 파트만 이동 허용)
    total_capacity = sum(caps)
    total_attending = sum(len(people[p]) for p in ("S", "A", "T", "B"))
    
    if total_attending > total_capacity:
        excess_count = total_attending - total_capacity
        
        # 알토 파트 중 우선순위 대상 먼저 선정
        alto_out_candidates = [m for m in ALTO_OUT_PRIORITY if m in people["A"]]
        for m in people["A"]:
            if m not in alto_out_candidates:
                alto_out_candidates.append(m)
                
        moved_alto = []
        while excess_count > 0 and alto_out_candidates:
            m = alto_out_candidates.pop(0)
            moved_alto.append(m)
            people["A"].remove(m)
            excess_count -= 1
            
        leftover["A"] = moved_alto

    rows = [[], [], [], []]

    # 2. 1, 2열 여성(S, A) 동수 배정 (3열 고정 소프라노 제외)
    s_row3_fixed = [m for m in SOPRANO_ROW3_PRIORITY if m in people["S"]]
    for m in s_row3_fixed:
        people["S"].remove(m)

    female_a = [m for m in BASE_SEAT_ORDER[1] + BASE_SEAT_ORDER[2] if m in people["A"]]
    for m in people["A"]:
        if m not in female_a: female_a.append(m)
        
    female_s = [m for m in BASE_SEAT_ORDER[1] + BASE_SEAT_ORDER[2] if m in people["S"]]
    for m in people["S"]:
        if m not in female_s: female_s.append(m)

    max_per_row = caps[0]
    total_females = len(female_a) + len(female_s)
    front_per_row = min(max_per_row, total_females // 2)

    half_a = len(female_a) // 2
    count_a_per_row = min(half_a, front_per_row)
    
    r1_members = [("A", m) for m in female_a[:count_a_per_row]]
    r2_members = [("A", m) for m in female_a[count_a_per_row:count_a_per_row * 2]]
    
    unplaced_a = female_a[count_a_per_row * 2:]

    need_s = front_per_row - len(r1_members)
    r1_members += [("S", m) for m in female_s[:need_s]]
    r2_members += [("S", m) for m in female_s[need_s:need_s * 2]]
    
    unplaced_s = female_s[need_s * 2:]

    rows[0] = r1_members
    rows[1] = r2_members

    people["A"] = unplaced_a
    people["S"] = unplaced_s

    # 3. 3열 배정 (혼성 - 고정 소프라노, 남은 여성 및 남성 단원)
    for name in s_row3_fixed:
        if len(rows[2]) >= caps[2]: break
        rows[2].append(("S", name))

    for name in BASE_SEAT_ORDER[3]:
        if len(rows[2]) >= caps[2]: break
        if name in s_row3_fixed: continue
        for p in ("S", "A", "T", "B"):
            if name in people[p]:
                rows[2].append((p, name))
                people[p].remove(name)
                break

    while len(rows[2]) < caps[2]:
        cand = [p for p in ("S", "A", "T", "B") if people[p]]
        if not cand: break
        p = cand[0]
        rows[2].append((p, people[p].pop(0)))

    # 4. 4열 배정 (테너 T, 베이스 B 파트만 배정)
    for name in BASE_SEAT_ORDER[4]:
        if len(rows[3]) >= caps[3]: break
        for p in ("T", "B"):
            if name in people[p]:
                rows[3].append((p, name))
                people[p].remove(name)
                break

    while len(rows[3]) < caps[3]:
        cand = [p for p in ("T", "B") if people[p]]
        if not cand: break
        p = cand[0]
        rows[3].append((p, people[p].pop(0)))

    # 파트별 및 오르간 위치 정렬
    final_rows = []
    part_order = {"ORG": 0, "S": 1, "A": 2, "T": 3, "B": 4}
    
    for r_idx in range(4):
        curr_row = list(rows[r_idx])
        if r_idx == organ_row:
            curr_row.append(("ORG", "오르간"))
        
        sorted_row = sorted(curr_row, key=lambda x: part_order[x[0]])
        final_rows.append(sorted_row)

    return final_rows, leftover, rt, organ_row

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
            "organ_row": org + 1,  # 프론트엔드 표출용 1-based index (3열 혹은 4열)
            "rows": rows,
            "leftover": leftover
        }
    except Exception as e:
        logger.error(f"오류 발생: {str(e)}")
        raise HTTPException(status_code=400, detail=str(e))