# -*- coding: utf-8 -*-
import os
import re
import json
import time
import base64
import logging
import threading
import datetime as dt
from typing import Optional
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("hcss_backend")

app = FastAPI(title="화원교회 찬양대 좌석 배치 API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

SCOPES = [
    "https://www.googleapis.com/auth/drive.readonly",
    "https://www.googleapis.com/auth/spreadsheets.readonly",
]
DRIVE_FILE_NAME = "26년 할렐루야 출석부"
SHEET_NAMES = ["소프라노", "알토", "테너", "베이스"]

BASE_SEATS = {
    1: {
        "S": {"direction": "R2L", "names": ["남현숙","김경미","백시원","박지현","이서연","이은진","고태옥","황진서","장영자","정은희","황재연"]},
        "A": {"direction": "L2R", "names": ["심윤정","박혜숙","서정영","오정민","박소윤","유미영","노인숙","김혜진","임은애"]}
    },
    2: {
        "S": {"direction": "R2L", "names": ["정진순","김은영","조한주","손국희","김서영","김은희","박진희","황주경","김은현","박도희"]},
        "A": {"direction": "L2R", "names": ["우다연","김지영","방민주","김준희","강미화","박미선","김현경","권순예","조경화","장형미"]}
    },
    3: {
        "S": {"direction": "R2L", "names": ["김지은","노희령","류은채","박유림","윤선미"]},
        "T": {"direction": "R2L", "names": ["강종훈","권재훈","하효동","권영훈","김기형"]},
        "B": {"direction": "L2R", "names": ["이재관","정동근","목현성","전병대","성준호","권상대b","이입교"]}
    },
    4: {
        "T": {"direction": "R2L", "names": ["시진규","남명호","이병륜","김동근","문지민","김경남","박상훈","권상대a"]},
        "B": {"direction": "L2R_CENTER", "names": ["김대성","김기훈","이의헌","최한열","장현재","박선일","박동민","이종성"]}
    },
}

def base_priority(row, part):
    cfg = BASE_SEATS.get(row, {}).get(part)
    if not cfg: 
        return []
    names = list(cfg["names"])
    direction = cfg.get("direction", "L2R")

    # 4열 베이스: 양쪽 끝에서 시작하여 중앙으로 채우는 순서
    if direction == "L2R_CENTER":
        out = []
        left = 0
        right = len(names) - 1
        while left <= right:
            out.append(names[left])
            left += 1
            if left <= right:
                out.append(names[right])
                right -= 1
        return out
    
    # R2L(오른쪽->왼쪽)은 화면상(왼쪽->오른쪽) 역순으로 정렬 출력
    if direction == "R2L":
        return list(reversed(names))
        
    return names

ALTO_OUT_PRIORITY = ["박소윤", "오정민", "박혜숙", "심윤정", "박미선", "강미화", "유미영", "배수련"]
SOPRANO_ROW3_PRIORITY = ["류은채", "노희령", "김지은"]

_cache_lock = threading.Lock()
_cache = {
    "drive_service": None,
    "sheets_service": None,
    "file_id": None,
    "file_id_cached_at": 0.0,
    "sheet_values": None,
    "attendance_cache": {},
}

FILE_ID_TTL = 3600
ATTENDANCE_CACHE_TTL = 300

def get_google_credentials():
    b64_str = os.environ.get("SERVICE_ACCOUNT_BASE64", "").strip()
    if not b64_str:
        raise ValueError("SERVICE_ACCOUNT_BASE64 환경변수가 설정되지 않았습니다.")
    decoded_json = base64.b64decode(b64_str).decode("utf-8")
    return Credentials.from_service_account_info(json.loads(decoded_json), scopes=SCOPES)

def get_drive_service():
    with _cache_lock:
        if _cache["drive_service"] is not None:
            return _cache["drive_service"]
    creds = get_google_credentials()
    service = build("drive", "v3", credentials=creds, cache_discovery=False)
    with _cache_lock:
        _cache["drive_service"] = service
    return service

def get_sheets_service():
    with _cache_lock:
        if _cache["sheets_service"] is not None:
            return _cache["sheets_service"]
    creds = get_google_credentials()
    service = build("sheets", "v4", credentials=creds, cache_discovery=False)
    with _cache_lock:
        _cache["sheets_service"] = service
    return service

def find_drive_file(service):
    escaped_name = DRIVE_FILE_NAME.replace("'", "\\'")
    q = f"name = '{escaped_name}' and trashed = false"
    r = service.files().list(
        q=q, spaces="drive",
        fields="files(id,name,modifiedTime)",
        orderBy="modifiedTime desc"
    ).execute()
    fs = r.get("files", [])
    if not fs:
        q_old = "name = '할렐루야 출석부' and trashed = false"
        r_old = service.files().list(
            q=q_old, spaces="drive",
            fields="files(id,name,modifiedTime)",
            orderBy="modifiedTime desc"
        ).execute()
        fs_old = r_old.get("files", [])
        if not fs_old:
            raise FileNotFoundError(f"구글 드라이브에서 '{DRIVE_FILE_NAME}' 파일을 찾지 못했습니다.")
        return fs_old[0]
    return fs[0]

def get_cached_file_id(drive_service):
    now = time.time()
    with _cache_lock:
        cached = _cache["file_id"]
        cached_at = _cache["file_id_cached_at"]
    if cached and (now - cached_at < FILE_ID_TTL):
        return cached
    file_info = find_drive_file(drive_service)
    with _cache_lock:
        _cache["file_id"] = file_info["id"]
        _cache["file_id_cached_at"] = now
    return file_info["id"]

def fetch_sheet_values(sheets_service, spreadsheet_id):
    ranges = [f"'{name}'!A1:ZZ500" for name in SHEET_NAMES]
    result = sheets_service.spreadsheets().values().batchGet(
        spreadsheetId=spreadsheet_id, ranges=ranges
    ).execute()
    value_ranges = result.get("valueRanges", [])
    sheet_values = {}
    for name, vr in zip(SHEET_NAMES, value_ranges):
        sheet_values[name] = vr.get("values", [])
    return sheet_values

def get_sheet_data(sheets_service, spreadsheet_id, force_refresh=False):
    if not force_refresh:
        with _cache_lock:
            if _cache["sheet_values"] is not None:
                return _cache["sheet_values"]

    sheet_values = fetch_sheet_values(sheets_service, spreadsheet_id)
    with _cache_lock:
        _cache["sheet_values"] = sheet_values
        _cache["attendance_cache"].clear()
    return sheet_values

def get_cached_attendance(sheet_values, date_str, target_date):
    now = time.time()
    key = (date_str,)
    with _cache_lock:
        entry = _cache["attendance_cache"].get(key)
        if entry and (now - entry["cached_at"] < ATTENDANCE_CACHE_TTL):
            return entry["data"]

    people = read_attendance(sheet_values, target_date)
    with _cache_lock:
        _cache["attendance_cache"][key] = {
            "data": people,
            "cached_at": now,
        }
    return people

def clear_all_cache(force_file_lookup=False):
    with _cache_lock:
        _cache["file_id"] = None
        _cache["file_id_cached_at"] = 0.0
        _cache["sheet_values"] = None
        _cache["attendance_cache"].clear()

def parse_cell_date(val, target_year):
    if val is None: return None
    if isinstance(val, dt.datetime): return val.date()
    if isinstance(val, dt.date): return val
    if isinstance(val, (int, float)):
        try: return (dt.datetime(1899, 12, 30) + dt.timedelta(days=val)).date()
        except: pass
    val_str = str(val).strip()
    if not val_str: return None
    val_str_compact = re.sub(r"\s*([.\-/])\s*", r"\1", val_str)
    for candidate in (val_str_compact, val_str):
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%Y.%m.%d", "%Y/%m/%d", "%Y%m%d"):
            try: return dt.datetime.strptime(candidate, fmt).date()
            except ValueError: pass
    match = re.search(r"(\d{1,2})[\/\.\-\s월]+(\d{1,2})", val_str_compact)
    if match:
        try: return dt.date(target_year, int(match.group(1)), int(match.group(2)))
        except ValueError: pass
    return None

def read_attendance(sheet_values, target_date):
    sheet_map = {"소프라노": "S", "알토": "A", "테너": "T", "베이스": "B"}
    people = {p: [] for p in "SATB"}
    target_year = target_date.year
    found_date_col = False

    for sheet_name, part_code in sheet_map.items():
        rows = sheet_values.get(sheet_name)
        if not rows: continue

        date_col = None
        for header_row in rows[:2]:
            for col_idx, cell_val in enumerate(header_row, start=1):
                parsed_d = parse_cell_date(cell_val, target_year)
                if parsed_d and parsed_d == target_date:
                    date_col = col_idx
                    break
            if date_col: break
        if not date_col: continue

        found_date_col = True

        for row_vals in rows[2:]:
            name_val = row_vals[1] if len(row_vals) > 1 else None
            if not name_val: continue
            name_str = str(name_val).strip()
            if not name_str: continue
            if name_str in ("구분", "월통계", "총인원") or "통계" in name_str: break

            c_val = row_vals[date_col - 1] if len(row_vals) >= date_col else None
            is_present = False
            if isinstance(c_val, bool): is_present = c_val
            elif isinstance(c_val, str):
                if c_val.strip().upper() in ("TRUE", "O", "○", "ㅇ", "출석", "참석", "Y", "YES", "1", "●"):
                    is_present = True
            elif isinstance(c_val, (int, float)) and c_val == 1: is_present = True

            if is_present: people[part_code].append(name_str)

    if not found_date_col:
        raise ValueError(f"{target_date.isoformat()} 날짜의 출석 데이터를 찾을 수 없습니다.")

    return people

def row_targets(total):
    caps = [18, 18, 17, 16]
    target_total = min(max(1, int(total)), sum(caps))
    candidates = []
    for r_equal in range(0, min(caps[0], caps[1]) + 1):
        r1 = r2 = r_equal
        for r3 in range(0, caps[2] + 1):
            if r2 < r3 or (r2 - r3) > 1: continue
            r4 = target_total - r1 - r2 - r3
            if 0 <= r4 <= caps[3]:
                if r3 < r4 or (r3 - r4) > 1: continue
                penalty = (r2 - r3) * 10 + (r3 - r4) * 10
                candidates.append((penalty, [r1, r2, r3, r4]))

    if not candidates:
        for r_equal in range(min(caps[0], caps[1]), -1, -1):
            r1 = r2 = r_equal
            for r3 in range(min(r2, caps[2]), -1, -1):
                r4 = target_total - r1 - r2 - r3
                if 0 <= r4 <= caps[3] and r3 >= r4:
                    penalty = (r2 - r3) * 100 + (r3 - r4) * 10
                    candidates.append((penalty, [r1, r2, r3, r4]))

    if candidates:
        return min(candidates, key=lambda x: x[0])[1]

    remaining = target_total
    result = []
    for cap in caps:
        take = min(cap, remaining)
        result.append(take)
        remaining -= take
    return result

def plan_front_row_size(female_count, male_count, fixed_row3_count, cap_front):
    best = None
    max_r = min(cap_front, female_count // 2)
    for r in range(0, max_r + 1):
        leftover_f = female_count - 2 * r
        total34 = leftover_f + male_count + fixed_row3_count
        row3 = (total34 + 1) // 2
        row4 = total34 // 2
        min_row3 = leftover_f + fixed_row3_count
        if row3 < min_row3:
            row3 = min_row3
            row4 = total34 - row3

        if row4 < row3:
            row3_disp, row4_disp = row3, row4 + 1
        else:
            row3_disp, row4_disp = row3 + 1, row4

        d23 = r - row3_disp
        d34 = row3_disp - row4_disp

        order_penalty = 0
        if d23 < 0: order_penalty += 1000 + abs(d23) * 50
        if d34 < 0: order_penalty += 1000 + abs(d34) * 50

        excess23 = max(0, d23 - 1)
        excess34 = max(0, d34 - 1)
        dev_penalty = (excess23 ** 2 + excess34 ** 2) * 20

        penalty = order_penalty + dev_penalty
        cand = (penalty, -r)
        if best is None or cand < best[0]:
            best = (cand, r)
    return best[1] if best else 0

def allocate(people_dict, total_seats):
    people = {p: list(people_dict[p]) for p in ("S", "A", "T", "B")}

    male_count = len(people["T"]) + len(people["B"])
    rt_config = row_targets(total_seats)
    organ_row_config = 3 if male_count < rt_config[3] else 2
    caps_config = rt_config[:]
    if caps_config[organ_row_config] > 0: caps_config[organ_row_config] -= 1
    total_capacity_config = sum(caps_config)
    total_attending = sum(len(people[p]) for p in ("S", "A", "T", "B"))

    leftover = {"S": [], "A": [], "T": [], "B": []}

    if total_attending > total_capacity_config:
        excess_count = total_attending - total_capacity_config
        alto_out_candidates = [m for m in ALTO_OUT_PRIORITY if m in people["A"]]
        for m in people["A"]:
            if m not in alto_out_candidates: alto_out_candidates.append(m)
        moved_alto = []
        while excess_count > 0 and alto_out_candidates:
            m = alto_out_candidates.pop(0)
            moved_alto.append(m)
            people["A"].remove(m)
            excess_count -= 1
        leftover["A"] = moved_alto

    rows = [[], [], [], []]

    name_to_part = {}
    for m in people["A"]: name_to_part[m] = "A"
    for m in people["S"]: name_to_part[m] = "S"

    _seen_f = set()
    female_pool = []
    for m in (base_priority(1,"A") + base_priority(1,"S") + base_priority(2,"A") + base_priority(2,"S")):
        if m in name_to_part and m not in _seen_f:
            female_pool.append((name_to_part[m], m))
            _seen_f.add(m)
    for m in people["A"] + people["S"]:
        if m not in _seen_f:
            female_pool.append((name_to_part[m], m))
            _seen_f.add(m)

    people["A"] = []
    people["S"] = []

    female_count = len(female_pool)
    male_count_remaining = len(people["T"]) + len(people["B"])
    r = plan_front_row_size(female_count, male_count_remaining, 0, cap_front=18)

    def _split_by_named_list(list1_pool, list2_pool, unlisted, cap1, cap2):
        taken1, excess1 = list1_pool[:cap1], list1_pool[cap1:]
        taken2, excess2 = list2_pool[:cap2], list2_pool[cap2:]
        fill = excess1 + excess2 + unlisted
        need1 = cap1 - len(taken1)
        if need1 > 0:
            taken1 = taken1 + fill[:need1]
            fill = fill[need1:]
        need2 = cap2 - len(taken2)
        if need2 > 0:
            taken2 = taken2 + fill[:need2]
            fill = fill[need2:]
        return taken1, taken2, fill

    part_of = dict((n, p) for p, n in female_pool)
    all_alto = [n for p, n in female_pool if p == "A"]
    all_soprano = [n for p, n in female_pool if p == "S"]

    alto_row1_pool = [n for n in base_priority(1,"A") if n in all_alto]
    alto_row2_pool = [n for n in base_priority(2,"A") if n in all_alto]
    alto_listed = set(alto_row1_pool) | set(alto_row2_pool)
    alto_unlisted = [n for n in all_alto if n not in alto_listed]

    alto_total = len(all_alto)
    alto_target = min(alto_total, 2 * r)
    if len(alto_row1_pool) >= len(alto_row2_pool):
        cap1_alto = min(r, (alto_target + 1) // 2)
        cap2_alto = min(r, alto_target // 2)
    else:
        cap1_alto = min(r, alto_target // 2)
        cap2_alto = min(r, (alto_target + 1) // 2)

    alto_row1, alto_row2, alto_overflow = _split_by_named_list(
        alto_row1_pool, alto_row2_pool, alto_unlisted, cap1_alto, cap2_alto)

    s_row1_pool = [n for n in base_priority(1,"S") if n in all_soprano]
    s_row2_pool = [n for n in base_priority(2,"S") if n in all_soprano]
    s_listed = set(s_row1_pool) | set(s_row2_pool)
    s_unlisted = [n for n in all_soprano if n not in s_listed]

    rem1, rem2 = r - len(alto_row1), r - len(alto_row2)
    soprano_row1, soprano_row2, soprano_overflow = _split_by_named_list(
        s_row1_pool, s_row2_pool, s_unlisted, rem1, rem2)

    female_overflow_names = alto_overflow + soprano_overflow
    female_overflow = [(part_of[n], n) for n in female_overflow_names]

    rows[0] = [("A", n) for n in alto_row1] + [("S", n) for n in soprano_row1]
    rows[1] = [("A", n) for n in alto_row2] + [("S", n) for n in soprano_row2]

    for part, name in female_overflow:
        rows[2].append((part, name))

    men_order = []
    seen_men = set()
    for row_num in (3, 4):
        for p in ("S", "T", "B"):
            for name in base_priority(row_num, p):
                if name in seen_men: continue
                if name in people[p]:
                    men_order.append((p, name)); seen_men.add(name)
    for p in ("T", "B"):
        for name in people[p]:
            if name not in seen_men:
                men_order.append((p, name))
                seen_men.add(name)

    people["T"], people["B"] = [], []
    for part, name in men_order:
        if len(rows[2]) <= len(rows[3]):
            rows[2].append((part, name))
        else:
            rows[3].append((part, name))

    def _pop_last_n(row_list, part, n):
        removed = []
        idx = len(row_list) - 1
        while n > 0 and idx >= 0:
            p, nm = row_list[idx]
            if p == part:
                removed.append(row_list.pop(idx))
                n -= 1
            idx -= 1
        removed.reverse()
        return removed

    a1_count = sum(1 for p, _ in rows[0] if p == "A")
    a2_count = sum(1 for p, _ in rows[1] if p == "A")
    b3_count = sum(1 for p, _ in rows[2] if p == "B")
    a3_count = sum(1 for p, _ in rows[2] if p == "A")
    s3_count = sum(1 for p, _ in rows[2] if p == "S")

    avg_a12 = (a1_count + a2_count) / 2
    offset = round(avg_a12 - (b3_count + a3_count))

    if offset >= 3:
        move = min(offset, a1_count, a2_count, s3_count, 2)
        move1 = (move + 1) // 2
        move2 = move - move1

        moved_alto = _pop_last_n(rows[0], "A", move1) + _pop_last_n(rows[1], "A", move2)
        moved_soprano = _pop_last_n(rows[2], "S", len(moved_alto))

        rows[2].extend(moved_alto)
        for i, item in enumerate(moved_soprano):
            (rows[0] if i < move1 else rows[1]).append(item)

    if people["A"]:
        leftover["A"].extend(people["A"])
        people["A"] = []
    if people["S"]:
        leftover["S"].extend(people["S"])
        people["S"] = []

    real_row3, real_row4 = len(rows[2]), len(rows[3])
    if real_row4 < real_row3:
        organ_row = 3
    else:
        organ_row = 2

    rt = [len(rows[0]), len(rows[1]), real_row3, real_row4]

    # 물리적 좌→우 배치 파트 순서
    row_display_order = {
        0: ["A", "S"],
        1: ["A", "S"],
        2: ["B", "A", "S", "T", "ORG"],
        3: ["B", "T", "ORG"],
    }

    final_rows, row_part_counts = [], []
    for r_idx in range(4):
        curr_row = list(rows[r_idx])
        if r_idx == organ_row:
            curr_row.append(("ORG", "오르간"))

        part_order = row_display_order[r_idx]
        part_grouped = {p: [] for p in part_order}
        for part, name in curr_row:
            part_grouped[part].append(name)

        sorted_row = []
        for p in part_order:
            members = part_grouped[p]
            for m in members:
                sorted_row.append((p, m))

        final_rows.append(sorted_row)
        p_counts = {p: len(part_grouped.get(p, [])) for p in ("S", "A", "T", "B", "ORG")}
        row_part_counts.append(p_counts)

    return final_rows, leftover, rt, organ_row, row_part_counts

class RequestModel(BaseModel):
    date_str: str
    total_seats: int = Field(default=67, ge=1, description="전체 좌석 정원")
    balance_rule: str = "1=2>=3>=4"
    force_refresh: bool = Field(default=False, description="출석부 캐시를 무시하고 새로 불러올지 여부")

@app.get("/health")
def health_check():
    return {"status": "ok", "code_version": "2026-09-15-seat-direction-final-fix2"}

@app.post("/api/allocate")
def run_allocation(req: RequestModel):
    try:
        target = dt.date.fromisoformat(req.date_str)
        sheets_service = get_sheets_service()

        if req.force_refresh:
            clear_all_cache(force_file_lookup=False)

        drive_service = get_drive_service()
        try:
            file_id = get_cached_file_id(drive_service)
        except Exception:
            clear_all_cache(force_file_lookup=True)
            file_id = get_cached_file_id(drive_service)

        sheet_values = get_sheet_data(
            sheets_service,
            file_id,
            force_refresh=req.force_refresh,
        )

        people = get_cached_attendance(sheet_values, req.date_str, target)

        attending_by_part = {p: len(people[p]) for p in ("S", "A", "T", "B")}
        rows, leftover, rt, org, row_part_counts = allocate(
            people, req.total_seats
        )

        attending_total = sum(attending_by_part.values())
        recommended_max_seats = min(
            attending_total + 1, sum([18, 18, 17, 16])
        )
        recommended_min_seats = min(60, recommended_max_seats)

        return {
            "status": "success",
            "code_version": "2026-09-15-seat-direction-final-fix2",
            "date": str(target),
            "attending_count": attending_total,
            "attending_by_part": attending_by_part,
            "row_targets": rt,
            "organ_row": org + 1,
            "rows": rows,
            "row_part_counts": row_part_counts,
            "sungdoseok": leftover,
            "recommended_total_seats": recommended_max_seats,
            "recommended_min_seats": recommended_min_seats,
            "recommended_max_seats": recommended_max_seats,
            "data_source": "Google Sheets API",
            "refresh_mode": "force" if req.force_refresh else "cache",
        }
    except Exception as e:
        logger.error(f"오류 발생: {str(e)}")
        raise HTTPException(status_code=400, detail=str(e))