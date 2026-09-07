# -*- coding: utf-8 -*-
import os
import io
import re
import json
import time
import base64
import logging
import threading
import datetime as dt
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from openpyxl import load_workbook

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("hcss_backend")

app = FastAPI(title="화원교회 찬양대 좌석 배치 API")

# 인증서를 쓰지 않으므로 credentials=True + origins=* 조합(스펙 위반)을 피한다.
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]
DRIVE_FILE_NAME = "26년 할렐루야 출석부.xlsx"

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

ALTO_OUT_PRIORITY = ["박소윤", "오정민", "박혜숙", "심윤정", "박미선", "강미화", "유미영", "배수련"]
SOPRANO_ROW3_PRIORITY = ["류은채", "노희령", "김지은"]

# ---------------------------------------------------------------------------
# 캐시 (구글 API 재호출 및 재다운로드로 인한 지연을 줄이기 위함)
# ---------------------------------------------------------------------------
_cache_lock = threading.Lock()
_cache = {
    "service": None,            # 구글 드라이브 서비스 객체 (재사용)
    "file_id": None,            # 검색된 엑셀 파일 ID
    "file_id_cached_at": 0.0,
    "file_mtime": None,         # 마지막으로 다운로드했을 때의 수정시간
    "file_bytes": None,         # 마지막으로 다운로드한 엑셀 바이트
    "attendance_cache": {},     # (mtime, date_str) -> {"data": {...}, "cached_at": ts}
}

FILE_ID_TTL = 3600          # 파일 검색(전체 드라이브 탐색) 결과 캐시 유지 시간(초)
ATTENDANCE_CACHE_TTL = 300  # 동일 (파일버전, 날짜) 파싱 결과 캐시 유지 시간(초)


def get_google_credentials():
    b64_str = os.environ.get("SERVICE_ACCOUNT_BASE64", "").strip()
    if not b64_str:
        raise ValueError("SERVICE_ACCOUNT_BASE64 환경변수가 설정되지 않았습니다.")
    decoded_json = base64.b64decode(b64_str).decode("utf-8")
    return Credentials.from_service_account_info(json.loads(decoded_json), scopes=SCOPES)


def get_drive_service():
    """구글 드라이브 서비스 객체를 프로세스 내에서 재사용한다.

    build()에 cache_discovery=False를 주지 않으면 매 호출마다 디스커버리 문서를
    로컬 파일 캐시에서 읽고 쓰려고 시도하면서 지연이 발생할 수 있다
    (특히 Render처럼 파일시스템이 제한된 환경에서 체감이 크다).
    """
    with _cache_lock:
        if _cache["service"] is not None:
            return _cache["service"]
    creds = get_google_credentials()
    service = build("drive", "v3", credentials=creds, cache_discovery=False)
    with _cache_lock:
        _cache["service"] = service
    return service


def find_drive_file(service):
    escaped_name = DRIVE_FILE_NAME.replace("'", "\\'")
    q = f"name = '{escaped_name}' and trashed = false"
    r = service.files().list(q=q, spaces="drive", fields="files(id,name)").execute()
    fs = r.get("files", [])
    if not fs:
        q_old = "name = '할렐루야 출석부.xlsx' and trashed = false"
        r_old = service.files().list(q=q_old, spaces="drive", fields="files(id,name)").execute()
        fs_old = r_old.get("files", [])
        if not fs_old:
            raise FileNotFoundError(f"구글 드라이브에서 '{DRIVE_FILE_NAME}' 파일을 찾지 못했습니다.")
        return fs_old[0]
    return fs[0]


def get_cached_file_id(service):
    now = time.time()
    with _cache_lock:
        cached = _cache["file_id"]
        cached_at = _cache["file_id_cached_at"]
    if cached and (now - cached_at < FILE_ID_TTL):
        return cached
    file_info = find_drive_file(service)
    with _cache_lock:
        _cache["file_id"] = file_info["id"]
        _cache["file_id_cached_at"] = now
    return file_info["id"]


def download_excel_to_stream(service, file_id):
    request = service.files().get_media(fileId=file_id)
    file_stream = io.BytesIO()
    downloader = MediaIoBaseDownload(file_stream, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    file_stream.seek(0)
    return file_stream


def get_excel_bytes(service, file_id):
    """수정시간(modifiedTime)만 가볍게 확인해서, 파일이 바뀌지 않았다면
    다시 다운로드하지 않고 캐시된 바이트를 재사용한다."""
    meta = service.files().get(fileId=file_id, fields="modifiedTime").execute()
    mtime = meta.get("modifiedTime")
    with _cache_lock:
        if _cache["file_bytes"] is not None and _cache["file_mtime"] == mtime:
            return _cache["file_bytes"], mtime
    stream = download_excel_to_stream(service, file_id)
    file_bytes = stream.getvalue()
    with _cache_lock:
        _cache["file_bytes"] = file_bytes
        _cache["file_mtime"] = mtime
    return file_bytes, mtime


def get_cached_attendance(file_bytes, mtime, date_str, target_date):
    now = time.time()
    key = (mtime, date_str)
    with _cache_lock:
        entry = _cache["attendance_cache"].get(key)
        if entry and (now - entry["cached_at"] < ATTENDANCE_CACHE_TTL):
            return entry["data"]
    people = read_attendance(io.BytesIO(file_bytes), target_date)
    with _cache_lock:
        # 오래된 캐시 정리 (메모리 누수 방지)
        expired = [k for k, v in _cache["attendance_cache"].items()
                   if now - v["cached_at"] > ATTENDANCE_CACHE_TTL]
        for k in expired:
            del _cache["attendance_cache"][k]
        _cache["attendance_cache"][key] = {"data": people, "cached_at": now}
    return people


def clear_all_cache():
    with _cache_lock:
        _cache["file_id"] = None
        _cache["file_id_cached_at"] = 0.0
        _cache["file_mtime"] = None
        _cache["file_bytes"] = None
        _cache["attendance_cache"].clear()


# ---------------------------------------------------------------------------
# 날짜/출석 파싱
# ---------------------------------------------------------------------------
def parse_cell_date(val, target_year):
    if val is None: return None
    if isinstance(val, dt.datetime): return val.date()
    if isinstance(val, dt.date): return val
    if isinstance(val, (int, float)):
        try: return (dt.datetime(1899, 12, 30) + dt.timedelta(days=val)).date()
        except: pass
    val_str = str(val).strip()
    if not val_str: return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%Y.%m.%d", "%Y/%m/%d", "%Y%m%d"):
        try: return dt.datetime.strptime(val_str, fmt).date()
        except ValueError: pass
    match = re.search(r"(\d{1,2})[\/\.\-\s월]+(\d{1,2})", val_str)
    if match:
        try: return dt.date(target_year, int(match.group(1)), int(match.group(2)))
        except ValueError: pass
    return None


def read_attendance(file_stream, target_date):
    wb = load_workbook(file_stream, data_only=True)
    sheet_map = {"소프라노": "S", "알토": "A", "테너": "T", "베이스": "B"}
    people = {p: [] for p in "SATB"}
    target_year = target_date.year

    for sheet_name, part_code in sheet_map.items():
        if sheet_name not in wb.sheetnames: continue
        ws = wb[sheet_name]
        date_col = None
        for col in range(1, ws.max_column + 1):
            cell_val = ws.cell(row=2, column=col).value or ws.cell(row=1, column=col).value
            parsed_d = parse_cell_date(cell_val, target_year)
            if parsed_d and parsed_d == target_date:
                date_col = col
                break
        if not date_col: continue

        for row in range(3, ws.max_row + 1):
            name_val = ws.cell(row=row, column=2).value
            if not name_val: continue
            name_str = str(name_val).strip()
            if name_str in ("구분", "월통계", "총인원") or "통계" in name_str: break

            c_val = ws.cell(row=row, column=date_col).value
            is_present = False
            if c_val is True: is_present = True
            elif isinstance(c_val, str):
                if c_val.strip().upper() in ("TRUE", "O", "○", "ㅇ", "출석", "참석", "Y", "YES", "1", "●"):
                    is_present = True
            elif isinstance(c_val, (int, float)) and c_val == 1: is_present = True

            if is_present: people[part_code].append(name_str)

    return people


# ---------------------------------------------------------------------------
# 좌석 배치 로직
# ---------------------------------------------------------------------------
def row_targets(total):
    caps = [18, 18, 17, 16]
    target_total = min(max(1, int(total)), sum(caps))

    candidates = []
    # 1열 = 2열 >= 3열 >= 4열 조건을 엄격하게 적용 (열 간 편차가 1 이내)
    for r_equal in range(0, min(caps[0], caps[1]) + 1):
        r1 = r2 = r_equal
        for r3 in range(0, caps[2] + 1):
            if r2 < r3 or (r2 - r3) > 1: continue  # 편차 1 이내
            r4 = target_total - r1 - r2 - r3
            if 0 <= r4 <= caps[3]:
                if r3 < r4 or (r3 - r4) > 1: continue  # 편차 1 이내
                penalty = (r2 - r3) * 10 + (r3 - r4) * 10
                candidates.append((penalty, [r1, r2, r3, r4]))

    if not candidates:
        # 조건이 완화되는 예외 경우 백업 탐색
        for r_equal in range(min(caps[0], caps[1]), -1, -1):
            r1 = r2 = r_equal
            for r3 in range(min(r2, caps[2]), -1, -1):
                r4 = target_total - r1 - r2 - r3
                if 0 <= r4 <= caps[3] and r3 >= r4:
                    penalty = (r2 - r3) * 100 + (r3 - r4) * 10
                    candidates.append((penalty, [r1, r2, r3, r4]))

    if candidates:
        return min(candidates, key=lambda x: x[0])[1]

    # 정원이 매우 적어(예: 1~3명) 위 조건들을 만족하는 조합이 전혀 없는 경우의
    # 최종 안전망: 1열부터 순서대로 채운다. (편차 규칙보다 "누락 없이 배치"가 우선)
    remaining = target_total
    result = []
    for cap in caps:
        take = min(cap, remaining)
        result.append(take)
        remaining -= take
    return result


def allocate(people_dict, total_seats):
    people = {p: list(people_dict[p]) for p in ("S", "A", "T", "B")}
    rt = row_targets(total_seats)

    male_count = len(people["T"]) + len(people["B"])
    organ_row = 3 if male_count < rt[3] else 2

    caps = rt[:]
    if caps[organ_row] > 0: caps[organ_row] -= 1

    leftover = {"S": [], "A": [], "T": [], "B": []}
    total_capacity = sum(caps)
    total_attending = sum(len(people[p]) for p in ("S", "A", "T", "B"))

    if total_attending > total_capacity:
        excess_count = total_attending - total_capacity
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

    s_row3_fixed = [m for m in SOPRANO_ROW3_PRIORITY if m in people["S"]]
    for m in s_row3_fixed: people["S"].remove(m)

    female_a = [m for m in BASE_SEAT_ORDER[1] + BASE_SEAT_ORDER[2] if m in people["A"]]
    for m in people["A"]:
        if m not in female_a: female_a.append(m)

    female_s = [m for m in BASE_SEAT_ORDER[1] + BASE_SEAT_ORDER[2] if m in people["S"]]
    for m in people["S"]:
        if m not in female_s: female_s.append(m)

    max_per_row = caps[0]
    total_females = len(female_a) + len(female_s)
    front_per_row = min(max_per_row, total_females // 2)

    count_a_per_row = min(len(female_a) // 2, front_per_row)

    rows[0] = [("A", m) for m in female_a[:count_a_per_row]]
    rows[1] = [("A", m) for m in female_a[count_a_per_row:count_a_per_row * 2]]

    unplaced_a = female_a[count_a_per_row * 2:]
    need_s = front_per_row - len(rows[0])
    rows[0] += [("S", m) for m in female_s[:need_s]]
    rows[1] += [("S", m) for m in female_s[need_s:need_s * 2]]

    # 1,2열 배치 후 남은 인원 (알토는 기존처럼 people["A"]에 남겨 3열 배정 단계에서 처리)
    people["A"] = unplaced_a
    leftover_s = female_s[need_s * 2:]

    # 3열 배정: 소프라노는 1,2,3열 안에서 반드시 소화하고 성도석으로 보내지 않는다.
    # 정원(caps[2])을 넘더라도 소프라노는 전원 3열에 배치한다.
    for name in s_row3_fixed:
        rows[2].append(("S", name))
    for name in leftover_s:
        rows[2].append(("S", name))

    # 3열 나머지 자리는 좌석 배치도 순서(BASE_SEAT_ORDER)를 따라 A/T/B 배정
    # (소프라노가 이미 정원을 채웠다면 이 구간은 자연히 건너뛰게 된다)
    for name in BASE_SEAT_ORDER[3]:
        if len(rows[2]) >= caps[2]: break
        if name in s_row3_fixed: continue
        for p in ("A", "T", "B"):
            if name in people[p]:
                rows[2].append((p, name))
                people[p].remove(name)
                break

    while len(rows[2]) < caps[2]:
        cand = [p for p in ("A", "T", "B") if people[p]]
        if not cand: break
        rows[2].append((cand[0], people[cand[0]].pop(0)))

    # 4열 배정
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
        rows[3].append((cand[0], people[cand[0]].pop(0)))

    # 성도석(leftover)은 알토 파트만 배정한다.
    # 테너/베이스는 소프라노와 마찬가지로 정원을 넘더라도 성도석으로 보내지 않고
    # 3,4열 안에서만(다른 열에는 배치하지 않고) 배치한다. 단, 4열에만 몰아넣지 않고
    # 매번 인원이 더 적은 열(3열 vs 4열)에 채워서 두 열 간 편차가 1 이하로 유지되게 한다.
    extra_men = [("T", n) for n in people["T"]] + [("B", n) for n in people["B"]]
    people["T"], people["B"] = [], []
    for part, name in extra_men:
        if len(rows[2]) <= len(rows[3]):
            rows[2].append((part, name))
        else:
            rows[3].append((part, name))

    # 안전망: 소프라노가 3열을 정원 이상으로 채워 알토가 자리를 못 찾았다면 성도석으로 보낸다.
    if people["A"]:
        leftover["A"].extend(people["A"])
        people["A"] = []

    reversed_part_order = ["B", "T", "A", "S", "ORG"]

    final_rows, row_part_counts = [], []
    for r_idx in range(4):
        curr_row = list(rows[r_idx])
        if r_idx == organ_row:
            curr_row.append(("ORG", "오르간"))

        part_grouped = {p: [] for p in reversed_part_order}
        for part, name in curr_row:
            part_grouped[part].append(name)

        sorted_row = []
        for p in reversed_part_order:
            members = list(reversed(part_grouped[p]))
            for m in members:
                sorted_row.append((p, m))

        final_rows.append(sorted_row)
        p_counts = {p: len(part_grouped[p]) for p in ("S", "A", "T", "B", "ORG")}
        row_part_counts.append(p_counts)

    return final_rows, leftover, rt, organ_row, row_part_counts


class RequestModel(BaseModel):
    date_str: str
    total_seats: int = Field(default=67, ge=1, description="전체 좌석 정원")
    balance_rule: str = "1=2>=3>=4"
    force_refresh: bool = Field(default=False, description="출석부 캐시를 무시하고 새로 불러올지 여부")


@app.get("/health")
def health_check():
    """Render 무료 플랜은 일정 시간 요청이 없으면 서버가 잠든다(cold start).
    이 엔드포인트는 구글 드라이브 접근 없이 즉시 응답하므로, 외부 무료 핑 서비스
    (UptimeRobot, cron-job.org 등)로 5~10분 간격 주기 호출을 걸어두면 실제 좌석
    배치 요청이 콜드스타트로 30초~1분씩 걸리는 것을 크게 줄일 수 있다."""
    return {"status": "ok"}


@app.post("/api/allocate")
def run_allocation(req: RequestModel):
    try:
        target = dt.date.fromisoformat(req.date_str)

        if req.force_refresh:
            clear_all_cache()

        service = get_drive_service()
        try:
            file_id = get_cached_file_id(service)
            file_bytes, mtime = get_excel_bytes(service, file_id)
        except Exception:
            # 캐시된 파일 ID가 더 이상 유효하지 않을 수 있으므로(파일 삭제/이름변경 등)
            # 캐시를 비우고 한 번 더 시도한다.
            clear_all_cache()
            file_id = get_cached_file_id(service)
            file_bytes, mtime = get_excel_bytes(service, file_id)

        people = get_cached_attendance(file_bytes, mtime, req.date_str, target)

        attending_by_part = {p: len(people[p]) for p in ("S", "A", "T", "B")}
        rows, leftover, rt, org, row_part_counts = allocate(people, req.total_seats)

        return {
            "status": "success",
            "date": str(target),
            "attending_count": sum(attending_by_part.values()),
            "attending_by_part": attending_by_part,
            "row_targets": rt,
            "organ_row": org + 1,
            "rows": rows,
            "row_part_counts": row_part_counts,
            "sungdoseok": leftover
        }
    except Exception as e:
        logger.error(f"오류 발생: {str(e)}")
        raise HTTPException(status_code=400, detail=str(e))