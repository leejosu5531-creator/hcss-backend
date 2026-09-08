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
    1: ["남현숙", "김경미", "백시원", "박지현", "이서연", "이은진", "고태옥", "황진서", "장영자", "황재연", "박유림",
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


def plan_front_row_size(female_count, male_count, fixed_row3_count, cap_front, alto_count):
    """1,2열 크기(r, r=1열=2열)를 탐색해 2,3,4열 편차가 최대한 1 이하가 되도록 정한다.

    1,2열은 여성(알토+소프라노)만 앉을 수 있고, 4열은 남성(테너+베이스)만 앉을 수
    있으며, 3열은 남는 여성 + 남성을 함께 받는다는 물리적 제약을 반영한다.
    단순히 총원을 4등분하면(성비를 무시하면) 실제로는 달성 불가능한 목표가 나올 수
    있어, 실제 여성/남성 인원수를 넣고 후보 r을 모두 시도해 가장 균형 잡힌 조합을 고른다.
    (2열>=3열>=4열 순서를 우선 지키고, 그 안에서 편차 1을 넘는 만큼을 제곱으로 벌점을
    줘서 어느 한 곳에 편차가 몰리지 않도록 한다.)

    알토는 3열에 앉지 않으므로, 밸런스를 이유로 r을 알토 인원의 절반보다 작게 잡으면
    안 된다(그러면 정원이 남는데도 알토가 성도석으로 밀려나게 됨). 따라서 r의 하한을
    ceil(alto_count/2)로 둔다.
    """
    best = None
    min_r = min(cap_front, (alto_count + 1) // 2)  # 알토는 반드시 1,2열에 다 들어가야 함
    max_r = min(cap_front, female_count // 2)
    max_r = max(max_r, min_r)
    for r in range(min_r, max_r + 1):
        leftover_f = female_count - 2 * r
        total34 = leftover_f + male_count + fixed_row3_count
        row3 = (total34 + 1) // 2
        row4 = total34 // 2
        min_row3 = leftover_f + fixed_row3_count  # 이미 확정된 여성 인원보다 3열이 작을 수 없음
        if row3 < min_row3:
            row3 = min_row3
            row4 = total34 - row3

        d23 = r - row3     # 0 또는 1이 이상적 (2열 >= 3열)
        d34 = row3 - row4  # 0 또는 1이 이상적 (3열 >= 4열)

        order_penalty = 0
        if d23 < 0: order_penalty += 1000 + abs(d23) * 50   # 순서 역전(3열>2열)은 강하게 회피
        if d34 < 0: order_penalty += 1000 + abs(d34) * 50   # 순서 역전(4열>3열)은 강하게 회피

        excess23 = max(0, d23 - 1)
        excess34 = max(0, d34 - 1)
        dev_penalty = (excess23 ** 2 + excess34 ** 2) * 20  # 허용 편차(1)를 넘는 만큼 제곱 벌점

        penalty = order_penalty + dev_penalty
        cand = (penalty, -r)  # 동률이면 r이 큰 쪽(앞줄을 더 채우는 쪽) 선호
        if best is None or cand < best[0]:
            best = (cand, r)
    return best[1] if best else 0


def allocate(people_dict, total_seats):
    people = {p: list(people_dict[p]) for p in ("S", "A", "T", "B")}

    # 1단계: 설정된 찬양대석 정원(total_seats) 기준으로 "전체 초과 여부"만 판단한다.
    # (오늘 실제 출석 인원이 정원 자체를 넘는 극단적 경우에만 알토를 성도석으로 보냄)
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

    s_row3_fixed = [m for m in SOPRANO_ROW3_PRIORITY if m in people["S"]]
    for m in s_row3_fixed: people["S"].remove(m)

    # 이름 중복 방지(seen-set): 원본 좌석 순서표에 같은 이름이 두 번 들어있어도
    # 한 자리에만 배치되도록 안전장치를 둔다.
    _seen_a = set()
    female_a = []
    for m in BASE_SEAT_ORDER[1] + BASE_SEAT_ORDER[2]:
        if m in people["A"] and m not in _seen_a:
            female_a.append(m)
            _seen_a.add(m)
    for m in people["A"]:
        if m not in _seen_a:
            female_a.append(m)
            _seen_a.add(m)

    _seen_s = set()
    female_s = []
    for m in BASE_SEAT_ORDER[1] + BASE_SEAT_ORDER[2]:
        if m in people["S"] and m not in _seen_s:
            female_s.append(m)
            _seen_s.add(m)
    for m in people["S"]:
        if m not in _seen_s:
            female_s.append(m)
            _seen_s.add(m)

    # female_a/female_s가 사본을 들고 있으므로 원본은 비워둔다.
    # (비워두지 않으면 이후 최종 안전망에서 이미 배치된 인원이 중복으로 성도석에 추가된다)
    people["A"] = []
    people["S"] = []

    # 2단계: 실제 여성/남성 인원수를 반영해 1열=2열 크기(r)를 다시 산출한다.
    # (설정 정원만으로 총원을 4등분하면 성비에 따라 2,3열 편차가 커질 수 있음)
    female_count = len(female_a) + len(female_s)
    male_count_remaining = len(people["T"]) + len(people["B"])
    r = plan_front_row_size(female_count, male_count_remaining, len(s_row3_fixed), cap_front=18,
                             alto_count=len(female_a))

    # 오르간 배치 열 결정 및 이후 표시용 row_targets: 위에서 정한 r을 반영해 재구성한다.
    leftover_f_preview = max(0, female_count - 2 * r)
    total34_preview = leftover_f_preview + male_count_remaining + len(s_row3_fixed)
    row3_preview = max(leftover_f_preview + len(s_row3_fixed), (total34_preview + 1) // 2)
    row4_preview = total34_preview - row3_preview
    rt = [r, r, row3_preview, row4_preview]
    organ_row = 3 if male_count < rt[3] else 2

    max_row = r  # 1,2열 목표 크기(각 열)

    # 1,2열 배정: 알토를 최우선으로 확보하고(알토는 1,2열에만 앉을 수 있음),
    # 남는 자리는 소프라노로 채운다. 두 열에 한 자리씩 번갈아 배정해서
    # 1열=2열이 최대한 정확히 맞도록 한다(총원이 홀수일 때만 1명 차이).
    combined_front = [("A", m) for m in female_a] + [("S", m) for m in female_s]
    max_total_front = max_row * 2
    front_pool = combined_front[:max_total_front]
    front_overflow = combined_front[max_total_front:]  # 1,2열 물리적 정원을 넘는 극단적 초과분

    for i, (part, name) in enumerate(front_pool):
        rows[i % 2].append((part, name))

    # 1,2열에 물리적으로 다 못 들어간 인원: 알토는 성도석으로, 소프라노는 3열로(유동적 조정)
    leftover_s = [name for part, name in front_overflow if part == "S"]
    overflow_a = [name for part, name in front_overflow if part == "A"]
    if overflow_a:
        leftover["A"].extend(overflow_a)

    # 3열 배정: 소프라노(고정 우선순위 + 1,2열에 못 들어간 유동적 인원)는
    # 성도석으로 보내지 않고 3열에 전원 배치한다. 알토는 1,2열에만 배정하므로
    # 여기서는 더 이상 배치하지 않는다.
    for name in s_row3_fixed:
        rows[2].append(("S", name))
    for name in leftover_s:
        rows[2].append(("S", name))

    # 3열/4열 지정석(BASE_SEAT_ORDER)은 "누가 먼저 앉는지" 순서 우선순위로만 쓰고,
    # 실제로 3열에 앉을지 4열에 앉을지는 항상 그 순간 인원이 더 적은 열로 배정한다.
    # (실제 출석자 대부분이 지정석 명단에 있어서, 예전처럼 명단이 곧 열을 확정해버리면
    #  균형 로직이 사실상 작동하지 않고 지정석 순서 그대로(불균형하게) 배치되는 문제가 있었음)
    men_order = []
    seen_men = set()
    for name in BASE_SEAT_ORDER[3] + BASE_SEAT_ORDER[4]:
        if name in seen_men: continue
        for p in ("T", "B"):
            if name in people[p]:
                men_order.append((p, name))
                seen_men.add(name)
                break
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

    # 성도석(leftover)은 알토 파트만 배정한다(단, 위에서 이미 전원 배치했으므로
    # 실제로는 발생하지 않음 — 전체 정원 자체가 초과된 극단적 경우를 대비한 안전망).
    # 안전망: 그 외 사유로 자리를 못 찾은 인원이 있다면 성도석으로 보낸다.
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
