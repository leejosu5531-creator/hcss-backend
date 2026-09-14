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

# 인증서를 쓰지 않으므로 credentials=True + origins=* 조합(스펙 위반)을 피한다.
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

# 출석부가 엑셀(.xlsx) 파일에서 구글 스프레드시트로 바뀌면서, 드라이브에서 파일을
# 찾을 때(메타데이터/수정시간 확인용)는 drive.readonly, 실제 셀 값을 읽을 때는
# spreadsheets.readonly 스코프가 필요하다.
SCOPES = [
    "https://www.googleapis.com/auth/drive.readonly",
    "https://www.googleapis.com/auth/spreadsheets.readonly",
]
DRIVE_FILE_NAME = "26년 할렐루야 출석부"  # 구글 시트는 확장자가 없다
SHEET_NAMES = ["소프라노", "알토", "테너", "베이스"]

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
# 캐시 (Google Sheets API 불필요한 재호출을 줄이기 위함)
# ---------------------------------------------------------------------------
_cache_lock = threading.Lock()
_cache = {
    "drive_service": None,      # 구글 드라이브 서비스 객체 (재사용, 메타데이터/검색용)
    "sheets_service": None,     # 구글 시트 서비스 객체 (재사용, 셀 값 읽기용)
    "file_id": None,            # 검색된 스프레드시트 ID
    "file_id_cached_at": 0.0,
    "sheet_values": None,       # 마지막으로 Google Sheets API에서 읽은 원본 값
    "attendance_cache": {},     # (date_str,) -> {"data": {...}, "cached_at": ts}
}

FILE_ID_TTL = 3600          # 파일 검색(전체 드라이브 탐색) 결과 캐시 유지 시간(초)
ATTENDANCE_CACHE_TTL = 300  # 동일 날짜의 출석 파싱 결과 캐시 유지 시간(초)


def get_google_credentials():
    b64_str = os.environ.get("SERVICE_ACCOUNT_BASE64", "").strip()
    if not b64_str:
        raise ValueError("SERVICE_ACCOUNT_BASE64 환경변수가 설정되지 않았습니다.")
    decoded_json = base64.b64decode(b64_str).decode("utf-8")
    return Credentials.from_service_account_info(json.loads(decoded_json), scopes=SCOPES)


def get_drive_service():
    """구글 드라이브 서비스 객체를 프로세스 내에서 재사용한다(파일 검색/수정시간 확인용).

    build()에 cache_discovery=False를 주지 않으면 매 호출마다 디스커버리 문서를
    로컬 파일 캐시에서 읽고 쓰려고 시도하면서 지연이 발생할 수 있다
    (특히 Render처럼 파일시스템이 제한된 환경에서 체감이 크다).
    """
    with _cache_lock:
        if _cache["drive_service"] is not None:
            return _cache["drive_service"]
    creds = get_google_credentials()
    service = build("drive", "v3", credentials=creds, cache_discovery=False)
    with _cache_lock:
        _cache["drive_service"] = service
    return service


def get_sheets_service():
    """구글 시트 서비스 객체를 프로세스 내에서 재사용한다(셀 값 읽기용)."""
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
    # 동일한 이름의 파일이 여러 개 있을 수 있으므로(재업로드 등으로 예전 파일이 남아있는 경우),
    # 수정시간(modifiedTime) 내림차순으로 정렬해서 항상 가장 최근 파일을 선택한다.
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
        if len(fs_old) > 1:
            logger.warning(f"동일한 이름의 파일이 {len(fs_old)}개 발견되어 가장 최근 수정된 파일을 사용합니다: {fs_old[0].get('id')}")
        return fs_old[0]
    if len(fs) > 1:
        logger.warning(f"동일한 이름의 파일이 {len(fs)}개 발견되어 가장 최근 수정된 파일을 사용합니다: {fs[0].get('id')}")
    return fs[0]


def get_cached_file_id(drive_service):
    now = time.time()
    with _cache_lock:
        cached = _cache["file_id"]
        cached_at = _cache["file_id_cached_at"]
    if cached and (now - cached_at < FILE_ID_TTL):
        logger.info(f"[파일ID 캐시 사용] file_id={cached} (캐시된 지 {now - cached_at:.0f}초)")
        return cached
    file_info = find_drive_file(drive_service)
    with _cache_lock:
        _cache["file_id"] = file_info["id"]
        _cache["file_id_cached_at"] = now
    logger.info(f"[파일ID 재검색] file_id={file_info['id']}, 파일명={file_info.get('name')}")
    return file_info["id"]


def fetch_sheet_values(sheets_service, spreadsheet_id):
    """4개 시트(소프라노/알토/테너/베이스)의 값을 한 번의 API 호출(batchGet)로 가져온다."""
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
    """Google Sheets 전용 데이터 조회.

    force_refresh=True이면 캐시를 무시하고 Google Sheets API에서 즉시 다시 읽는다.
    Drive modifiedTime을 확인하거나 1~3분 기다리지 않는다.
    """
    if not force_refresh:
        with _cache_lock:
            if _cache["sheet_values"] is not None:
                logger.info("[캐시] Google Sheets 원본 데이터 재사용")
                return _cache["sheet_values"]

    logger.info("[시트 새로고침] Google Sheets API에서 최신 셀 값을 직접 조회")
    sheet_values = fetch_sheet_values(sheets_service, spreadsheet_id)

    with _cache_lock:
        _cache["sheet_values"] = sheet_values
        _cache["attendance_cache"].clear()

    total_rows = sum(len(v) for v in sheet_values.values())
    logger.info(f"[시트 조회 완료] 총 {total_rows}행, 캐시 갱신")
    return sheet_values


def get_cached_attendance(sheet_values, date_str, target_date):
    """현재 Google Sheets 스냅샷에서 날짜별 출석자를 파싱한다."""
    now = time.time()
    key = (date_str,)
    with _cache_lock:
        entry = _cache["attendance_cache"].get(key)
        if entry and (now - entry["cached_at"] < ATTENDANCE_CACHE_TTL):
            logger.info(f"[출석 캐시 사용] key={key}")
            return entry["data"]

    logger.info(f"[출석 새로 파싱] date={date_str}")
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
    # 구글 시트(특히 한국어 로케일)는 날짜를 "2026. 9. 6"처럼 마침표+공백으로
    # 표시하는 경우가 많다. 구분자 뒤에 붙는 공백을 미리 제거해서
    # "2026.9.6" 형태로 정규화한 뒤 기존 형식들을 그대로 적용한다.
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
    """구글 시트 API로 읽어온 원본 값(시트이름 -> 2차원 리스트)에서 출석자를 뽑아낸다.
    시트 API는 빈 셀을 건너뛰므로 각 행의 길이가 제각각일 수 있어(ragged), 매번
    길이를 확인하고 인덱싱한다."""
    sheet_map = {"소프라노": "S", "알토": "A", "테너": "T", "베이스": "B"}
    people = {p: [] for p in "SATB"}
    target_year = target_date.year
    found_date_col = False

    for sheet_name, part_code in sheet_map.items():
        rows = sheet_values.get(sheet_name)
        if not rows:
            logger.info(f"[진단] '{sheet_name}' 시트에 데이터가 없습니다(빈 시트이거나 이름 불일치).")
            continue

        # 1~2행(0-index 0,1)에서 날짜 열 찾기 (date_col은 1부터 시작하는 열 번호로 유지)
        date_col = None
        for header_row in rows[:2]:
            for col_idx, cell_val in enumerate(header_row, start=1):
                parsed_d = parse_cell_date(cell_val, target_year)
                if parsed_d and parsed_d == target_date:
                    date_col = col_idx
                    break
            if date_col: break
        if not date_col:
            logger.info(f"[진단] '{sheet_name}' 시트에서 {target_date} 날짜 열을 못 찾음. "
                        f"1행={rows[0] if len(rows) > 0 else None}, "
                        f"2행={rows[1] if len(rows) > 1 else None}")
            continue

        found_date_col = True

        for row_vals in rows[2:]:  # 3행부터(0-index 2~)
            name_val = row_vals[1] if len(row_vals) > 1 else None  # B열
            if not name_val: continue
            name_str = str(name_val).strip()
            if not name_str: continue
            if name_str in ("구분", "월통계", "총인원") or "통계" in name_str: break

            c_val = row_vals[date_col - 1] if len(row_vals) >= date_col else None
            is_present = False
            if isinstance(c_val, bool):
                is_present = c_val
            elif isinstance(c_val, str):
                if c_val.strip().upper() in ("TRUE", "O", "○", "ㅇ", "출석", "참석", "Y", "YES", "1", "●"):
                    is_present = True
            elif isinstance(c_val, (int, float)) and c_val == 1:
                is_present = True

            if is_present: people[part_code].append(name_str)

    if not found_date_col:
        # 네 개 시트(소프라노/알토/테너/베이스) 어디에서도 해당 날짜 열을 찾지 못한 경우
        # 빈 데이터로 조용히 진행하지 않고, 명확하게 알려준다.
        raise ValueError(f"{target_date.isoformat()} 날짜의 출석 데이터를 찾을 수 없습니다. "
                          f"출석부에 해당 날짜 열이 있는지 확인해 주세요.")

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


def plan_front_row_size(female_count, male_count, fixed_row3_count, cap_front):
    """1,2열 크기(r, r=1열=2열)를 탐색해 2,3,4열 편차가 최대한 1 이하가 되도록 정한다.

    1,2열과 3열은 여성(알토+소프라노)이 앉을 수 있고, 4열은 남성(테너+베이스)만
    앉을 수 있다는 물리적 제약을 반영한다. 단순히 총원을 4등분하면(성비를
    무시하면) 실제로는 달성 불가능한 목표가 나올 수 있어, 실제 여성/남성
    인원수를 넣고 후보 r을 모두 시도해 가장 균형 잡힌 조합을 고른다.
    (2열>=3열>=4열 순서를 우선 지키고, 그 안에서 편차 1을 넘는 만큼을 제곱으로 벌점을
    줘서 어느 한 곳에 편차가 몰리지 않도록 한다.)
    """
    best = None
    max_r = min(cap_front, female_count // 2)
    for r in range(0, max_r + 1):
        leftover_f = female_count - 2 * r
        total34 = leftover_f + male_count + fixed_row3_count
        row3 = (total34 + 1) // 2
        row4 = total34 // 2
        min_row3 = leftover_f + fixed_row3_count  # 이미 확정된 여성 인원보다 3열이 작을 수 없음
        if row3 < min_row3:
            row3 = min_row3
            row4 = total34 - row3

        # 오르간이 실제로 어느 열에 들어갈지까지 반영해서(최종 화면 표시 기준으로) 편차를 계산한다.
        # (사람 배치 이후 오르간 배치 규칙과 동일한 규칙: 4열이 3열보다 적을 때만 4열에,
        #  그 외에는 3열에 넣어 "4열>3열"이 되는 일을 절대 만들지 않는다)
        if row4 < row3:
            row3_disp, row4_disp = row3, row4 + 1
        else:
            row3_disp, row4_disp = row3 + 1, row4

        d23 = r - row3_disp     # 0 또는 1이 이상적 (2열 >= 3열, 오르간 포함 기준)
        d34 = row3_disp - row4_disp  # 0 또는 1이 이상적 (3열 >= 4열, 오르간 포함 기준)

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

    # 이제 알토도 3열 후보가 될 수 있다(음향적으로 3열에서 성부가 섞이는 것이 낫다는
    # 판단, 그리고 성도석 인접성 때문에 "정원 초과 시엔 알토 우선"이라는 규칙만 유지).
    # 알토와 소프라노를 하나의 여성 풀로 합쳐서 1,2,3열에 걸쳐 유동적으로 배치한다.
    name_to_part = {}
    for m in people["A"]: name_to_part[m] = "A"
    for m in people["S"]: name_to_part[m] = "S"

    # 이름 중복 방지(seen-set): 원본 좌석 순서표에 같은 이름이 두 번 들어있어도
    # 한 자리에만 배치되도록 안전장치를 둔다.
    _seen_f = set()
    female_pool = []  # (part, name) 튜플, BASE_SEAT_ORDER[1]+[2] 순서 우선
    for m in BASE_SEAT_ORDER[1] + BASE_SEAT_ORDER[2]:
        if m in name_to_part and m not in _seen_f:
            female_pool.append((name_to_part[m], m))
            _seen_f.add(m)
    for m in people["A"] + people["S"]:
        if m not in _seen_f:
            female_pool.append((name_to_part[m], m))
            _seen_f.add(m)

    # female_pool이 사본을 들고 있으므로 원본은 비워둔다.
    people["A"] = []
    people["S"] = []

    # 2단계: 실제 여성/남성 인원수를 반영해 1열=2열 크기(r)를 다시 산출한다.
    # (설정 정원만으로 총원을 4등분하면 성비에 따라 2,3열 편차가 커질 수 있음)
    female_count = len(female_pool)
    male_count_remaining = len(people["T"]) + len(people["B"])
    r = plan_front_row_size(female_count, male_count_remaining, 0, cap_front=18)

    max_row = r  # 1,2열 목표 크기(각 열)

    # 1,2열 배정: 먼저 원래 좌석 순서표대로 BASE_SEAT_ORDER[1]/[2] 우선순위로
    # "1,2열에 들어갈 전체 인원(앞줄 존)"을 정한 뒤, 그 안에서 알토/소프라노 각각을
    # 1열-2열에 최대한 균등하게 나눈다. 열의 총원뿐 아니라 파트별 인원까지 두 열이
    # 서로 맞아야, 같은 파트가 앞뒤로 겹쳐 서서 소리를 맞추기 쉽기 때문이다
    # (예: 1열 A8,S9 / 2열 A6,S11처럼 총원만 같고 파트 구성이 다르면 좋지 않음).
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
        return taken1, taken2, fill  # fill: 두 열을 다 채우고도 남는 인원(초과)

    names_in_pool = {n for _, n in female_pool}
    part_of = dict((n, p) for p, n in female_pool)

    row1_pool = [n for n in BASE_SEAT_ORDER[1] if n in names_in_pool]
    row2_pool = [n for n in BASE_SEAT_ORDER[2] if n in names_in_pool]
    listed = set(row1_pool) | set(row2_pool)
    unlisted_pool = [n for _, n in female_pool if n not in listed]

    front1, front2, female_overflow_names = _split_by_named_list(row1_pool, row2_pool, unlisted_pool, r, r)
    front_zone = [(part_of[n], n) for n in (front1 + front2)]  # 1,2열에 들어갈 전체 인원(순서는 아래에서 재정렬)
    female_overflow = [(part_of[n], n) for n in female_overflow_names]

    # 앞줄 존 안에서 알토/소프라노를 각각 절반씩 1열-2열에 나눈다(파트 구성까지 맞춤).
    # 총원이 2r로 정확히 맞아떨어지도록, 알토가 홀수로 남으면 그 잉여를 1열에,
    # 소프라노 잉여는 2열에 주는 식으로 서로 상쇄시켜 1열=2열 총원을 정확히 유지한다.
    front_alto = [n for p, n in front_zone if p == "A"]
    front_soprano = [n for p, n in front_zone if p == "S"]

    alto_row1, alto_row2 = front_alto[0::2], front_alto[1::2]      # 알토는 1열이 홀수 잉여를 가져감
    soprano_row2, soprano_row1 = front_soprano[0::2], front_soprano[1::2]  # 소프라노는 2열이 홀수 잉여를 가져감(상쇄)

    rows[0] = [("A", n) for n in alto_row1] + [("S", n) for n in soprano_row1]
    rows[1] = [("A", n) for n in alto_row2] + [("S", n) for n in soprano_row2]

    # 1,2열에 못 들어간 여성(알토+소프라노 섞여 있음)은 3열로 유동적으로 배치한다.
    # (물리적으로 정말 못 들어가는 경우는 거의 없다 — 3열이 사실상 여성을 계속
    #  받아줄 수 있기 때문. 전체 정원 자체가 초과된 경우는 이미 위에서 알토를
    #  성도석으로 먼저 뺐으므로 여기까지 오지 않는다)
    for part, name in female_overflow:
        rows[2].append((part, name))

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
    if people["S"]:
        # 소프라노는 원칙적으로 성도석에 가지 않지만, 혹시 모를 예외 상황에 대비한 안전망
        leftover["S"].extend(people["S"])
        people["S"] = []

    # 오르간 배치 열 결정: 사람 배치가 다 끝난 뒤 정한다(미리 정하면 화면 표시
    # 총원(인원+오르간)이 오히려 불균형해지는 문제가 있었음).
    # "4열이 3열보다 커지는 일은 절대 없어야 한다"가 최우선 조건이므로:
    #   - 실제 4열 인원이 이미 3열보다 적으면(그래서 +1 해도 3열을 못 넘으면) 4열에 배치
    #   - 그 외(3열=4열 동률 등)에는 반드시 3열에 배치해서 4열이 3열을 넘지 않게 한다
    real_row3, real_row4 = len(rows[2]), len(rows[3])
    if real_row4 < real_row3:
        organ_row = 3
    else:
        # 실제 4열이 3열보다 적지 않은 경우(동률 등): 4열에 넣으면 4열이 3열을 넘어설
        # 위험이 있으므로 무조건 3열에 배치한다. (2열=3열이 될 수는 있어도
        # 4열이 3열을 넘는 것보다는 낫다고 판단)
        organ_row = 2

    rt = [len(rows[0]), len(rows[1]), real_row3, real_row4]

    # 열별로 좌우 배치 순서를 다르게 둔다.
    #   3열: 베이스 - 알토 - 소프라노 - 테너 - 오르간(필요시)
    #   4열: 베이스 - 테너 - 오르간(필요시)
    # 3열과 4열 모두 베이스가 왼쪽, 테너가 그다음(오르간 제외하면 오른쪽 끝)에 오도록
    # 통일해서, 인원수가 매주 달라지더라도 두 열의 베이스/테너 무리가 같은 쪽에
    # 모이도록 한다.
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
            members = list(reversed(part_grouped[p]))
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
    """Render 무료 플랜은 일정 시간 요청이 없으면 서버가 잠든다(cold start).
    이 엔드포인트는 구글 드라이브 접근 없이 즉시 응답하므로, 외부 무료 핑 서비스
    (UptimeRobot, cron-job.org 등)로 5~10분 간격 주기 호출을 걸어두면 실제 좌석
    배치 요청이 콜드스타트로 30초~1분씩 걸리는 것을 크게 줄일 수 있다.

    또한 배포가 최신 코드로 반영됐는지 확인하는 용도로도 쓸 수 있다.
    브라우저에서 https://<서비스주소>/health 로 직접 접속해서
    code_version 값을 확인하면 된다."""
    return {"status": "ok", "code_version": "2026-09-13-google-sheets-direct-refresh-fix13"}


@app.post("/api/allocate")
def run_allocation(req: RequestModel):
    try:
        target = dt.date.fromisoformat(req.date_str)
        logger.info(
            f"[요청 시작] date={req.date_str}, total_seats={req.total_seats}, "
            f"force_refresh={req.force_refresh}"
        )

        sheets_service = get_sheets_service()

        # force_refresh=True = 사용자가 HCSS에서 '출석부 새로고침'을 누른 경우.
        # Google Sheets API에서 현재 값을 직접 다시 읽는다.
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
            "code_version": "2026-09-13-google-sheets-direct-refresh-fix13",
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
