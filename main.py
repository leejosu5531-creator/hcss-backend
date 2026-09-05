import os
import datetime as dt
import logging
import traceback
from fastapi import FastAPI, HTTPException
from pydantic import BaseModel

# 1. 상세 디버깅을 위한 로깅 설정
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("allocation_debug")

app = FastAPI(title="화원교회 찬양대 좌석 배치 API")

# [기존 Pydantic 모델 및 설정 부분 유지]
# ...

@app.post("/api/allocate")
def run_allocation(req: RequestModel):
    logger.info("=== [1/6] 좌석 배치 요청 수신 ===")
    logger.info(f"전달받은 요청 데이터: date_str={req.date_str}, total_seats={req.total_seats}")

    try:
        # 1. 날짜 파싱 점검
        logger.info("=== [2/6] 날짜 데이터 변환 시도 ===")
        target = dt.date.fromisoformat(req.date_str)
        logger.info(f"변환된 target 날짜: {target}")

        # 2. 인증 파일 및 Google Drive API 연결 점검
        logger.info("=== [3/6] Google Drive 서비스 인증 시도 ===")
        logger.info(f"인증 파일 경로 확인: {SERVICE_ACCOUNT_FILE}")
        if not os.path.exists(SERVICE_ACCOUNT_FILE):
            logger.error(f"❌ 인증 파일이 존재하지 않습니다: {SERVICE_ACCOUNT_FILE}")
            raise FileNotFoundError(f"인증 파일 없음: {SERVICE_ACCOUNT_FILE}")
        
        service = google_drive_service()
        logger.info("Google Drive 서비스 생성 성공")

        # 3. 드라이브 파일 검색 점검
        logger.info(f"=== [4/6] 드라이브 파일 검색 시도 ('{DRIVE_FILE_NAME}') ===")
        file_info = find_drive_file(service)
        logger.info(f"찾은 파일 정보: ID={file_info.get('id')}, Name={file_info.get('name')}")

        # 4. 엑셀 파일 다운로드 점검
        logger.info("=== [5/6] 엑셀 파일 임시 다운로드 시도 ===")
        tmp_path = os.path.join(BASE_DIR, "_attendance_temp.xlsx")
        download_excel(service, file_info["id"], tmp_path)
        logger.info(f"임시 파일 저장 완료: {tmp_path}")

        # 5. 출석 데이터 읽기 및 배정 로직 점검
        logger.info("=== [6/6] 출석 데이터 분석 및 좌석 배치 연산 시도 ===")
        people = read_attendance(tmp_path, target)
        rows, leftover, rt, org = allocate(people, req.total_seats)

        # 임시 파일 삭제
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
            logger.info("임시 파일 삭제 완료")

        logger.info("=== ✅ 좌석 배치 성공적으로 완료됨 ===")
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
        # 모든 예외 발생 시 상세 Traceback을 Render Logs에 출력
        logger.error("❌ 처리 중 오류(Exception) 발생!")
        logger.error(f"오류 메시지: {str(e)}")
        logger.error("=== 상세 에러 스택 트레이스 (Traceback) ===")
        logger.error(traceback.format_exc())
        
        raise HTTPException(status_code=400, detail=str(e))
