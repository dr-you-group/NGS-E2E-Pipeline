import os
import json
import time
import logging
import config

logger = logging.getLogger("app")

def safe_remove_file(file_path: str, max_retries: int = 5, delay: float = 0.5) -> bool:
    """
    파일을 안전하게 삭제하는 함수 (Windows 파일 잠금 문제 해결)
    
    Args:
        file_path: 삭제할 파일 경로
        max_retries: 최대 재시도 횟수
        delay: 재시도 간격 (초)
    
    Returns:
        bool: 삭제 성공 여부
    """
    if not os.path.exists(file_path):
        return True

    for attempt in range(max_retries):
        try:
            os.remove(file_path)
            logger.info(f"임시 파일 삭제 완료: {file_path}")
            return True
        except PermissionError as e:
            if attempt < max_retries - 1:
                logger.warning(f"파일 삭제 재시도 {attempt + 1}/{max_retries}: {file_path}")
                time.sleep(delay)
            else:
                logger.error(f"파일 삭제 실패 (최대 재시도 도달): {file_path} - {e}")
                return False
        except Exception as e:
            logger.error(f"파일 삭제 중 예상치 못한 오류: {file_path} - {e}")
            return False
    return False


def save_json_file(specimen_id, report_data):
    """JSON 파일로 저장하는 함수"""
    try:
        json_file_path = config.JSON_DIR / f"{specimen_id}.json"
        with open(json_file_path, "w", encoding="utf-8") as json_file:
            json.dump(report_data, json_file, indent=4, ensure_ascii=False)
        logger.info(f"JSON 파일 저장 완료: {json_file_path}")
        return True
    except Exception as e:
        logger.error(f"JSON 파일 저장 실패: {str(e)}")
        return False
