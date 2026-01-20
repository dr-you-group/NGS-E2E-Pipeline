import os
import logging
from pathlib import Path

# 1. 기본 경로 설정 (pathlib 사용 권장)
# 현재 파일(config.py)의 부모 디렉토리를 프로젝트 루트로 설정
BASE_DIR = Path(__file__).resolve().parent

DB_PATH = BASE_DIR / "ngs_reports.db"
JSON_DIR = BASE_DIR / "json"
TMP_DIR = BASE_DIR / "tmp"
STATIC_DIR = BASE_DIR / "static"
TEMPLATE_DIR = BASE_DIR / "templates"

# 2. 로깅(Console) 설정
# print() 대신 사용할 로거 설정을 여기서 정의합니다.
def setup_logging():
    logging.basicConfig(
        level=logging.INFO,
        # 현재 print()와 최대한 비슷하게 보이도록 포맷을 단순화 (시간, 레벨 포함 가능)
        format='%(message)s',
        handlers=[
            logging.StreamHandler()  # 콘솔 출력
        ]
    )
    # 불필요한 라이브러리 로그는 숨김
    logging.getLogger("multipart").setLevel(logging.WARNING)
    logging.getLogger("uvicorn").setLevel(logging.WARNING)

    # 3. PPT 생성 설정 (pptx_generator.py에서 가져옴)
    PPT_CONFIG = {
        "FONT_NAME": "Arial",
        "MARGIN_LEFT_CM": 1.0,
        "DEFAULT_WIDTH_CM": 24.0,
    }