# pptx_generator.py (새로 생성)

import os
import copy
from io import BytesIO
import pandas as pd
from pptx import Presentation
from pptx.util import Pt, Cm
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.xmlchemy import OxmlElement


class NGS_PPT_Generator:
    def __init__(self):
        # 템플릿 경로 설정 (app.py와 동일한 위치의 templates 폴더 가정)
        self.base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        self.template_dir = os.path.join(self.base_dir, "resources")

    def _set_cell_border(self, cell, border_color="000000", border_width='12700'):
        """셀 테두리 스타일 지정 헬퍼 함수"""
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        for lines in ['a:lnL', 'a:lnR', 'a:lnT', 'a:lnB']:
            ln = OxmlElement(lines)
            ln.set('w', border_width)
            ln.set('cap', 'flat')
            ln.set('cmpd', 'sng')
            ln.set('algn', 'ctr')

            solidFill = OxmlElement('a:solidFill')
            srgbClr = OxmlElement('a:srgbClr')
            srgbClr.set('val', border_color)
            solidFill.append(srgbClr)
            ln.append(solidFill)

            tcPr.append(ln)
        return cell

    def generate(self, report_data: dict) -> BytesIO:
        """
        JSON 데이터를 받아 PPT 바이너리 스트림(BytesIO)을 반환
        """
        # 1. 패널 타입에 따른 템플릿 로드
        panel_type = report_data.get('panel_type', 'GE')
        template_name = "blank_SA_report.pptx" if panel_type == 'SA' else "blank_GE_report.pptx"
        template_path = os.path.join(self.template_dir, template_name)

        if not os.path.exists(template_path):
            raise FileNotFoundError(f"템플릿 파일을 찾을 수 없습니다: {template_path}")

        prs = Presentation(template_path)

        # 2. 데이터 매핑 (기존 make_pptx_result.py 로직 이식)
        self._fill_clinical_info(prs, report_data['clinical_info'])
        self._fill_qc_info(prs, report_data.get('qc', {}))
        self._fill_biomarkers(prs, report_data.get('biomarkers', {}))

        # 3. [핵심] 변이 테이블 동적 생성 (기존에 없던 로직 추가)
        # 예시로 SNV Clinical 테이블만 구현 (다른 변이도 동일한 방식으로 추가 가능)
        if 'snv_clinical' in report_data and report_data['snv_clinical']['data']:
            self._create_variant_table(prs, report_data['snv_clinical'], "SNV (Clinical Significance)")

        # 4. 결과를 메모리에 저장
        output = BytesIO()
        prs.save(output)
        output.seek(0)
        return output

    def _fill_clinical_info(self, prs, info):
        """환자 기본 정보 채우기 (슬라이드 1페이지 가정)"""
        slide = prs.slides[0]
        # (좌표 등은 템플릿에 맞춰 조정 필요, 기존 코드 참조하여 구현)
        # 예시: 텍스트 박스나 표를 찾아 값 입력
        # 실제 구현 시에는 make_pptx_result.py의 report_clinical_info_table_1 함수 로직을 여기에 넣습니다.
        pass

    def _fill_qc_info(self, prs, qc_data):
        """QC 정보 채우기 (슬라이드 3페이지 가정)"""
        if len(prs.slides) > 3:
            slide = prs.slides[3]
            # QC 테이블 그리기 로직
        pass

    def _fill_biomarkers(self, prs, bio_data):
        """TMB, MSI 정보 채우기"""
        pass

    def _create_variant_table(self, prs, variant_data, title):
        """
        변이 데이터 테이블 생성 및 슬라이드 추가 로직
        데이터가 많으면 자동으로 새 슬라이드를 만듭니다.
        """
        rows_per_slide = 10  # 슬라이드당 최대 행 수
        headers = variant_data.get('headers', [])
        data = variant_data.get('data', [])

        # 데이터를 페이지 단위로 자르기
        chunks = [data[i:i + rows_per_slide] for i in range(0, len(data), rows_per_slide)]

        for i, chunk in enumerate(chunks):
            # 2번째 장부터는 새 슬라이드 추가 (빈 슬라이드 레이아웃 사용)
            if i > 0:
                slide_layout = prs.slide_layouts[6]  # 6: 빈 슬라이드
                slide = prs.slides.add_slide(slide_layout)
            else:
                slide = prs.slides[4]  # 첫 장은 템플릿의 5번째 슬라이드(인덱스 4) 사용 가정

            # 테이블 그리기
            rows = len(chunk) + 1  # 헤더 포함
            cols = len(headers)
            left = Cm(1.0)
            top = Cm(4.0)  # 제목 아래 위치
            width = Cm(24.0)
            height = Cm(0.8 * rows)

            table = slide.shapes.add_table(rows, cols, left, top, width, height).table

            # 헤더 입력
            for col_idx, header_text in enumerate(headers):
                cell = table.cell(0, col_idx)
                cell.text = str(header_text)
                self._set_cell_border(cell)
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(200, 200, 200)  # 회색 배경

            # 데이터 입력
            for row_idx, row_data in enumerate(chunk):
                for col_idx, cell_value in enumerate(row_data):
                    cell = table.cell(row_idx + 1, col_idx)
                    cell.text = str(cell_value)
                    self._set_cell_border(cell)