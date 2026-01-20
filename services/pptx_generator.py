import copy
import os
from io import BytesIO

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Pt, Cm


class PPTReportConfig:
    """보고서 생성에 필요한 상수, 스타일, 규칙 등을 관리하는 설정 클래스"""

    # 1. 레이아웃 & 폰트 설정
    MARGIN_LEFT = Cm(1.0)
    DEFAULT_WIDTH = Cm(24.0)
    BODY_TOP_START = Cm(4.5)
    BODY_BOTTOM_LIMIT = Cm(18.0)

    FONT_NAME = "Arial"
    FONT_SIZE_TITLE = Pt(14)
    FONT_SIZE_HEADER = Pt(12)
    FONT_SIZE_BODY = Pt(9)

    COLOR_BLACK = RGBColor(0, 0, 0)
    COLOR_RED = RGBColor(255, 0, 0)
    COLOR_NAVY = RGBColor(0, 51, 102)
    COLOR_GRAY_BG = RGBColor(230, 230, 230)

    # 간격 설정
    SPACE_SECTION = Cm(0.1)  # 섹션 간 간격
    SPACE_TABLE_BOTTOM = Cm(0.1)  # 테이블 하단 간격
    SPACE_TITLE_BOTTOM = Cm(0.05)  # 제목 하단 간격

    # 2. 데이터 매핑
    CLINICAL_INFO_MAPPING = {
        "검체 정보": "검체 정보", "성별": "성별", "나이": "나이",
        "Unit NO.": "Unit NO.", "환자명": "환자명", "채취 장기": "채취 장기",
        "원발 장기": "원발 장기", "진단": "진단", "의뢰의": "의뢰의",
        "의뢰의 소속": "의뢰의 소속", "검체 유형": "검체 유형",
        "검체의 적절성여부": "검체의 적절성여부", "검체접수일": "검체접수일",
        "결과보고일": "결과보고일",
        "Tumor Mutation Burden": "tmb",
        "Microsatellite Instability": "msi"
    }

    # 3. 섹션 순서 및 키
    VARIANT_SECTIONS = [
        {"key": "snv_clinical", "title": "SNVs & Indels"},
        {"key": "fusion_clinical", "title": "Fusion gene"},
        {"key": "cnv_clinical", "title": "Copy number variation"},
        {"key": "lr_brca_clinical", "title": "Large rearrangements in BRCA1/2"},
        {"key": "splice_clinical", "title": "Splice variant"}
    ]

    # 4. 테이블 식별 규칙 (확장성 핵심)
    # required: 반드시 포함되어야 할 헤더 키워드 (OR 조건은 튜플로 묶음)
    # excluded: 포함되면 안 되는 헤더 키워드
    TABLE_IDENTIFICATION_RULES = [
        {
            "key": "snv_clinical",
            "required": ["GENE", ("MUTATION", "AA CHANGE")],
            "excluded": ["BURDEN"]  # TMB 제외
        },
        {
            "key": "fusion_clinical",
            "required": ["FUSION", "BREAKPOINT"],
            "excluded": []
        },
        {
            "key": "cnv_clinical",
            "required": ["COPY", "FOLD"],
            "excluded": ["EXON"]  # Large Rearrangement와 구분
        },
        {
            "key": "lr_brca_clinical",
            "required": [("BRCA", "EXON"), "FOLD"],  # BRCA 혹은 (EXON and FOLD) - 로직에서 처리
            "excluded": []
        },
        {
            "key": "splice_clinical",
            "required": [("SPLICE", "EXON"), "BREAKPOINT"],
            "excluded": []
        }
    ]

    # 5. 삭제할 소제목 키워드
    REMOVE_KEYWORDS = ["SNVs", "Fusion", "Copy", "Splice", "Rearrangement"]


class LayoutAnalyzer:
    def __init__(self, prs, slide):
        self.prs = prs
        self.slide = slide
        self.page_width = prs.slide_width
        self.page_height = prs.slide_height
        self.config = PPTReportConfig()

        # 위치 초기화
        self.body_top = self.config.BODY_TOP_START
        self.body_bottom = self.config.BODY_BOTTOM_LIMIT

        self.existing_elements = {
            "main_title": None,
            "prototypes": {}
        }

        self._analyze_and_extract()

    def _analyze_and_extract(self):
        shapes_to_remove = []
        section_2_top = self.page_height

        # 1. 섹션 2 위치 파악
        for shape in self.slide.shapes:
            if shape.has_text_frame:
                text = shape.text_frame.text.strip()
                if "2. Variants of unknown significance" in text:
                    section_2_top = shape.top

        # 2. 요소 분석 및 추출
        for shape in self.slide.shapes:
            # 2-1. 텍스트 분석
            if shape.has_text_frame:
                text = shape.text_frame.text.strip()

                if "1. Variants of clinical significance" in text:
                    self.existing_elements["main_title"] = shape.top + shape.height + self.config.SPACE_TITLE_BOTTOM
                    self.body_top = self.existing_elements["main_title"]

                if "검사기관" in text or "세브란스병원" in text:
                    if shape.top < self.page_height:
                        self.body_bottom = shape.top - Cm(0.5)

                # 소제목 삭제 (섹션 2 위쪽만)
                if shape.top < section_2_top:
                    if text.startswith("- ") and any(k in text for k in self.config.REMOVE_KEYWORDS):
                        shapes_to_remove.append(shape)

            # 2-2. 테이블 분석 (규칙 기반 식별)
            if shape.has_table:
                if shape.top >= section_2_top:
                    continue

                tbl = shape.table
                try:
                    if len(tbl.rows) > 0:
                        headers = [cell.text_frame.text.strip().upper() for cell in tbl.rows[0].cells]
                        header_str = " ".join(headers)

                        target_key = self._identify_table_type(header_str)

                        if target_key:
                            self.existing_elements["prototypes"][target_key] = copy.deepcopy(shape.element)
                            shapes_to_remove.append(shape)
                except Exception as e:
                    print(f"Table analysis warning: {e}")

        # 추출된 요소 제거
        for shape in shapes_to_remove:
            try:
                sp = shape._element
                sp.getparent().remove(sp)
            except ValueError:
                pass

    def _identify_table_type(self, header_str: str) -> str:
        """Config 규칙에 따라 테이블 타입을 식별하는 헬퍼 메서드"""
        for rule in self.config.TABLE_IDENTIFICATION_RULES:
            # Required 조건 체크 (모든 조건 만족해야 함)
            is_match = True
            for req in rule["required"]:
                if isinstance(req, tuple):  # Tuple은 OR 조건
                    if not any(r in header_str for r in req):
                        is_match = False
                        break
                else:  # String은 AND 조건
                    if req not in header_str:
                        is_match = False
                        break

            # Excluded 조건 체크 (하나라도 있으면 탈락)
            if is_match:
                for exc in rule["excluded"]:
                    if exc in header_str:
                        is_match = False
                        break

            if is_match:
                return rule["key"]
        return None


class LayoutContext:
    def __init__(self, prs, start_slide, analyzer: LayoutAnalyzer):
        self.prs = prs
        self.current_slide = start_slide
        self.top = analyzer.body_top
        self.bottom_limit = analyzer.body_bottom
        self.margin = analyzer.config.MARGIN_LEFT
        self.width = analyzer.config.DEFAULT_WIDTH
        self.config = analyzer.config

    def check_space(self, height):
        if self.top + height > self.bottom_limit:
            self.add_new_slide()

    def add_new_slide(self):
        # 6번 레이아웃(Blank) 사용 시도
        layout_idx = 6 if len(self.prs.slide_layouts) > 6 else -1
        blank_layout = self.prs.slide_layouts[layout_idx]
        self.current_slide = self.prs.slides.add_slide(blank_layout)
        self.top = Cm(1.5)

    def add_space(self, height):
        self.top += height


class NGS_PPT_Generator:
    def __init__(self):
        self.base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        self.template_dir = os.path.join(self.base_dir, "resources")
        self.config = PPTReportConfig()  # 설정 인스턴스

    def _set_cell_border(self, cell, border_color="000000", border_width='12700'):
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
        panel_type = report_data.get('panel_type', 'GE')
        template_name = "blank_SA_report.pptx" if panel_type == 'SA' else "blank_GE_report.pptx"
        template_path = os.path.join(self.template_dir, template_name)

        if not os.path.exists(template_path):
            raise FileNotFoundError(f"템플릿 파일을 찾을 수 없습니다: {template_path}")

        prs = Presentation(template_path)

        if len(prs.slides) > 0:
            self._fill_clinical_info(prs.slides[0], report_data)

        analyzer = LayoutAnalyzer(prs, prs.slides[0])
        self._process_clinical_variants(prs, report_data, analyzer)

        output = BytesIO()
        prs.save(output)
        output.seek(0)
        return output

    def _fill_clinical_info(self, slide, report_data):
        # 1. clinical_info 딕셔너리 추출 (없으면 빈 딕셔너리)
        inner_info = report_data.get('clinical_info', {})
        tables = [shape.table for shape in slide.shapes if shape.has_table]

        for ppt_label, data_key in self.config.CLINICAL_INFO_MAPPING.items():
            # 1순위: clinical_info 내부 검색
            value = inner_info.get(data_key)

            # 2순위: 데이터 루트에서 검색 (대소문자 호환 추가)
            if value is None:
                # 'tmb'로 찾고 없으면 'TMB'로도 찾아봄
                value = report_data.get(data_key) or report_data.get(data_key.upper())

            # 3순위: biomarkers 키 내부 검색 (대소문자 호환 추가)
            if value is None and 'biomarkers' in report_data:
                bio_data = report_data['biomarkers']
                value = bio_data.get(data_key) or bio_data.get(data_key.upper())

            final_value = str(value) if value is not None else ""

            self._search_and_fill_cell_below(tables, ppt_label, final_value)

    def _process_clinical_variants(self, prs, report_data, analyzer):
        layout = LayoutContext(prs, prs.slides[0], analyzer)

        if not analyzer.existing_elements["main_title"]:
            self._draw_main_section_title(layout, "1. Variants of clinical significance")

        for config in self.config.VARIANT_SECTIONS:
            key = config['key']
            title = config['title']
            section_data = report_data.get(key, {})

            rows = section_data.get('data', [])
            headers = section_data.get('headers', [])
            highlight = section_data.get('highlight', '')

            prototype_xml = analyzer.existing_elements["prototypes"].get(key)

            if rows and len(rows) > 0:
                self._render_section_header(layout, title, highlight=highlight)
                if prototype_xml:
                    self._render_table_using_prototype(layout, prototype_xml, rows)
                else:
                    self._render_table_from_scratch(layout, headers, rows)
            else:
                self._render_section_header(layout, title, is_none=True)

            layout.add_space(self.config.SPACE_SECTION)

    def _draw_main_section_title(self, layout, text):
        height = Cm(1.0)
        layout.check_space(height)
        tb = layout.current_slide.shapes.add_textbox(layout.margin, layout.top, layout.width, height)
        p = tb.text_frame.paragraphs[0]

        self._set_run_style(p.add_run(), text, is_bold=True,
                            font_size=self.config.FONT_SIZE_TITLE,
                            color=self.config.COLOR_NAVY)

        layout.add_space(height + self.config.SPACE_TITLE_BOTTOM)

    def _render_section_header(self, layout, title, highlight=None, is_none=False):
        height = Cm(0.8)
        layout.check_space(height)
        tb = layout.current_slide.shapes.add_textbox(layout.margin, layout.top, layout.width, height)
        p = tb.text_frame.paragraphs[0]

        # 1. 제목
        self._set_run_style(p.add_run(), f"- {title}", is_bold=True,
                            font_size=self.config.FONT_SIZE_HEADER,
                            color=self.config.COLOR_BLACK)

        # 2. 상태값
        if is_none:
            self._set_run_style(p.add_run(), ": None",
                                font_size=self.config.FONT_SIZE_HEADER,
                                color=self.config.COLOR_BLACK)
        elif highlight:
            self._set_run_style(p.add_run(), ": ", is_bold=True,
                                font_size=self.config.FONT_SIZE_HEADER,
                                color=self.config.COLOR_RED)

            # 3. 하이라이트 파싱 (첫 단어 이탤릭)
            variants = highlight.split(', ')
            for idx, variant in enumerate(variants):
                parts = variant.split(' ', 1)

                # 유전자명 (이탤릭)
                self._set_run_style(p.add_run(), parts[0], is_bold=True,
                                    font_size=self.config.FONT_SIZE_HEADER,
                                    color=self.config.COLOR_RED, italic=True)

                # 나머지 내용 (정자체)
                if len(parts) > 1:
                    self._set_run_style(p.add_run(), " " + parts[1], is_bold=True,
                                        font_size=self.config.FONT_SIZE_HEADER,
                                        color=self.config.COLOR_RED)

                if idx < len(variants) - 1:
                    self._set_run_style(p.add_run(), ", ", is_bold=True,
                                        font_size=self.config.FONT_SIZE_HEADER,
                                        color=self.config.COLOR_RED)

        layout.add_space(height)

    def _set_run_style(self, run, text, is_bold=False, font_size=None, color=None, italic=False):
        """텍스트 스타일 적용 헬퍼 메서드"""
        run.text = text
        run.font.name = self.config.FONT_NAME
        if font_size: run.font.size = font_size
        if color: run.font.color.rgb = color
        run.font.bold = is_bold
        run.font.italic = italic

    def _render_table_using_prototype(self, layout, prototype_xml, rows):
        row_height = Cm(0.8)
        header_height = Cm(0.8)

        available_height = layout.bottom_limit - layout.top

        if available_height < (header_height + row_height):
            layout.add_new_slide()
            available_height = layout.bottom_limit - layout.top

        max_rows = int((available_height - header_height) / row_height)

        if max_rows >= len(rows):
            self._insert_cloned_table(layout, prototype_xml, rows)
        else:
            current_batch = rows[:max_rows]
            next_batch = rows[max_rows:]
            self._insert_cloned_table(layout, prototype_xml, current_batch)
            layout.add_new_slide()
            self._render_table_using_prototype(layout, prototype_xml, next_batch)

    def _insert_cloned_table(self, layout, prototype_xml, rows):
        new_tbl_element = copy.deepcopy(prototype_xml)
        layout.current_slide.shapes._spTree.insert_element_before(new_tbl_element, 'p:extLst')

        table_shape = layout.current_slide.shapes[-1]
        table_shape.top = int(layout.top)
        table_shape.left = int(layout.margin)
        table = table_shape.table

        current_rows = len(table.rows)
        needed_rows = len(rows) + 1

        if needed_rows > current_rows:
            for _ in range(needed_rows - current_rows):
                self._duplicate_last_row(table)

        for r_idx, row_data in enumerate(rows):
            target_row = table.rows[r_idx + 1]
            for c_idx, val in enumerate(row_data):
                if c_idx < len(target_row.cells):
                    self._set_cell_text_preserving_style(
                        target_row.cells[c_idx],
                        str(val),
                        is_bold=True,
                        font_color=self.config.COLOR_RED
                    )

        table_height = sum([row.height for row in table.rows])
        layout.add_space(table_height + self.config.SPACE_TABLE_BOTTOM)

    def _set_cell_text_preserving_style(self, cell, text, is_bold=False, font_color=None):
        """텍스트 입력 전 기존 내용을 초기화하여 중복/깨짐 방지"""
        if not cell.text_frame.paragraphs:
            p = cell.text_frame.add_paragraph()
        else:
            p = cell.text_frame.paragraphs[0]

        p.text = ""

        p.alignment = PP_ALIGN.CENTER

        run = p.add_run()
        run.font.name = self.config.FONT_NAME
        run.font.size = self.config.FONT_SIZE_BODY

        if is_bold: run.font.bold = True
        if font_color: run.font.color.rgb = font_color
        run.text = text

    def _duplicate_last_row(self, table):
        new_row = copy.deepcopy(table._tbl.tr_lst[-1])
        for tc in new_row.tc_lst:
            txBody = tc.find('{http://schemas.openxmlformats.org/drawingml/2006/main}txBody')
            if txBody is not None:
                p_list = txBody.findall('{http://schemas.openxmlformats.org/drawingml/2006/main}p')
                for p in p_list:
                    for child in list(p):
                        p.remove(child)
        table._tbl.append(new_row)

    def _render_table_from_scratch(self, layout, headers, rows):
        rows_count = len(rows) + 1
        cols_count = len(headers)
        table_height = Cm(0.8 * rows_count)

        shape = layout.current_slide.shapes.add_table(
            rows_count, cols_count, layout.margin, layout.top, self.config.DEFAULT_WIDTH, table_height
        )
        table = shape.table

        for idx, h in enumerate(headers):
            cell = table.cell(0, idx)
            cell.text = str(h)
            self._set_cell_border(cell)
            cell.fill.solid()
            cell.fill.fore_color.rgb = self.config.COLOR_GRAY_BG

        for r, row_data in enumerate(rows):
            for c, val in enumerate(row_data):
                cell = table.cell(r + 1, c)
                cell.text = str(val)
                self._set_cell_border(cell)

        layout.add_space(table_height + self.config.SPACE_TABLE_BOTTOM)

    def _search_and_fill_cell_below(self, tables, target_label, value):
        clean_target = target_label.replace(" ", "").lower()

        for table in tables:
            for r_idx, row in enumerate(table.rows):
                for c_idx, cell in enumerate(row.cells):
                    cell_text = cell.text_frame.text.replace(" ", "").lower()

                    if clean_target in cell_text:
                        try:
                            if r_idx + 1 < len(table.rows):
                                target_cell = table.cell(r_idx + 1, c_idx)

                                # 기존 텍스트(단위) 가져오기 (예: "/Megabase")
                                original_text = target_cell.text_frame.text.strip()

                                # 입력할 값 문자열 정리
                                val_str = str(value).strip() if value is not None else ""

                                # [핵심 로직 수정] 중복 단위 방지
                                if val_str:
                                    # 1. 데이터(val_str) 안에 이미 단위(original_text)가 들어있는 경우
                                    if original_text and original_text in val_str:
                                        final_text = val_str

                                    # 2. 데이터에 단위가 없는 경우 (숫자만 있는 경우)
                                    elif original_text:
                                        final_text = f"{val_str} {original_text}"

                                    # 3. 기존 단위가 아예 없는 셀인 경우
                                    else:
                                        final_text = val_str
                                else:
                                    # 값이 없으면 기존 단위 유지
                                    final_text = original_text

                                # 스타일 적용하여 입력
                                self._set_cell_text_preserving_style(
                                    target_cell,
                                    final_text,
                                    is_bold=True
                                )
                                return
                        except Exception as e:
                            print(f"Warning: Failed to fill {target_label} - {e}")