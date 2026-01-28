import copy
import os
from io import BytesIO

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Pt, Cm


class PPTReportConfig:
    """보고서 생성에 필요한 상수, 스타일, 규칙 등을 관리하는 설정 클래스"""

    # 레이아웃 & 폰트 설정
    MARGIN_LEFT = Cm(1.0)
    DEFAULT_WIDTH = Cm(19.05)
    BODY_TOP_START = Cm(4.5)
    BODY_BOTTOM_LIMIT = Cm(18.0)

    FONT_NAME = "Arial"
    FONT_SIZE_TITLE = Pt(12)  # 14 -> 12로 변경 (범용성 강화)
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

    # 스타일 프리셋 정의
    STYLES = {
        "clinical": {"color": COLOR_RED, "bold": True},
        "unknown": {"color": COLOR_BLACK, "bold": False}
    }

    MARKER_BIOMARKERS = "Other Biomarkers"

    # 섹션 시작점을 찾기 위한 텍스트 매핑
    SECTION_START_MARKERS = {
        "clinical": "1. Variants of clinical significance",
        "unknown": "2. Variants of unknown significance"
    }

    # 고지문 시작 마커
    MARKER_DISCLAIMER = "*본 기관의 유전자 정보 검색"

    # 데이터 매핑
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

    # 섹션 순서 및 키
    VARIANT_SECTIONS = [
        # 1. Clinical Significance
        {"key": "snv_clinical", "title": "SNVs & Indels", "type": "clinical", "prototype_key": "snv_clinical"},
        {"key": "fusion_clinical", "title": "Fusion gene", "type": "clinical", "prototype_key": "fusion_clinical"},
        {"key": "cnv_clinical", "title": "Copy number variation", "type": "clinical", "prototype_key": "cnv_clinical"},
        {"key": "lr_brca_clinical", "title": "Large rearrangements in BRCA1/2", "type": "clinical",
         "prototype_key": "lr_brca_clinical"},
        {"key": "splice_clinical", "title": "Splice variant", "type": "clinical", "prototype_key": "splice_clinical"},

        # 2. Unknown Significance
        {"key": "snv_unknown", "title": "SNVs & Indels", "type": "unknown", "prototype_key": "snv_clinical"},
        {"key": "fusion_unknown", "title": "Fusion gene", "type": "unknown", "prototype_key": "fusion_clinical"},
        {"key": "cnv_unknown", "title": "Copy number variation", "type": "unknown", "prototype_key": "cnv_clinical"},
        {"key": "lr_brca_unknown", "title": "Large rearrangements in BRCA1/2", "type": "unknown",
         "prototype_key": "lr_brca_clinical"},
        {"key": "splice_unknown", "title": "Splice variant", "type": "unknown", "prototype_key": "splice_clinical"}
    ]

    # 테이블 식별 규칙 (확장성 핵심)
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

    # 삭제할 소제목 키워드
    REMOVE_KEYWORDS = ["SNVs", "Fusion", "Copy", "Splice", "Rearrangement", "Failed gene"]


class LayoutAnalyzer:
    def __init__(self, prs):
        self.prs = prs
        self.page_width = prs.slide_width
        self.page_height = prs.slide_height
        self.config = PPTReportConfig()

        # 위치 초기화
        self.body_top = self.config.BODY_TOP_START
        self.body_bottom = self.config.BODY_BOTTOM_LIMIT

        # 슬라이드별 하단 제한선 저장소 {slide_index: bottom_limit}
        self.slide_bottom_limits = {}

        self.existing_elements = {
            "prototypes": {},
            "disclaimer": None  # 고지문 요소 저장용
        }

        # 섹션별 시작 위치 저장소
        self.section_locations = {}
        self.slide_bottom_limits = {}
        self._analyze_all_slides()

    def _analyze_all_slides(self):
        for slide_idx, slide in enumerate(self.prs.slides):
            self._analyze_single_slide(slide, slide_idx)

    def _analyze_single_slide(self, slide, slide_idx):
        shapes_to_remove = []
        found_section_on_this_slide = False

        # 이 슬라이드에서 사용할 하단 한계선 초기화 (기본값: 전역 설정 18cm)
        current_slide_limit = self.config.BODY_BOTTOM_LIMIT

        # 1. 텍스트 분석 (섹션 마커 및 Footer)
        for shape in slide.shapes:
            if shape.has_text_frame:
                text = shape.text_frame.text.strip()

                # "Other Biomarker" 텍스트 감지 -> 하단 한계선 위로 올림
                if "- Other Biomarker" in text:
                    # 해당 텍스트의 상단(Top)에서 0.5cm 위를 한계선으로 설정
                    limit = shape.top - Cm(0.5)
                    # 기존보다 더 위쪽(작은 값)이면 업데이트
                    if limit > current_slide_limit:
                        current_slide_limit = limit

                # 섹션 시작 마커 찾기
                for section_type, marker in self.config.SECTION_START_MARKERS.items():
                    if marker in text:
                        self.section_locations[section_type] = {
                            "slide_index": slide_idx,
                            "slide_id": slide.slide_id,
                            "top": shape.top + shape.height + self.config.SPACE_TITLE_BOTTOM,
                            "has_title": True,
                            "title_shape": shape  # 제거를 위해 shape 참조 저장
                        }
                        found_section_on_this_slide = True

                        # Clinical 섹션인 경우 기본 시작점 업데이트
                        if section_type == "clinical":
                            self.body_top = self.section_locations[section_type]["top"]

                # Footer 위치 파악 (전역 설정도 업데이트)
                if "검사기관" in text or "세브란스병원" in text:
                    if shape.top < self.page_height:
                        limit = shape.top - Cm(0.5)
                        if limit < current_slide_limit:
                            current_slide_limit = limit
                        # 하위 호환성을 위해 전역 변수도 업데이트
                        self.body_bottom = current_slide_limit

                # 소제목 삭제 대상 수집
                if found_section_on_this_slide:
                    text_lower = text.lower()
                    is_remove_target = any(k.lower() in text_lower for k in self.config.REMOVE_KEYWORDS)
                    if is_remove_target:
                        shapes_to_remove.append(shape)
                
                # 고지문(Disclaimer) 감지 -> 저장 후 제거 (나중에 동적으로 다시 그림)
                if text.startswith(self.config.MARKER_DISCLAIMER):
                    self.existing_elements['disclaimer'] = copy.deepcopy(shape.element)
                    shapes_to_remove.append(shape)

        # 분석된 슬라이드의 한계선을 ID 기반으로 저장
        self.slide_bottom_limits[slide.slide_id] = current_slide_limit

        # 2. 테이블 요소 분석 및 프로토타입 추출
        for shape in slide.shapes:
            if shape.has_table:
                tbl = shape.table
                try:
                    if len(tbl.rows) > 0:
                        headers = [cell.text_frame.text.strip().upper() for cell in tbl.rows[0].cells]
                        header_str = " ".join(headers)

                        target_key = self._identify_table_type(header_str)

                        if target_key:
                            if target_key not in self.existing_elements['prototypes']:
                                self.existing_elements['prototypes'][target_key] = copy.deepcopy(shape.element)

                            if found_section_on_this_slide:
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
    def __init__(self, prs, analyzer: LayoutAnalyzer):
        self.prs = prs
        self.analyzer = analyzer
        self.current_slide = prs.slides[0]

        start_loc = analyzer.section_locations.get("clinical")
        if start_loc:
            self.current_slide_index = start_loc["slide_index"]
            self.top = start_loc["top"]
        else:
            self.current_slide_index = 0
            self.top = analyzer.body_top

        self.current_slide = prs.slides[self.current_slide_index]
        self.margin = analyzer.config.MARGIN_LEFT
        self.width = analyzer.config.DEFAULT_WIDTH
        self.config = analyzer.config

    @property
    def bottom_limit(self):
        """현재 보고 있는 슬라이드의 하단 한계선을 동적으로 가져옴"""
        # 슬라이드 ID로 조회
        return self.analyzer.slide_bottom_limits.get(
            self.current_slide.slide_id,
            self.analyzer.config.BODY_BOTTOM_LIMIT
        )

    def move_to_section(self, section_type):
        """특정 섹션의 지정된 위치(템플릿 페이지)로 이동"""
        loc = self.analyzer.section_locations.get(section_type)
        if loc:
            target_slide = self._find_slide_by_id(loc['slide_id'])
            if target_slide:
                self.current_slide = target_slide
                self.top = loc["top"]
                self.current_slide_index = loc["slide_index"]  # 인덱스 업데이트
                return True
        return False

    def check_space(self, height):
        if self.top + height > self.bottom_limit:
            self.add_new_slide()

    def add_new_slide(self):
        # 1. 새 슬라이드 생성 (기본적으로 맨 뒤에 추가됨)
        layout_idx = 6 if len(self.prs.slide_layouts) > 6 else -1
        blank_layout = self.prs.slide_layouts[layout_idx]
        new_slide = self.prs.slides.add_slide(blank_layout)

        # [테두리 복사] 템플릿의 2번째 슬라이드(인덱스 1)에서 배경/테두리 스타일 복사
        if len(self.prs.slides) > 1:
            reference_slide = self.prs.slides[1]
            self._copy_template_style(new_slide, reference_slide)
        
        # [공간 확장] 새로 생성된 슬라이드는 테두리 끝까지 쓸 수 있도록 하단 한계선을 24.5cm로 확장
        self.analyzer.slide_bottom_limits[new_slide.slide_id] = Cm(24.5)

        # 2. 목표 인덱스 계산 (현재 + 1)
        target_index = self.current_slide_index + 1

        # 3. 목표가 맨 뒤가 아니라면 슬라이드 이동
        if target_index < len(self.prs.slides) - 1:
            xml_slides = self.prs.slides._sldIdLst
            slides = list(xml_slides)
            xml_slides.remove(slides[-1])  # 마지막(방금 추가한 것) 제거
            xml_slides.insert(target_index, slides[-1])  # 목표 위치에 삽입
            
            # 새 위치에서 슬라이드 객체 다시 가져오기
            self.current_slide = self.prs.slides[target_index]
        else:
            self.current_slide = new_slide

        # 4. 트래킹 정보 업데이트
        self.current_slide_index = target_index
        self.top = Cm(0.5)

    def _copy_template_style(self, target_slide, source_slide):
        """소스 슬라이드에서 배경이나 테두리 같은 정적 도형을 복사"""
        for shape in source_slide.shapes:
            # 텍스트나 테이블이 아닌, 오토쉐이프(도형)나 그림만 복사 대상
            # (테두리는 보통 직사각형 도형으로 구현됨)
            if shape.shape_type in [1, 6]: # 1: AUTO_SHAPE, 6: GROUP
                # 테이블이나 텍스트 프레임이 있는 경우 제외 (단, 단순 도형도 텍스트 프레임 가질 수 있음)
                # 여기서는 '테두리'로 추정되는 크기나 특성을 필터링하면 좋겠지만, 
                # 우선 텍스트가 없는 도형을 복사하는 것으로 단순화
                if not shape.has_table:
                     if not shape.has_text_frame or not shape.text_frame.text.strip():
                         new_el = copy.deepcopy(shape.element)
                         target_slide.shapes._spTree.append(new_el)

    def add_space(self, height):
        self.top += height

    def _find_slide_by_id(self, slide_id):
        for slide in self.prs.slides:
            if slide.slide_id == slide_id:
                return slide
        return None


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

        analyzer = LayoutAnalyzer(prs)
        self._process_all_variants(prs, report_data, analyzer)

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

    def _extract_highlight_keywords(self, report_data):
        """Report Data에서 코멘트 강조를 위한 키워드(유전자명 등)를 추출합니다."""
        keywords = set()
        target_headers = ['GENE', 'MUTATION', 'AA CHANGE', 'FUSION', 'SPLICE', 'COPY', 'BREAKPOINT']
        
        # 제외할 일반적인 단어 리스트 (단독으로 Bold 처리되지 않도록)
        BLACKLIST = {"mutation", "fusion", "amplification", "deletion", "duplication", 
                     "insertion", "splice", "variant", "copy", "loss", "gain", 
                     "rearrangement", "skipping", "indel", "snv", "cnv", "none"}
        
        # 1. Variant Sections 순회
        for section_config in self.config.VARIANT_SECTIONS:
            key = section_config['key']
            section_data = report_data.get(key, {})
            rows = section_data.get('data', [])
            headers = section_data.get('headers', [])
            
            if not rows or not headers:
                continue
            
            # 헤더 인덱스 매핑 (대소문자 무관하게 처리)
            indices = {}
            for h_idx, h_val in enumerate(headers):
                h_upper = str(h_val).upper().strip()
                if h_upper in target_headers:
                    indices[h_upper] = h_idx
            
            for row in rows:
                gene = None
                
                # GENE
                if 'GENE' in indices and len(row) > indices['GENE']:
                    val = str(row[indices['GENE']]).strip()
                    if val and val.lower() != 'none':
                        gene = val
                        keywords.add(gene)
                        # 편의상 예상되는 조합형 키워드 추가
                        keywords.add(f"{gene} mutation") 
                        keywords.add(f"{gene} amplification")
                        keywords.add(f"{gene} fusion")
                
                # 그 외 컬럼들 (MUTATION, AA CHANGE, FUSION 등)
                for header_key in ['MUTATION', 'AA CHANGE', 'FUSION', 'SPLICE']:
                    if header_key in indices and len(row) > indices[header_key]:
                        val = str(row[indices[header_key]]).strip()
                        if val and val.lower() not in BLACKLIST:
                            # 1) 단독 키워드 추가 (블랙리스트에 없으면)
                            keywords.add(val)
                        
                        # 2) Gene 결합형 키워드 추가 (항상 추가)
                        if gene and val and val.lower() != 'none':
                            # 중복 공백 제거
                            composite = f"{gene} {val}"
                            keywords.add(composite)

        # 2. Failed Gene
        failed_gene = report_data.get('failed_gene')
        if failed_gene and str(failed_gene).lower() != 'none':
             keywords.add(str(failed_gene))
        
        # 3. 추가 기본 키워드 (사용자 요청 사례 등)
        keywords.add("MET exon14 skipping")
        
        # 최종 필터링: 너무 짧은 단어(2글자 이하) 등 제거
        final_keywords = [k for k in keywords if len(k) > 2]
        return list(set(final_keywords))

    def _process_all_variants(self, prs, report_data, analyzer):
        layout = LayoutContext(prs, analyzer)
        current_section_group = None

        for section_config in self.config.VARIANT_SECTIONS:
            key = section_config['key']
            title = section_config['title']
            style_type = section_config['type']
            prototype_key = section_config['prototype_key']

            # 스타일 설정 가져오기 (clinical: Red/Bold, unknown: Black/Normal)
            style_props = self.config.STYLES.get(style_type, self.config.STYLES["unknown"])

            if current_section_group != style_type:
                current_section_group = style_type
                
                # [리팩토링된 흐름 로직]
                # 템플릿의 앵커로 이동할지, 아니면 현재 흐름을 유지할지 결정
                target_loc = analyzer.section_locations.get(style_type)
                should_move = False
                
                if target_loc:
                    target_slide_idx = target_loc["slide_index"]
                    target_slide_id = target_loc["slide_id"]

                    # 1. 이미 현재 슬라이드가 목표 슬라이드인 경우 (예: Clinical 섹션 시작)
                    # 별도 이동 없이 앵커를 사용하도록 허용
                    if layout.current_slide.slide_id == target_slide_id:
                        should_move = True
                    # 2. 목표 슬라이드를 아직 지나지 않았을 때만 이동
                    # 엄격한 부등호(<) 사용: 오버플로우로 인해 현재 인덱스가 목표 인덱스와 같아진 경우 점프 방지
                    elif layout.current_slide_index < target_slide_idx:
                        should_move = True
                
                found_anchor = False
                if should_move:
                    found_anchor = layout.move_to_section(style_type)
                else:
                    # [정리 로직]
                    # 이동하지 않기로 결정했다면(오버플로우 등), 
                    # 중복 방지를 위해 원래 템플릿 슬라이드의 제목을 제거해야 함
                    if target_loc and "title_shape" in target_loc:
                        try:
                            sp = target_loc["title_shape"]
                            if sp.element.getparent() is not None:
                                sp.element.getparent().remove(sp.element)
                        except Exception as e:
                            print(f"Warning: Failed to remove unused title shape: {e}")
                
                # 이동하지 않았거나(오버플로우), 앵커를 못 찾았다면 직접 타이틀 그리기
                if not found_anchor:
                    main_title = self.config.SECTION_START_MARKERS.get(style_type)
                    if main_title:
                        # 기본값이 12pt Black으로 변경되었으므로 별도 파라미터 불필요
                        self._draw_main_section_title(layout, main_title)

            section_data = report_data.get(key, {})
            rows = section_data.get('data', [])
            headers = section_data.get('headers', [])
            highlight_val = section_data.get('highlight', [])

            prototype_xml = analyzer.existing_elements["prototypes"].get(prototype_key)

            if rows and len(rows) > 0:
                header_height = Cm(0.8)
                min_table_height = Cm(0.8)
                required_height = header_height + min_table_height

                if layout.top + required_height > layout.bottom_limit:
                    layout.add_new_slide()

                self._render_section_header(layout, title, highlight_data=highlight_val)
                if prototype_xml:
                    self._render_table_using_prototype(layout, prototype_xml, rows, style_props)
                else:
                    self._render_table_from_scratch(layout, headers, rows)
            else:
                self._render_section_header(layout, title, is_none=True)

            layout.add_space(self.config.SPACE_SECTION)

        # [복구] 3. Failed gene 섹션 (Variants 처리 루프 종료 후)
        failed_gene = report_data.get('failed_gene')
        self._draw_failed_gene(layout, failed_gene)

        # [복구] Comments 섹션
        comments = report_data.get('comments', [])
        highlight_keywords = self._extract_highlight_keywords(report_data)
        self._draw_comments(layout, comments, highlight_keywords)

    def _draw_main_section_title(self, layout, text, font_size=None, color=None):
        height = Cm(1.0)
        layout.check_space(height)
        tb = layout.current_slide.shapes.add_textbox(layout.margin, layout.top, layout.width, height)
        p = tb.text_frame.paragraphs[0]
        
        # 기본값 설정 (None일 경우 Config 기본값 사용)
        final_size = font_size if font_size else self.config.FONT_SIZE_TITLE
        final_color = color if color else self.config.COLOR_BLACK  # Navy -> Black 변경

        self._set_run_style(p.add_run(), text, is_bold=True,
                            font_size=final_size,
                            color=final_color)

        layout.add_space(height + self.config.SPACE_TITLE_BOTTOM)

    def _render_section_header(self, layout, title, highlight_data=None, is_none=False):
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
        elif highlight_data:
            # ": " 찍기
            self._set_run_style(p.add_run(), ": ", is_bold=True,
                                font_size=self.config.FONT_SIZE_HEADER,
                                color=self.config.COLOR_RED)

            # 1. 단순히 리스트를 하나씩 꺼내서 그림 (로직 없음)
            for segment in highlight_data:
                # 데이터에서 텍스트와 스타일을 꺼냄
                text = segment['text']
                style = segment.get('style', 'normal')

                # 스타일이 'italic'이면 True, 아니면 False
                is_italic = (style == 'italic')

                # 2. 그대로 출력 (조건문 분기 불필요)
                self._set_run_style(p.add_run(), text, is_bold=True,
                                    font_size=self.config.FONT_SIZE_HEADER,
                                    color=self.config.COLOR_RED,
                                    italic=is_italic)  # 스타일 적용

        layout.add_space(height)

    def _set_run_style(self, run, text, is_bold=False, font_size=None, color=None, italic=False):
        """텍스트 스타일 적용 헬퍼 메서드"""
        run.text = text
        run.font.name = self.config.FONT_NAME
        if font_size: run.font.size = font_size
        if color: run.font.color.rgb = color
        run.font.bold = is_bold
        run.font.italic = italic

    def _render_table_using_prototype(self, layout, prototype_xml, rows, style_props):
        row_height = Cm(0.8)
        header_height = Cm(0.8)

        available_height = layout.bottom_limit - layout.top

        if available_height < (header_height + row_height):
            layout.add_new_slide()
            available_height = layout.bottom_limit - layout.top

        max_rows = int((available_height - header_height) / row_height)

        if max_rows >= len(rows):
            self._insert_cloned_table(layout, prototype_xml, rows, style_props)
        else:
            current_batch = rows[:max_rows]
            next_batch = rows[max_rows:]
            self._insert_cloned_table(layout, prototype_xml, current_batch, style_props)
            layout.add_new_slide()
            self._render_table_using_prototype(layout, prototype_xml, next_batch, style_props)

    def _insert_cloned_table(self, layout, prototype_xml, rows, style_props):
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
                        is_bold=style_props['bold'],
                        font_color=style_props['color']
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
            rows_count, cols_count, layout.margin, layout.top, layout.width, table_height
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

    def _draw_failed_gene(self, layout, failed_gene_text):
        """3. Failed gene 섹션을 그립니다."""
        # 제목이자 내용이 한 줄에 들어가는 형태: "3. Failed gene: {내용}"
        height = Cm(1.0)
        layout.check_space(height)
        
        tb = layout.current_slide.shapes.add_textbox(layout.margin, layout.top, layout.width, height)
        p = tb.text_frame.paragraphs[0]
        
        # "3. Failed gene" 부분 (Unknown과 동일하게 12pt Black 적용)ㄴ
        self._set_run_style(p.add_run(), "3. Failed gene", is_bold=True,
                            font_size=self.config.FONT_SIZE_TITLE, color=self.config.COLOR_BLACK)
        
        # ": " 및 내용 부분
        content = f": {failed_gene_text}" if failed_gene_text else ": None"
        self._set_run_style(p.add_run(), content, is_bold=False,
                            font_size=self.config.FONT_SIZE_HEADER, color=self.config.COLOR_BLACK)
        
        layout.add_space(height + self.config.SPACE_SECTION)


    def _draw_comments(self, layout, comments, highlight_keywords=None):
        """Comments 섹션과 하단 고지문을 그립니다."""
        # 섹션 헤더: "▣ Comment"
        header_height = Cm(1.0)
        layout.check_space(header_height)
        
        tb_header = layout.current_slide.shapes.add_textbox(layout.margin, layout.top, layout.width, header_height)
        p_header = tb_header.text_frame.paragraphs[0]
        # 기본값이 12pt Black으로 변경되었으므로 Config 참조
        self._set_run_style(p_header.add_run(), "▣ Comment", is_bold=True, 
                            font_size=self.config.FONT_SIZE_TITLE, color=self.config.COLOR_BLACK)
        layout.add_space(header_height)

        # 주의사항/고지문 (Disclaimer) 템플릿 요소 가져오기
        disclaimer_xml = layout.analyzer.existing_elements.get('disclaimer')
        
        if disclaimer_xml:
            self._insert_disclaimer_element(layout, disclaimer_xml, comments, highlight_keywords)
        else:
            self._draw_default_disclaimer(layout, comments, highlight_keywords)

    def _insert_disclaimer_element(self, layout, element_xml, comments=None, highlight_keywords=None):
        """템플릿에서 추출한 XML 요소를 사용하여 고지문을 그립니다. (코멘트 포함)"""
        import re
        
        # 1. 요소 복제 및 추가
        new_sp = copy.deepcopy(element_xml)
        layout.current_slide.shapes._spTree.append(new_sp)
        
        # 2. Shape 래핑
        disclaimer_shape = layout.current_slide.shapes[-1]
        
        # 3. 코멘트 내용 주입 (TextFrame의 맨 앞에 추가)
        if comments:
            if isinstance(comments, str):
                comments = [comments]
            
            tf = disclaimer_shape.text_frame
            tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT  # 텍스트 양에 맞춰 도형 크기 자동 조절
            tf.word_wrap = True

            # 기존 문단(고지문) 앞에 빈 줄 추가 (구분용)
            # tf.paragraphs[0].insert_paragraph_before() -> 지원하지 않음. OxmlElement 사용
            p_xml = OxmlElement('a:p')
            tf.paragraphs[0]._p.addprevious(p_xml)
            p_empty = tf.paragraphs[0] # 새로 추가된 빈 문단
            
            # 코멘트 삽입 (역순으로 맨 앞에 계속 추가하면 정순서가 됨)
            for comment in reversed(comments):
                # 1. Spacer (빈 줄) 추가 (각 대상 사이 간격)
                spacer_xml = OxmlElement('a:p')
                # spacer_xml.set('h', '100') # 높이 조절 가능하나 기본값 사용
                tf.paragraphs[0]._p.addprevious(spacer_xml)

                # 2. Comment 문단 추가
                p_xml = OxmlElement('a:p')
                tf.paragraphs[0]._p.addprevious(p_xml)
                p = tf.paragraphs[0]
                
                p.alignment = PP_ALIGN.LEFT # 정렬 보정
                
                # 범용 볼드 처리 로직 (Extractor 기반)
                # 추출된 highlight_keywords 중 코멘트에 포함된 가장 긴 키워드를 찾아 볼드 처리
                # (중복 매칭 방지를 위해 하나만 찾거나, re.split으로 모두 처리)
                
                matched = False
                if highlight_keywords:
                    # 길이 역순 정렬 (긴 키워드 우선)
                    keywords_sorted = sorted(highlight_keywords, key=len, reverse=True)
                    
                    # 가장 먼저 매칭되는 키워드 하나를 찾아서 처리 (복잡한 중첩 처리 방지)
                    # "MET exon14 skipping"이 "MET"보다 먼저 매칭되도록 정렬됨
                    for kw in keywords_sorted:
                        if kw in comment:
                             parts = re.split(f"({re.escape(kw)})", comment)
                             for part in parts:
                                 is_bold_part = (part == kw)
                                 self._set_run_style(p.add_run(), part, is_bold=is_bold_part, 
                                                     font_size=Pt(8), color=self.config.COLOR_BLACK)
                             matched = True
                             break
                
                if not matched:
                    # 매칭되는 키워드가 없으면 2순위 Regex (Colon) 시도
                    match_col = re.match(r"^([^:]+)(:)(.*)$", comment, re.DOTALL)
                    if match_col:
                        subject = match_col.group(1)
                        colon = match_col.group(2)
                        rest = match_col.group(3)
                        self._set_run_style(p.add_run(), subject, is_bold=True, 
                                            font_size=Pt(8), color=self.config.COLOR_BLACK)
                        self._set_run_style(p.add_run(), colon + rest, is_bold=False, 
                                            font_size=Pt(8), color=self.config.COLOR_BLACK)
                    else:
                        # 그냥 출력
                        self._set_run_style(p.add_run(), comment, is_bold=False, 
                                            font_size=Pt(8), color=self.config.COLOR_BLACK)
            
            # [수정] 텍스트 박스 높이 수동 계산 및 보정
            # SHAPE_TO_FIT_TEXT가 파일 오픈 시점에 바로 반영되지 않는 문제 해결책
            # 대략적인 높이를 계산하여 명시적으로 top을 조정하거나 height를 늘려줌
            # Disclaimer 텍스트 길이 + 코멘트 길이 + 줄바꿈
            total_text_len = sum([len(c) for c in comments]) + 500 # Disclaimer approx 400-500 chars
            est_lines = (total_text_len / 70) + len(comments) * 2 + 5 # 70 chars/line, spacers, paragraphs
            est_height = Cm(est_lines * 0.5) # 0.5cm line height (approx for 8pt)
            
            # 기존 height보다 크면 반영 (단, SHAPE_TO_FIT_TEXT가 켜져있으면 무시될 수 있으나 초기 렌더링 힌트로 작용)
            if est_height > disclaimer_shape.height:
                 disclaimer_shape.height = int(est_height)

        # 4. 높이 확인 및 페이지 넘김
        item_height = disclaimer_shape.height
        
        if layout.top + item_height > layout.bottom_limit:
            sp = disclaimer_shape._element
            sp.getparent().remove(sp)
            layout.add_new_slide()
            
            # 재귀 호출로 다시 그림 (로직 재사용)
            # 단, 이미 코멘트를 주입한 XML 상태가 아니므로(copy본에 주입함), 
            # 원본 XML과 comments를 다시 넘겨야 함. 
            # 하지만 여기서 'new_sp'는 이미 수정된 상태임. 
            # 수정된 'new_sp'를 다시 쓰는 게 좋음.
            
            # 수정된 요소를 복제해서 새 슬라이드에 추가
            new_sp_2 = copy.deepcopy(new_sp) 
            layout.current_slide.shapes._spTree.append(new_sp_2)
            disclaimer_shape = layout.current_slide.shapes[-1]
            
        # 5. 위치 설정
        disclaimer_shape.top = int(layout.top)
        layout.add_space(item_height + Cm(0.2))

    def _draw_default_disclaimer(self, layout, comments=None, highlight_keywords=None):
        """템플릿 요소가 없을 경우 기본 텍스트로 그립니다."""
        # 코멘트 처리
        if comments:
            if isinstance(comments, str):
                comments = [comments]
            for comment in comments:
                # is_bold_keyword 제거 -> _draw_long_text 내부 범용 로직 사용
                self._draw_long_text(layout, comment, font_size=Pt(8), highlight_keywords=highlight_keywords)
            layout.add_space(Cm(0.2))
            
        disclaimer_text = (
            "*본 기관의 유전자 정보 검색 및 해석은 dbSNP, COSMIC, ClinVar, c-bioportal 등의 유전자 정보검색 사이트를 참고로 하고 있습니다. "
            "또한, 발견된 유전자 변이의 임상적 의미에 대하여, 아래 참고 문헌에 기반한 4단계 시스템 (Tier I: Strong clinical significance, "
            "Tier II: Potential Clinical significance, Tier III: Unknown Clinical significance, Tier IV: Benign or Likely Benign) 을 "
            "적용하여 보고하고 있습니다. 또한 normal cell sequencing이 이루어지지 않아 일부 변이의 경우 germline polymorphism의 가능성을 배제할 수 없습니다.\n"
            "Standards and Guidelines for the Interpretation and Reporting of Sequence Variants in Cancer. A Joint Consensus Recommendation "
            "of the Association for Molecular Pathology, American Society of Clinical Oncology and College of American Pathologists. "
            "J Mol Diagn. 2017; 19(1):4-23."
        )
        additional_info = (
            "• 본 검사의 raw data (BAM, FASTQ, VCF) 파일은 분자 병리 검사실 내 병리과 서버 컴퓨터에서 보관, 관리되고 있습니다.\n"
            "• 본 검사의 결과는 검체에 포함된 정상세포와 암세포의 비율에 따라 위음성의 결과를 배제 할 수 없습니다."
        )
        
        self._draw_long_text(layout, disclaimer_text, font_size=Pt(8))
        layout.add_space(Cm(0.2))
        self._draw_long_text(layout, additional_info, font_size=Pt(8))

    def _draw_long_text(self, layout, text, font_size=Pt(8), is_bold_keyword=None, highlight_keywords=None):
        height = Cm(0.8)
        layout.check_space(height)
        tb = layout.current_slide.shapes.add_textbox(layout.margin, layout.top, layout.width, height)
        p = tb.text_frame.paragraphs[0]
        
        matched = False
        if highlight_keywords:
            import re
            keywords_sorted = sorted(highlight_keywords, key=len, reverse=True)
            for kw in keywords_sorted:
                if kw in text:
                     parts = re.split(f"({re.escape(kw)})", text)
                     for part in parts:
                         is_bold = (part == kw)
                         # 폰트 사이즈는 파라미터를 따라감
                         self._set_run_style(p.add_run(), part, font_size=font_size, color=self.config.COLOR_BLACK, is_bold=is_bold)
                     matched = True
                     break
        
        if not matched:
            if is_bold_keyword and is_bold_keyword in text:
                 # (호환성을 위해 남겨둠)
                 import re
                 pattern = f"({re.escape(is_bold_keyword)})"
                 parts = re.split(pattern, text)
                 for part in parts:
                     if not part: continue
                     is_bold = (part == is_bold_keyword)
                     self._set_run_style(p.add_run(), part, font_size=font_size, color=self.config.COLOR_BLACK, is_bold=is_bold)
            else:
                 # 2순위: Colon 패턴
                 regex_col = r"^([^:]+)(:)(.*)$"
                 match_col = re.match(regex_col, text, re.DOTALL)
                 
                 if match_col:
                    subject = match_col.group(1)
                    colon = match_col.group(2)
                    rest = match_col.group(3)
                    self._set_run_style(p.add_run(), subject, is_bold=True, font_size=font_size, color=self.config.COLOR_BLACK)
                    self._set_run_style(p.add_run(), colon + rest, is_bold=False, font_size=font_size, color=self.config.COLOR_BLACK)
                 else:
                    self._set_run_style(p.add_run(), text, font_size=font_size, color=self.config.COLOR_BLACK)

        layout.add_space(height)
