from __future__ import annotations

from pathlib import Path

from PIL import Image as PILImage
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    Image,
    ListFlowable,
    ListItem,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)


def register_font() -> str:
    for candidate in (
        Path("C:/Windows/Fonts/malgun.ttf"),
        Path("C:/Windows/Fonts/NanumGothic.ttf"),
    ):
        if candidate.exists():
            font_name = "ManualFont"
            if font_name not in pdfmetrics.getRegisteredFontNames():
                pdfmetrics.registerFont(TTFont(font_name, str(candidate)))
            return font_name
    return "Helvetica"


def build_styles(font_name: str):
    styles = getSampleStyleSheet()
    styles.add(
        ParagraphStyle(
            name="ManualTitle",
            parent=styles["Title"],
            fontName=font_name,
            fontSize=20,
            leading=26,
            alignment=TA_CENTER,
            textColor=colors.HexColor("#17324d"),
            spaceAfter=12,
        )
    )
    styles.add(
        ParagraphStyle(
            name="ManualSubTitle",
            parent=styles["Heading2"],
            fontName=font_name,
            fontSize=13,
            leading=18,
            textColor=colors.HexColor("#17324d"),
            spaceBefore=12,
            spaceAfter=8,
        )
    )
    styles.add(
        ParagraphStyle(
            name="ManualBody",
            parent=styles["BodyText"],
            fontName=font_name,
            fontSize=10,
            leading=16,
            alignment=TA_LEFT,
            wordWrap="CJK",
            spaceAfter=4,
        )
    )
    styles.add(
        ParagraphStyle(
            name="ManualNote",
            parent=styles["BodyText"],
            fontName=font_name,
            fontSize=9,
            leading=14,
            textColor=colors.HexColor("#6b7280"),
            wordWrap="CJK",
        )
    )
    return styles


def p(text: str, style) -> Paragraph:
    return Paragraph(text.replace("\n", "<br/>"), style)


def bullet_list(items: list[str], style):
    return ListFlowable(
        [ListItem(p(item, style), leftIndent=8) for item in items],
        bulletType="bullet",
        start="circle",
        leftIndent=14,
    )


def make_info_table(rows: list[list[str]], col_widths):
    table = Table(rows, colWidths=col_widths, repeatRows=1)
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#d9e5f2")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#17324d")),
                ("FONTNAME", (0, 0), (-1, -1), "ManualFont"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("LEADING", (0, 0), (-1, -1), 13),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#9aa7b6")),
                ("BACKGROUND", (0, 1), (-1, -1), colors.white),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )
    return table


def scaled_image(path: Path, max_width: float, max_height: float):
    with PILImage.open(path) as image:
        width, height = image.size
    ratio = min(max_width / width, max_height / height)
    return Image(str(path), width=width * ratio, height=height * ratio)


def page_number(canvas, doc):
    canvas.saveState()
    canvas.setFont("ManualFont", 8)
    canvas.setFillColor(colors.HexColor("#6b7280"))
    canvas.drawRightString(doc.pagesize[0] - 18 * mm, 10 * mm, f"{doc.page}")
    canvas.restoreState()


def build_story(styles):
    project_root = Path(__file__).resolve().parents[1]
    sample_dir = project_root / "SAMPLE"
    output_dir = project_root / "output" / "doc"
    erp_image = sample_dir / "ERP 스크린샷.png"
    work_image = output_dir / "manual_work_tab.png"
    settings_image = output_dir / "manual_settings_tab.png"

    story = []
    story.append(Spacer(1, 8 * mm))
    story.append(p("유상사급 타처보관 확인서 생성 프로그램<br/>사용자 매뉴얼", styles["ManualTitle"]))
    story.append(p("기준 화면: 현재 프로그램 기준 업무 탭, 설정 탭, ERP 조회 화면", styles["ManualNote"]))
    story.append(Spacer(1, 6 * mm))

    story.append(p("1. 문서 개요", styles["ManualSubTitle"]))
    story.append(
        p(
            "이 문서는 유상사급 타처보관 확인서 생성 프로그램의 실제 화면을 기준으로 작성한 사용자 매뉴얼입니다. "
            "사용자가 사업장별 자재유형별수불집계 파일을 불러오고, 거래처를 선택한 뒤, 재고자산확인서 Excel/PDF를 생성하고 메일까지 발송하는 전체 흐름을 설명합니다.",
            styles["ManualBody"],
        )
    )
    story.append(
        bullet_list(
            [
                "입력 파일: 사업장별 자재유형별수불집계 Excel 파일",
                "출력 파일: 거래처별 재고자산확인서 Excel, PDF",
                "메일 기능: Outlook 즉시 발송, Gmail SMTP 발송 지원",
            ],
            styles["ManualBody"],
        )
    )

    story.append(p("2. 입력 소스 다운로드 경로", styles["ManualSubTitle"]))
    story.append(p("입력 소스 파일은 ERP에서 직접 다운로드합니다. 실제 업무 경로는 다음과 같습니다.", styles["ManualBody"]))
    story.append(
        make_info_table(
            [
                ["구분", "내용"],
                ["ERP 메뉴 경로", "ERP > 자재관리 > 자재실적관리 > 자재유형별수불집계"],
                ["조회 조건", "사업장, 해당년월, 자재대분류/중분류/소분류 등을 필요에 따라 설정"],
                ["수불유형", "자재매출 대상이 포함되도록 조회"],
                ["다운로드 대상", "거래처명, 자재코드, 색상, 자재명이 포함된 엑셀 파일"],
            ],
            [34 * mm, 136 * mm],
        )
    )
    if erp_image.exists():
        story.append(Spacer(1, 4 * mm))
        story.append(p("ERP 소스 다운로드 화면", styles["ManualSubTitle"]))
        story.append(scaled_image(erp_image, 170 * mm, 95 * mm))
        story.append(Spacer(1, 2 * mm))
        story.append(p("좌측 메뉴의 자재실적관리 아래 자재유형별수불집계를 선택한 뒤 조회 후 엑셀 파일로 저장합니다.", styles["ManualNote"]))

    story.append(p("3. 업무 탭 화면 구성", styles["ManualSubTitle"]))
    story.append(
        make_info_table(
            [
                ["영역", "설명"],
                ["상단 요약", "등록된 소스 파일 수, 불러온 거래처 수, 현재 선택한 업체 수를 표시합니다."],
                ["소스 파일", "사업장별 자재유형별수불집계 파일을 추가하고 사업장명을 확인 또는 수정합니다."],
                ["거래처 선택", "거래처명을 검색하고 체크박스로 여러 업체를 동시에 선택합니다."],
                ["출력 옵션", "기준일과 저장 폴더를 지정합니다. 기준일 기본값은 전월 말일입니다."],
                ["선택 업체 미리보기", "선택 업체별 탭으로 자재 목록을 검토합니다."],
                ["하단 실행 버튼", "선택 업체 반영, 전체 미리보기, 산출 실행, 메일 발송을 수행합니다."],
            ],
            [34 * mm, 136 * mm],
        )
    )
    if work_image.exists():
        story.append(Spacer(1, 4 * mm))
        story.append(p("업무 탭 스냅샷", styles["ManualSubTitle"]))
        story.append(scaled_image(work_image, 170 * mm, 102 * mm))
        story.append(Spacer(1, 2 * mm))
        story.append(p("업무 탭에서 소스 파일 불러오기, 거래처 선택, 미리보기, 산출 실행, 메일 발송 흐름을 처리합니다.", styles["ManualNote"]))

    story.append(PageBreak())

    story.append(p("4. 업무 처리 절차", styles["ManualSubTitle"]))
    story.append(
        bullet_list(
            [
                "1단계. [파일 추가]를 눌러 사업장별 자재유형별수불집계 파일을 등록합니다.",
                "2단계. 자동 추출된 사업장명이 맞는지 확인하고 필요 시 수정합니다.",
                "3단계. [거래처 불러오기]를 눌러 거래처 목록을 생성합니다.",
                "4단계. 거래처명 검색창으로 원하는 업체를 찾아 체크박스로 선택합니다.",
                "5단계. 기준일과 저장 폴더를 확인합니다.",
                "6단계. [선택 업체 반영]으로 하단 미리보기 탭을 갱신합니다.",
                "7단계. [산출 실행]으로 거래처별 Excel/PDF를 생성합니다.",
                "8단계. 필요하면 [메일 발송]으로 이어서 메일을 보냅니다.",
            ],
            styles["ManualBody"],
        )
    )

    story.append(p("5. 설정 탭 구성", styles["ManualSubTitle"]))
    story.append(
        make_info_table(
            [
                ["영역", "설명"],
                ["사용자 이름", "메일 본문에 들어갈 발신자 이름을 저장합니다."],
                ["기본 발송 방식", "Outlook 또는 Gmail(SMTP) 기본값을 저장합니다."],
                ["SMTP 설정", "Gmail SMTP 호스트, 포트, 계정, 비밀번호, TLS 사용 여부를 입력합니다."],
                ["메일 템플릿", "제목 템플릿과 본문 템플릿 기본값을 관리합니다."],
                ["업체 이메일 관리", "거래처별 이메일 주소를 저장하고 재사용합니다."],
            ],
            [34 * mm, 136 * mm],
        )
    )
    if settings_image.exists():
        story.append(Spacer(1, 4 * mm))
        story.append(p("설정 탭 스냅샷", styles["ManualSubTitle"]))
        story.append(scaled_image(settings_image, 170 * mm, 102 * mm))
        story.append(Spacer(1, 2 * mm))
        story.append(p("설정 탭에서 사용자 이름, SMTP 설정, 제목/본문 템플릿, 업체 이메일을 관리합니다.", styles["ManualNote"]))

    story.append(p("6. 메일 발송 기능", styles["ManualSubTitle"]))
    story.append(
        make_info_table(
            [
                ["항목", "동작 규칙"],
                ["발송 방식", "Outlook 즉시 발송 또는 Gmail SMTP 발송"],
                ["기본 방식", "사용자별 기본값 저장 후 필요 시 변경"],
                ["첨부파일", "거래처별 Excel + PDF 2개를 항상 첨부"],
                ["본문", "기본 시드를 제공하고 발송 전 사용자가 수정 가능"],
                ["회신 요청일", "메일 발송 창에서 직접 지정"],
                ["이메일 누락", "해당 업체를 발송 대상에서 제외하고 경고 팝업 표시"],
                ["발송 결과", "이력 저장 없이 결과 팝업으로만 표시"],
            ],
            [34 * mm, 136 * mm],
        )
    )
    story.append(
        bullet_list(
            [
                "Outlook은 최종 확인 후 즉시 발송합니다.",
                "Gmail은 SMTP 계정과 앱 비밀번호가 필요합니다.",
                "업체 이메일은 설정 탭의 이메일 관리 표에서 최초 입력 후 재사용합니다.",
            ],
            styles["ManualBody"],
        )
    )

    story.append(PageBreak())

    story.append(p("7. 출력 결과 확인", styles["ManualSubTitle"]))
    story.append(
        make_info_table(
            [
                ["출력물", "설명"],
                ["Excel", "헤더는 상단 1회만 출력되고 본문은 아래로 계속 이어집니다."],
                ["PDF", "페이지가 넘어가면 헤더가 반복되고 마지막 페이지에만 합계/서명 영역이 출력됩니다."],
                ["폴더 구조", "선택 거래처별 하위 폴더가 생성되고 각 폴더에 Excel/PDF가 저장됩니다."],
            ],
            [34 * mm, 136 * mm],
        )
    )

    story.append(p("8. 자주 발생하는 상황", styles["ManualSubTitle"]))
    story.append(
        make_info_table(
            [
                ["상황", "조치 방법"],
                ["거래처가 보이지 않음", "먼저 [거래처 불러오기]를 실행했는지 확인합니다."],
                ["파일이 읽히지 않음", "ERP에서 저장한 원본 파일인지, Sheet1이 존재하는지 확인합니다."],
                ["사업장명이 잘못 잡힘", "소스 파일 표의 사업장 칼럼에서 직접 수정합니다."],
                ["기존 파일이 있음", "산출 실행 시 덮어쓰기 확인 창이 표시됩니다."],
                ["이메일 발송 대상이 줄어듦", "설정 탭에서 해당 업체 이메일이 등록되어 있는지 확인합니다."],
                ["Gmail 발송 실패", "SMTP 계정, 앱 비밀번호, TLS 설정을 다시 확인합니다."],
            ],
            [34 * mm, 136 * mm],
        )
    )

    story.append(p("9. 운영 체크 포인트", styles["ManualSubTitle"]))
    story.append(
        bullet_list(
            [
                "사업장별 소스 파일은 동일 분기 기준 파일만 함께 사용합니다.",
                "ERP에서 내려받은 파일의 컬럼명은 임의 변경하지 않습니다.",
                "최종 산출 전에는 반드시 탭 미리보기 또는 전체 미리보기로 업체별 목록을 확인합니다.",
                "메일 발송 전에는 설정 탭의 사용자 이름, 발송 방식, 업체 이메일이 최신 상태인지 확인합니다.",
            ],
            styles["ManualBody"],
        )
    )

    story.append(Spacer(1, 6 * mm))
    story.append(
        p(
            "작성 비고: 본 매뉴얼은 현재 프로그램 화면과 ERP 경로 화면을 기준으로 작성되었습니다. "
            "추후 화면 구성이나 메일 기능 정책이 변경되면 매뉴얼도 함께 갱신해야 합니다.",
            styles["ManualNote"],
        )
    )
    return story


def main() -> None:
    project_root = Path(__file__).resolve().parents[1]
    output_dir = project_root / "output" / "doc"
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / "유상사급타처보관_사용자매뉴얼.pdf"

    font_name = register_font()
    styles = build_styles(font_name)
    doc = SimpleDocTemplate(
        str(output_path),
        pagesize=A4,
        leftMargin=18 * mm,
        rightMargin=18 * mm,
        topMargin=16 * mm,
        bottomMargin=16 * mm,
        title="유상사급 타처보관 사용자 매뉴얼",
    )
    doc.build(build_story(styles), onFirstPage=page_number, onLaterPages=page_number)
    print(output_path)


if __name__ == "__main__":
    main()
