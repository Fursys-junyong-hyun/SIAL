"""유상사급 타처보관 확인서 — 사용자 매뉴얼 PDF 생성기.

이 스크립트는 docs/사용매뉴얼.md 와 같은 구조의 친근한 톤으로 PDF 매뉴얼을 만든다.
실행: python scripts/generate_user_manual_pdf.py
산출물: output/doc/유상사급타처보관_사용자매뉴얼.pdf
"""
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
    KeepTogether,
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
            fontSize=22,
            leading=28,
            alignment=TA_CENTER,
            textColor=colors.HexColor("#17324d"),
            spaceAfter=14,
        )
    )
    styles.add(
        ParagraphStyle(
            name="ManualPartHeading",
            parent=styles["Heading1"],
            fontName=font_name,
            fontSize=16,
            leading=22,
            textColor=colors.HexColor("#0f4c81"),
            spaceBefore=18,
            spaceAfter=10,
        )
    )
    styles.add(
        ParagraphStyle(
            name="ManualStepHeading",
            parent=styles["Heading2"],
            fontName=font_name,
            fontSize=13,
            leading=18,
            textColor=colors.HexColor("#17324d"),
            spaceBefore=12,
            spaceAfter=6,
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
            name="ManualBodyBold",
            parent=styles["BodyText"],
            fontName=font_name,
            fontSize=10,
            leading=16,
            alignment=TA_LEFT,
            wordWrap="CJK",
            spaceAfter=4,
            textColor=colors.HexColor("#0f4c81"),
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
    styles.add(
        ParagraphStyle(
            name="ManualWarn",
            parent=styles["BodyText"],
            fontName=font_name,
            fontSize=10,
            leading=15,
            textColor=colors.HexColor("#9c3a3a"),
            wordWrap="CJK",
            backColor=colors.HexColor("#fbeaea"),
            borderColor=colors.HexColor("#e7c0c0"),
            borderWidth=0.6,
            borderPadding=6,
            spaceBefore=4,
            spaceAfter=6,
        )
    )
    return styles


def p(text: str, style) -> Paragraph:
    return Paragraph(text.replace("\n", "<br/>"), style)


def numbered_list(items: list[str], style):
    return ListFlowable(
        [ListItem(p(item, style), leftIndent=8) for item in items],
        bulletType="1",
        leftIndent=18,
    )


def bullet_list(items: list[str], style):
    return ListFlowable(
        [ListItem(p(item, style), leftIndent=8) for item in items],
        bulletType="bullet",
        start="circle",
        leftIndent=14,
    )


def info_table(rows: list[list[str]], col_widths, font_name: str = "ManualFont"):
    table = Table(rows, colWidths=col_widths, repeatRows=1)
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#d9e5f2")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#17324d")),
                ("FONTNAME", (0, 0), (-1, -1), font_name),
                ("FONTSIZE", (0, 0), (-1, -1), 9.5),
                ("LEADING", (0, 0), (-1, -1), 14),
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


def first_existing(*candidates: Path) -> Path | None:
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return None


def page_number(canvas, doc):
    canvas.saveState()
    canvas.setFont("ManualFont", 8)
    canvas.setFillColor(colors.HexColor("#6b7280"))
    canvas.drawRightString(doc.pagesize[0] - 18 * mm, 10 * mm, f"- {doc.page} -")
    canvas.restoreState()


# ---------------- Story builders ----------------

def build_title(styles, body_style, note_style):
    story = []
    story.append(Spacer(1, 10 * mm))
    story.append(p("유상사급 타처보관 확인서<br/>사용 매뉴얼", styles["ManualTitle"]))
    story.append(
        p(
            "이 프로그램은 거래처별 재고자산확인서를 자동으로 만들어서 메일로 보내는 도구입니다.<br/>"
            "처음에 한 번 설정해 두면, 다음부터는 버튼을 5번 누르면 끝납니다.",
            body_style,
        )
    )
    story.append(Spacer(1, 4 * mm))
    story.append(
        p(
            "이 매뉴얼은 비IT 담당자가 보는 것을 가정하고 작성되었습니다. "
            "차례 → 0부(ERP 다운로드) → 1부(처음 한 번만 설정) → 2부(매번 하는 작업) → 3부(문제 해결).",
            note_style,
        )
    )
    return story


def build_requirements(styles, body_style):
    story = []
    story.append(p("준비물", styles["ManualStepHeading"]))
    story.append(
        p(
            "작업을 시작하기 전에 두 가지가 필요합니다.",
            body_style,
        )
    )
    story.append(
        numbered_list(
            [
                "사업장별 <b>자재유형별수불집계</b> 엑셀 파일 (1개 이상) — 다음 0부에서 받는 법 설명",
                "본인 회사 <b>Gmail 주소</b> (예: name@fursys.com)",
            ],
            body_style,
        )
    )
    return story


def build_part0_erp(styles, body_style, note_style, warn_style, erp_image: Path | None):
    story = []
    story.append(p("0부 · 소스 파일 다운로드 (ERP 에서 받기)", styles["ManualPartHeading"]))
    story.append(
        p(
            "확인서를 만들려면 먼저 ERP 에서 사업장별 자재유형별수불집계 파일을 받아야 합니다. "
            "사업장 1곳당 파일 1개씩 받아두면 됩니다.",
            body_style,
        )
    )

    story.append(p("0-1. ERP 메뉴 위치", styles["ManualStepHeading"]))
    story.append(p("ERP 시스템에 로그인한 뒤 왼쪽 메뉴에서 다음 순서로 들어갑니다.", body_style))
    story.append(
        p(
            "<b>자재관리 &gt; 자재실적관리 &gt; 자재유형별수불집계</b>",
            styles["ManualBodyBold"],
        )
    )
    story.append(
        p(
            "화면 상단 탭 이름이 <b>자재유형별수불집계(MRS0020_M01)</b> 인지 확인합니다.",
            body_style,
        )
    )

    story.append(p("0-2. 조회 조건 설정", styles["ManualStepHeading"]))
    story.append(p("화면 위쪽 조회 조건 영역에서 다음을 지정합니다.", body_style))
    story.append(
        info_table(
            [
                ["항목", "값"],
                ["사업장", "받으려는 공장 선택 (예: 퍼시스충주1공장)"],
                ["해당년월", "조회할 분기 시작일 ~ 종료일 (예: 2026-04-01 ~ 2026-06-30 은 2분기)"],
                ["수불유형", "화면 좌측 목록에서 자재매출 한 줄만 체크"],
                ["카테고리대/중/소", "보통 그대로 두기 (전체)"],
            ],
            [40 * mm, 130 * mm],
        )
    )
    story.append(Spacer(1, 3 * mm))
    story.append(
        p(
            "<b>분기별 기준일 예시</b> · 1분기: 01-01~03-31 · 2분기: 04-01~06-30 · 3분기: 07-01~09-30 · 4분기: 10-01~12-31",
            note_style,
        )
    )

    if erp_image and erp_image.exists():
        story.append(Spacer(1, 4 * mm))
        story.append(p("[ ERP 화면 예시 ]", styles["ManualStepHeading"]))
        story.append(scaled_image(erp_image, 170 * mm, 110 * mm))
        story.append(Spacer(1, 1 * mm))
        story.append(
            p(
                "좌측 메뉴 '자재실적관리 → 자재유형별수불집계' 진입 후, 조회 조건에서 사업장/기간/수불유형(자재매출)을 설정합니다.",
                note_style,
            )
        )

    story.append(p("0-3. 조회 후 엑셀로 저장", styles["ManualStepHeading"]))
    story.append(p("오른쪽 위 <b>조회</b> 버튼(빨간색)을 누릅니다.", body_style))
    story.append(
        p(
            "데이터가 표시되면 ERP 화면의 엑셀 내보내기 기능으로 파일을 저장합니다. "
            "파일명에 사업장 이름이 들어가도록 합니다.",
            body_style,
        )
    )
    story.append(p("예: <b>자재유형별수불집계_퍼시스충주1공장.xls</b>", body_style))

    story.append(p("0-4. 여러 사업장이면 반복", styles["ManualStepHeading"]))
    story.append(
        p(
            "다른 사업장 파일도 필요한 만큼 같은 방법으로 받습니다. 파일은 모두 같은 폴더에 모아두면 편합니다.",
            body_style,
        )
    )

    return story


def build_part1_setup(styles, body_style, note_style, warn_style):
    story = [PageBreak()]
    story.append(p("1부 · 처음 한 번만 하는 설정 (메일 보내기 준비)", styles["ManualPartHeading"]))
    story.append(p("이 1부는 처음 사용할 때 한 번만 하면 됩니다. 다음부터는 안 해도 됩니다.", body_style))

    story.append(p("1-1. 프로그램을 켭니다", styles["ManualStepHeading"]))
    story.append(
        p(
            "바탕화면의 <b>유상사급 타처보관 확인서</b> 아이콘을 더블 클릭합니다.",
            body_style,
        )
    )

    story.append(p("1-2. \"메일 설정\" 탭을 엽니다", styles["ManualStepHeading"]))
    story.append(
        p(
            "프로그램 위쪽에 [업무] [메일 설정] 두 탭이 있습니다. 오른쪽 <b>메일 설정</b> 을 클릭합니다.",
            body_style,
        )
    )

    story.append(p("1-3. 내 이름을 입력합니다", styles["ManualStepHeading"]))
    story.append(
        p(
            "<b>사용자 이름</b> 칸에 본인 이름을 적습니다 (예: 홍길동). 이 이름이 보내는 메일 본문에 표시됩니다.",
            body_style,
        )
    )

    story.append(p("1-4. Gmail 앱 비밀번호를 새로 만듭니다", styles["ManualStepHeading"]))
    story.append(p("화면에서 파란색 글자로 된 <b>Gmail 앱 비밀번호 발급 페이지 열기</b> 를 클릭합니다.", body_style))
    story.append(p("인터넷 브라우저(크롬 등)가 자동으로 열립니다. 브라우저에서 다음을 진행합니다.", body_style))
    story.append(
        numbered_list(
            [
                "본인 회사 Google 계정으로 로그인합니다.",
                "<b>앱 이름</b> 칸에 알아보기 쉬운 이름을 적습니다 (예: 재고확인서메일).",
                "<b>만들기</b> (또는 Generate) 버튼을 누릅니다.",
                "화면에 16자리 비밀번호가 나타납니다 (예: abcd efgh ijkl mnop).",
                "이 비밀번호를 마우스로 드래그해서 복사합니다 (Ctrl + C).",
            ],
            body_style,
        )
    )
    story.append(
        p(
            "⚠️ <b>중요</b> · 이 비밀번호는 화면이 닫히면 다시 볼 수 없습니다. 닫기 전에 반드시 복사해서 다음 단계로 진행하세요.",
            warn_style,
        )
    )

    story.append(p("1-5. 프로그램에 입력합니다", styles["ManualStepHeading"]))
    story.append(p("프로그램으로 돌아옵니다.", body_style))
    story.append(
        bullet_list(
            [
                "<b>Gmail 주소</b> 칸: 본인 회사 Gmail 주소를 적습니다 (예: hong@fursys.com).",
                "<b>앱 비밀번호</b> 칸: 방금 복사한 16자리를 붙여넣기 합니다 (Ctrl + V). 공백은 자동 정리됩니다.",
            ],
            body_style,
        )
    )

    story.append(p("1-6. 연결이 되는지 확인합니다", styles["ManualStepHeading"]))
    story.append(p("<b>Gmail 연결 테스트</b> 버튼을 누릅니다.", body_style))
    story.append(
        bullet_list(
            [
                "✅ <b>Gmail 로그인 성공</b> 이 나오면 잘 된 것입니다.",
                "❌ 오류가 나오면 1-4 부터 다시 시도합니다.",
            ],
            body_style,
        )
    )

    story.append(p("1-7. 설정을 저장합니다", styles["ManualStepHeading"]))
    story.append(
        p(
            "화면 오른쪽 아래 <b>설정 저장</b> 버튼을 누릅니다. "
            "'메일 설정과 업체 이메일을 저장했습니다.' 라는 메시지가 나오면 끝입니다.",
            body_style,
        )
    )
    story.append(
        p(
            "이제 1부는 모두 끝났습니다. 다음부터는 이 화면을 다시 열 필요가 없습니다.",
            note_style,
        )
    )
    return story


def build_part2_workflow(styles, body_style, note_style, work_image: Path | None):
    story = [PageBreak()]
    story.append(p("2부 · 매번 하는 작업 (확인서 만들고 메일 보내기)", styles["ManualPartHeading"]))
    story.append(p("화면 위쪽에 있는 순서를 그대로 따라가면 됩니다.", body_style))
    story.append(
        p(
            "<b>① 소스 파일 추가 → ② 거래처 불러오기 → ③ 거래처 선택 → ④ 산출 실행 → ⑤ 메일 발송</b>",
            styles["ManualBodyBold"],
        )
    )

    if work_image and work_image.exists():
        story.append(Spacer(1, 3 * mm))
        story.append(scaled_image(work_image, 170 * mm, 95 * mm))
        story.append(p("프로그램 '업무' 탭 화면 구성", note_style))

    story.append(p("① 엑셀 파일 추가하기", styles["ManualStepHeading"]))
    story.append(p("<b>업무</b> 탭이 선택되어 있는지 확인합니다.", body_style))
    story.append(
        p(
            "화면 왼쪽 위 <b>소스 파일</b> 영역에서 <b>파일 추가</b> 버튼을 누릅니다. "
            "받아온 자재유형별수불집계 엑셀 파일을 고릅니다.",
            body_style,
        )
    )
    story.append(
        bullet_list(
            [
                "한 번에 여러 개를 골라도 됩니다 (Ctrl 키를 누르면서 클릭).",
                "잘못 추가한 파일은 그 줄을 클릭한 뒤 <b>선택 제거</b> 버튼으로 뺍니다.",
                "사업장 이름이 다르면 그 칸을 더블 클릭해 직접 고칠 수 있습니다.",
            ],
            body_style,
        )
    )

    story.append(p("② 거래처 불러오기", styles["ManualStepHeading"]))
    story.append(
        p(
            "소스 파일 영역의 오른쪽 아래에 있는 <b>② 거래처 불러오기</b> 버튼을 누릅니다. "
            "잠깐 기다리면 화면 오른쪽에 거래처 목록이 나타납니다.",
            body_style,
        )
    )

    story.append(p("③ 보낼 거래처 고르기", styles["ManualStepHeading"]))
    story.append(
        p(
            "화면 오른쪽 <b>거래처 선택</b> 영역에서, 보내려는 업체 앞의 네모(체크박스)를 클릭해서 ☑로 만듭니다.",
            body_style,
        )
    )
    story.append(
        bullet_list(
            [
                "검색 칸에 업체 이름의 일부를 입력하면 해당 업체만 보입니다.",
                "한 번에 다 선택하려면 <b>현재 목록 전체 선택</b> 을 누릅니다.",
            ],
            body_style,
        )
    )

    story.append(p("③.5 저장 폴더 정하기", styles["ManualStepHeading"]))
    story.append(p("화면 왼쪽 아래 <b>출력 옵션</b> 영역에서:", body_style))
    story.append(
        bullet_list(
            [
                "<b>기준일</b>: 확인서 기준 날짜 (전월 말일이 자동 입력됨).",
                "<b>저장 폴더</b>: <b>폴더 선택</b> 버튼을 눌러 산출물을 저장할 폴더를 고릅니다.",
            ],
            body_style,
        )
    )

    story.append(p("④ 산출 실행 (Excel / PDF 만들기)", styles["ManualStepHeading"]))
    story.append(
        p(
            "화면 맨 아래 오른쪽 끝의 <b>④ 산출 실행</b> 버튼을 누릅니다. "
            "'총 N개 업체의 Excel/PDF를 생성했습니다.' 메시지가 나오면 성공입니다.",
            body_style,
        )
    )
    story.append(
        p(
            "정한 저장 폴더 안에 들어가 보면 업체 이름별 폴더가 만들어져 있고, "
            "각 폴더 안에 엑셀 파일과 PDF 파일이 각각 1개씩 있습니다.",
            body_style,
        )
    )

    story.append(p("⑤ 메일 발송", styles["ManualStepHeading"]))
    story.append(
        p(
            "<b>④ 산출 실행</b> 버튼 오른쪽의 <b>⑤ 메일 발송</b> 버튼을 누르면 새 창이 열립니다.",
            body_style,
        )
    )
    story.append(p("창은 위에서 아래로 세 부분입니다.", body_style))
    story.append(
        info_table(
            [
                ["영역", "내용"],
                ["1단계 · 발송 설정", "발송 방식(보통 그대로), 회신 요청일(다음 주 금요일 자동), 미리보기 업체"],
                ["2단계 · 메일 내용", "왼쪽: 편집할 템플릿 / 오른쪽: 실제 발송될 모양 미리보기"],
                ["3단계 · 발송 대상", "보낼 업체와 이메일 주소 목록, 각 줄의 상태 아이콘"],
            ],
            [40 * mm, 130 * mm],
        )
    )
    story.append(Spacer(1, 3 * mm))
    story.append(p("<b>발송 대상 표의 상태 아이콘 의미</b>", body_style))
    story.append(
        info_table(
            [
                ["아이콘", "의미"],
                ["✅ 발송 가능", "이상 없음. 이대로 발송하면 됨"],
                ["✉️ 이메일 미입력", "그 거래처 이메일을 입력해야 함"],
                ["⚠️ 이메일 형식 오류", "주소를 잘못 적었거나 @가 빠짐"],
                ["📎 첨부파일 없음", "산출 실행이 안 됐음. 창을 닫고 ④ 산출 실행을 다시 함"],
            ],
            [50 * mm, 120 * mm],
        )
    )
    story.append(Spacer(1, 3 * mm))
    story.append(
        p(
            "<b>이메일이 빈 거래처가 있다면</b> 표에서 그 줄의 <b>이메일</b> 칸을 더블 클릭해서 입력합니다. "
            "한 번 입력해 두면 다음번에도 자동으로 채워집니다.",
            body_style,
        )
    )
    story.append(
        p(
            "<b>보내기</b>: 오른쪽 아래 <b>발송</b> 버튼을 누른 뒤, '이대로 발송하시겠습니까?' 가 나오면 <b>예</b> 를 누릅니다. "
            "결과 창에 ✅성공 / ❌실패 / ⏭️제외 건수가 표시됩니다.",
            body_style,
        )
    )
    return story


def build_part3_troubleshooting(styles, body_style, note_style, settings_image: Path | None):
    story = [PageBreak()]
    story.append(p("3부 · 문제가 생겼을 때", styles["ManualPartHeading"]))
    story.append(
        info_table(
            [
                ["증상", "해결 방법"],
                [
                    "거래처가 없거나 빈 화면이 떠요",
                    "① 소스 파일을 먼저 추가했는지, ② 거래처 불러오기를 눌렀는지 확인합니다.",
                ],
                [
                    "⑤ 메일 발송 버튼이 회색이라 안 눌려요",
                    "먼저 ④ 산출 실행을 눌러서 엑셀/PDF를 만들어야 보낼 수 있습니다. "
                    "버튼 위에 마우스를 올리면 무엇이 빠졌는지 알려주는 안내가 나옵니다.",
                ],
                [
                    "일부 거래처가 '이메일 미입력'으로 제외돼요",
                    "메일 발송 화면의 표에서 그 거래처 줄의 이메일 칸을 더블 클릭해서 적습니다. "
                    "한 번 적어두면 다음 작업부터 자동 채움.",
                ],
                [
                    "'Gmail 인증 실패' 라고 나와요",
                    "1부의 1-4 단계부터 다시 진행해서 앱 비밀번호를 새로 만듭니다. "
                    "새 비밀번호 입력 후 Gmail 연결 테스트로 확인.",
                ],
                [
                    "정말로 안 보내져요 (네트워크 등)",
                    "메일 발송 창의 발송 방식을 'EML 초안 파일로 저장' 으로 바꿔서 발송합니다. "
                    "저장 폴더에 _메일초안_EML 폴더가 생기고, .eml 파일을 더블 클릭하면 메일 프로그램에서 첨부까지 다 들어간 채로 열림.",
                ],
                [
                    "잘못 만들어서 다시 만들고 싶어요",
                    "같은 저장 폴더에서 ④ 산출 실행을 다시 누르고 덮어쓰기 확인 시 '예'.",
                ],
            ],
            [55 * mm, 115 * mm],
        )
    )

    if settings_image and settings_image.exists():
        story.append(Spacer(1, 5 * mm))
        story.append(p("[ 메일 설정 화면 구성 참고 ]", styles["ManualStepHeading"]))
        story.append(scaled_image(settings_image, 170 * mm, 100 * mm))
        story.append(
            p(
                "메일 설정 탭의 기본 구성. '고급 옵션'과 'Gmail Google 로그인 (대체 옵션)'은 화살표 ▶ 클릭으로 펼칠 수 있습니다.",
                note_style,
            )
        )

    story.append(Spacer(1, 6 * mm))
    story.append(p("도움이 필요할 때", styles["ManualStepHeading"]))
    story.append(p("이 매뉴얼로 해결이 안 되면 IT 담당자에게 다음 정보를 알려주세요.", body_style))
    story.append(
        bullet_list(
            [
                "어느 단계에서 문제가 생겼는지 (예: ③ 거래처 선택 후 ④ 산출 실행에서 막힘)",
                "화면에 나온 오류 메시지 전체",
                "처음 시도한 거래처 수와 실제 성공/실패 수",
            ],
            body_style,
        )
    )
    return story


def build_story(styles):
    project_root = Path(__file__).resolve().parents[1]
    sample_dir = project_root / "SAMPLE"
    output_dir = project_root / "output" / "doc"

    # ERP 스크린샷은 신규 파일이 있으면 우선, 없으면 기존 파일 사용
    erp_image = first_existing(
        sample_dir / "ERP_자재매출선택.png",
        sample_dir / "ERP 스크린샷.png",
    )
    work_image = first_existing(output_dir / "manual_work_tab.png")
    settings_image = first_existing(output_dir / "manual_settings_tab.png")

    body = styles["ManualBody"]
    note = styles["ManualNote"]
    warn = styles["ManualWarn"]

    story = []
    story += build_title(styles, body, note)
    story += build_requirements(styles, body)
    story += build_part0_erp(styles, body, note, warn, erp_image)
    story += build_part1_setup(styles, body, note, warn)
    story += build_part2_workflow(styles, body, note, work_image)
    story += build_part3_troubleshooting(styles, body, note, settings_image)
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
        topMargin=18 * mm,
        bottomMargin=16 * mm,
        title="유상사급 타처보관 사용 매뉴얼",
    )
    doc.build(build_story(styles), onFirstPage=page_number, onLaterPages=page_number)
    print(f"Generated: {output_path}")


if __name__ == "__main__":
    main()
