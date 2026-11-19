import streamlit as st
import json, os, subprocess, tempfile
from datetime import date
from docx import Document
from docx.shared import Mm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from pypdf import PdfWriter

st.set_page_config(page_title="공사감리일지 & 체크리스트", page_icon="🏗️", layout="centered")
st.title("🏗️ 공사감리일지 & 체크리스트 (v2.1)")

with open("checklist_db.json", "r", encoding="utf-8") as f:
    DB = json.load(f)

def parse_item(raw):
    if isinstance(raw, str):
        return "기본업무", raw
    if isinstance(raw, dict):
        cat = raw.get("cat") or raw.get("category") or raw.get("구분") or raw.get("type") or "기본업무"
        text = raw.get("item") or raw.get("text") or raw.get("content") or raw.get("검사항목") or raw.get("title")
        if not text:
            vals = [str(v) for v in raw.values() if v]
            text = vals[-1] if vals else ""
        return str(cat), str(text)
    return "기본업무", str(raw)

# 1. 감리 구분 및 공사 개요 (담당자 이름 모두 공란 처리)
st.subheader("1. 감리 구분 및 공사 개요")
gubun = st.radio("감리업무 구분", ["비상주 감리", "책임상주 감리"], horizontal=True)
site_name = st.text_input("공사명", value="서탄산업단지 (주)노바테크놀러지스 신축공사")

c1, c2, c3 = st.columns(3)
work_date = c1.date_input("작성일", value=date.today())
weather = c2.selectbox("날씨", ["맑음", "구름많음", "흐림", "비", "눈"])
location = c3.text_input("시공 층/위치", value="지상층 / 부지내")

c4, c5, c6 = st.columns(3)
inspector = c4.text_input("건축사보(감리원)", value="", placeholder="성명 직접 입력")
chief_supervisor = c5.text_input("총괄감리책임자", value="", placeholder="성명 직접 입력")
contractor = c6.text_input("총괄 시공 책임자", value="", placeholder="성명 직접 입력")

# 2. 공종별 체크리스트
st.subheader(f"2. [{gubun}] 공종별 체크리스트")
selected_procs = st.multiselect("금일 감리 공종 선택", list(DB[gubun].keys()))

check_results = {}
for proc in selected_procs:
    with st.expander(f"📋 {proc} 점검항목", expanded=True):
        check_results[proc] = []
        for idx, raw_item in enumerate(DB[gubun][proc]):
            cat_str, item_str = parse_item(raw_item)
            col_t, col_r = st.columns([3, 1])
            col_t.write(f"**[{cat_str}]** {item_str}")
            res = col_r.selectbox("결과", ["적합", "부적합", "해당없음"], key=f"{proc}_{idx}", label_visibility="collapsed")
            check_results[proc].append({"cat": cat_str, "item": item_str, "res": res})

# 3. 특기사항 및 지적사항
st.subheader("3. 특기사항 및 지적사항")
special_note = st.text_area("특기사항", value="- 해당공정 인부 안전관리 철저 주문\n- 용접공정이나 기타 인화물질자재 사용시 화재 안전기준 철저 주문")
issue_note = st.text_area("지적사항 및 처리결과", value="")

# 4. 현장 사진 업로드
st.subheader("4. 📸 현장 사진 업로드 (사진대지 자동 병합)")
photos = st.file_uploader("현장 사진 선택 (갤러리에서 여러 장 선택 가능)", type=["png", "jpg", "jpeg"], accept_multiple_files=True)
if photos:
    st.success(f"총 {len(photos)}장의 현장 사진이 첨부되었습니다.")

# 5. 서명(도장) 첨부
st.subheader("5. ✍️ 결재 도장/서명 첨부")
s1, s2 = st.columns(2)
stamp_insp = s1.file_uploader("건축사보(감리원) 도장/서명", type=["png", "jpg", "jpeg"])
stamp_chief = s2.file_uploader("총괄감리책임자 도장/서명", type=["png", "jpg", "jpeg"])

# --- 워드 문서 생성 유틸리티 ---
def set_cell_bg(cell, hex_color):
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shading)

def set_cell_margins(cell, top=35, bottom=35, left=75, right=75):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def style_cell(cell, text="", size=9, bold=False, align=WD_ALIGN_PARAGRAPH.LEFT, bg=None, img_path=None, img_w=12):
    cell.text = ""
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    set_cell_margins(cell)
    if bg:
        set_cell_bg(cell, bg)
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.12
    if text:
        run = p.add_run(text)
        run.font.name = "NanumGothic"
        run._element.rPr.rFonts.set(qn('w:eastAsia'), "NanumGothic")
        run.font.size = Pt(size)
        run.bold = bold
    if img_path and os.path.exists(img_path):
        run_img = p.add_run()
        run_img.add_picture(img_path, width=Mm(img_w))

def build_diary_docx(tmpdir, date_str, si_path, sc_path):
    doc = Document()
    for sec in doc.sections:
        sec.top_margin = Mm(10)
        sec.bottom_margin = Mm(10)
        sec.left_margin = Mm(12)
        sec.right_margin = Mm(12)

    p_top = doc.add_paragraph("■ 건축공사 감리세부기준 [별지 제2호서식] <개정 2017. 2. 4.>")
    p_top.paragraph_format.space_after = Pt(2)
    p_top.runs[0].font.size = Pt(8)

    tbl = doc.add_table(rows=7 + len(check_results), cols=4)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.style = "Table Grid"
    widths = [Mm(32), Mm(94), Mm(30), Mm(30)]
    for row in tbl.rows:
        for idx, w in enumerate(widths):
            row.cells[idx].width = w

    c_title = tbl.cell(0, 0).merge(tbl.cell(1, 1))
    style_cell(c_title, "공 사 감 리 일 지", size=16, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    style_cell(tbl.cell(0, 2), "건축사보", size=9, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, bg="F2F2F2")
    style_cell(tbl.cell(0, 3), "총괄감리책임자", size=9, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, bg="F2F2F2")
    style_cell(tbl.cell(1, 2), f"{inspector}\n(서명 또는 인)", size=8.5, align=WD_ALIGN_PARAGRAPH.CENTER, img_path=si_path, img_w=11)
    style_cell(tbl.cell(1, 3), f"{chief_supervisor}\n(서명 또는 인)", size=8.5, align=WD_ALIGN_PARAGRAPH.CENTER, img_path=sc_path, img_w=11)

    style_cell(tbl.cell(2, 0), "공 사 명", size=9, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, bg="F2F2F2")
    style_cell(tbl.cell(2, 1).merge(tbl.cell(2, 3)), site_name, size=9.5, bold=True)

    style_cell(tbl.cell(3, 0), "작 성 일", size=9, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, bg="F2F2F2")
    style_cell(tbl.cell(3, 1), date_str, size=9)
    style_cell(tbl.cell(3, 2), "날 씨", size=9, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, bg="F2F2F2")
    style_cell(tbl.cell(3, 3), weather, size=9, align=WD_ALIGN_PARAGRAPH.CENTER)

    style_cell(tbl.cell(4, 0), f"공종 및 세부공정\n({location})", size=9, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, bg="F2F2F2")
    style_cell(tbl.cell(4, 1).merge(tbl.cell(4, 3)), "감리내용 (감리항목 및 확인결과)", size=9, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, bg="F2F2F2")

    r_idx = 5
    for proc, items in check_results.items():
        lines = [f"• {x['item']} ({x['res']})" for x in items if x['res'] != "해당없음"]
        style_cell(tbl.cell(r_idx, 0), proc, size=9, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
        style_cell(tbl.cell(r_idx, 1).merge(tbl.cell(r_idx, 3)), "\n".join(lines), size=8.5)
        r_idx += 1

    style_cell(tbl.cell(r_idx, 0), "특 기 사 항", size=9, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, bg="F2F2F2")
    style_cell(tbl.cell(r_idx, 1).merge(tbl.cell(r_idx, 3)), special_note, size=8.5)
    r_idx += 1

    style_cell(tbl.cell(r_idx, 0), "지적사항 및\n처리결과", size=9, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, bg="F2F2F2")
    style_cell(tbl.cell(r_idx, 1).merge(tbl.cell(r_idx, 3)), issue_note, size=8.5)

    p_guide = doc.add_paragraph(
        "작성방법\n"
        "1. 공종에는 주요공종 및 단위공종 그리고 해당 층수를 기재합니다.\n"
        "2. 감리항목은 공종별 감리 체크리스트를 기반으로 감리항목을 기재합니다.\n"
        "3. 감리내용에는 각 감리항목에 대하여 육안검사·입회·시험 등을 통한 감리내용을 구체적으로 기재하고, 그 결과와 업무수행 근거자료를 상세히 기재합니다.\n"
        "4. 특기사항은 특별히 명기되어 있지 아니한 내용의 발생·조치사항 등을 기재합니다.\n"
        "5. 지적사항 및 처리결과는 재시공 및 공사중지 등 구두 또는 서면에 의한 지시내용과 처리결과를 기재합니다."
    )
    p_guide.paragraph_format.space_before = Pt(3)
    p_guide.paragraph_format.line_spacing = 1.05
    for r in p_guide.runs:
        r.font.size = Pt(7.5)

    out_path = os.path.join(tmpdir, "diary.docx")
    doc.save(out_path)
    return out_path

def build_checklist_docx(tmpdir, idx, proc, items, short_date):
    doc = Document()
    for sec in doc.sections:
        sec.top_margin = Mm(10)
        sec.bottom_margin = Mm(10)
        sec.left_margin = Mm(12)
        sec.right_margin = Mm(12)

    p_top = doc.add_paragraph(f"■ 건축공사 감리세부기준 [별표 1] ({gubun})")
    p_top.paragraph_format.space_after = Pt(2)
    p_top.runs[0].font.size = Pt(8.5)

    tbl = doc.add_table(rows=5 + len(items), cols=5)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.style = "Table Grid"
    w2 = [Mm(24), Mm(96), Mm(16), Mm(18), Mm(32)]
    for row in tbl.rows:
        for i, w in enumerate(w2):
            row.cells[i].width = w

    style_cell(tbl.cell(0, 0).merge(tbl.cell(0, 3)), "공 종 별  감 리  체 크 리 스 트", size=14, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    style_cell(tbl.cell(0, 4), "문서번호 :\n", size=8)

    style_cell(tbl.cell(1, 0), "공 종", size=9, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, bg="F2F2F2")
    style_cell(tbl.cell(1, 1), proc, size=9, bold=True)
    style_cell(tbl.cell(1, 2).merge(tbl.cell(1, 3)), "부위/위치", size=9, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, bg="F2F2F2")
    style_cell(tbl.cell(1, 4), location, size=9)

    style_cell(tbl.cell(2, 0), "구 분", size=9, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, bg="F2F2F2")
    style_cell(tbl.cell(2, 1), "검 사 항 목", size=9, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, bg="F2F2F2")
    style_cell(tbl.cell(2, 2), "적합", size=8.5, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, bg="F2F2F2")
    style_cell(tbl.cell(2, 3), "부적합", size=8.5, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, bg="F2F2F2")
    style_cell(tbl.cell(2, 4), "조치사항", size=9, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, bg="F2F2F2")

    r_idx = 3
    for x in items:
        style_cell(tbl