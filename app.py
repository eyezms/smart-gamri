import streamlit as st
import json, os, subprocess, tempfile, io, re
from datetime import date
from PIL import Image, ImageOps
from docx import Document
from docx.shared import Mm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from pypdf import PdfWriter

st.set_page_config(page_title="공사감리일지 & 체크리스트", page_icon="🏗️", layout="centered")
st.title("🏗️ 공사감리일지 & 체크리스트 (v2.7)")

with open("checklist_db.json", "r", encoding="utf-8") as f:
    DB = json.load(f)

def parse_item(raw):
    if isinstance(raw, str):
        cat, text = "기본업무", raw
    elif isinstance(raw, dict):
        cat = raw.get("cat") or raw.get("category") or raw.get("구분") or "기본업무"
        text = raw.get("item") or raw.get("text") or raw.get("content") or raw.get("검사항목") or ""
        if not text:
            vals = [str(v) for v in raw.values() if v]
            text = vals[-1] if vals else ""
    else:
        cat, text = "기본업무", str(raw)
    text = re.sub(r'^[☑☐■□✓✔•\-\s]+', '', str(text)).strip()
    return str(cat), text

# 1. 감리 구분 및 공사 개요 (담당자 모두 공란)
st.subheader("1. 감리 구분 및 공사 개요")
gubun = st.radio("감리업무 구분", ["비상주 감리", "책임상주 감리"], horizontal=True)
site_name = st.text_input("공사명", value="서탄산업단지 (주)노바테크놀러지스 신축공사")

c1, c2, c3 = st.columns(3)
work_date = c1.date_input("작성일", value=date.today())
weather = c2.selectbox("날씨", ["맑음", "구름많음", "흐림", "비", "눈"])
location = c3.text_input("시공 층/위치", value="지상층 / 부지내")

c4, c5, c6 = st.columns(3)
inspector = c4.text_input("건축사보(감리원)", value="", placeholder="성명 입력")
chief_supervisor = c5.text_input("총괄감리책임자", value="", placeholder="성명 입력")
contractor = c6.text_input("총괄 시공 책임자", value="", placeholder="성명 입력")

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

# 4. 현장 사진 업로드 (모바일 호환성 강화 + 세션 영구 보관)
st.subheader("4. 📸 현장 사진 업로드 (사진대지 자동 병합)")
if "saved_photos" not in st.session_state:
    st.session_state["saved_photos"] = []

uploaded_photos = st.file_uploader(
    "현장 사진 선택 (갤러리/파일에서 여러 장 선택)",
    accept_multiple_files=True,
    key="photo_uploader"
)
if uploaded_photos:
    new_list = []
    for pf in uploaded_photos:
        try:
            raw_b = pf.getvalue()
            img = ImageOps.exif_transpose(Image.open(io.BytesIO(raw_b))).convert("RGB")
            img.thumbnail((1000, 1000))
            buf = io.BytesIO()
            img.save(buf, format="JPEG", quality=85)
            new_list.append((pf.name, buf.getvalue()))
        except Exception as e:
            st.error(f"이미지 처리 오류 ({pf.name}): {e}")
    if new_list:
        st.session_state["saved_photos"] = new_list

photo_data_list = []
if st.session_state["saved_photos"]:
    col_info, col_clr = st.columns([3, 1])
    col_info.success(f"✅ 총 {len(st.session_state['saved_photos'])}장 사진 준비 완료 (PDF 맨 뒤에 병합됩니다)")
    if col_clr.button("사진 초기화"):
        st.session_state["saved_photos"] = []
        st.rerun()

    cols = st.columns(2)
    for i, (fname, clean_bytes) in enumerate(st.session_state["saved_photos"]):
        with cols[i % 2]:
            st.image(clean_bytes, caption=f"사진 {i+1}", use_container_width=True)
            cap = st.text_input(f"사진 {i+1} 설명", value=f"시공 현황 ({location})", key=f"cap_{i}")
            photo_data_list.append((clean_bytes, cap))

# 5. 결재 도장/서명 첨부
st.subheader("5. ✍️ 결재 도장/서명 첨부")
s1, s2 = st.columns(2)
stamp_insp = s1.file_uploader("건축사보(감리원) 도장/서명", type=["png", "jpg", "jpeg"])
stamp_chief = s2.file_uploader("총괄감리책임자 도장/서명", type=["png", "jpg", "jpeg"])

# --- 원본 법정 서식 정밀 렌더링 유틸리티 ---
def fix_table_layout(table, col_widths_mm):
    table.autofit = False
    tblPr = table._tbl.tblPr
    tblLayout = OxmlElement('w:tblLayout')
    tblLayout.set(qn('w:type'), 'fixed')
    tblPr.append(tblLayout)

    tblGrid = table._tbl.tblGrid
    for child in list(tblGrid):
        tblGrid.remove(child)
    for w_mm in col_widths_mm:
        gridCol = OxmlElement('w:gridCol')
        gridCol.set(qn('w:w'), str(int(w_mm * 56.7)))
        tblGrid.append(gridCol)

    for row in table.rows:
        for idx, w_mm in enumerate(col_widths_mm):
            cell = row.cells[idx]
            cell.width = Mm(w_mm)
            tcPr = cell._tc.get_or_add_tcPr()
            for old in tcPr.findall(qn('w:tcW')):
                tcPr.remove(old)
            tcW = OxmlElement('w:tcW')
            tcW.set(qn('w:w'), str(int(w_mm * 56.7)))
            tcW.set(qn('w:type'), 'dxa')
            tcPr.append(tcW)

def set_row_height(row, height_mm):
    trPr = row._tr.get_or_add_trPr()
    trHeight = OxmlElement('w:trHeight')
    trHeight.set(qn('w:val'), str(int(height_mm * 56.7)))
    trHeight.set(qn('w:hRule'), 'atLeast')
    trPr.append(trHeight)

def set_cell_margins(cell, top=45, bottom=45, left=70, right=70):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def style_cell(cell, text="", size=9, bold=False, align=WD_ALIGN_PARAGRAPH.LEFT, bg=None, img_path=None):
    cell.text = ""
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    set_cell_margins(cell)
    if bg:
        cell._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{bg}"/>'))
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
        p.add_run().add_picture(img_path, width=Mm(12))

def make_doc():
    d = Document()
    for s in d.sections:
        s.page_width, s.page_height = Mm(210), Mm(297)
        s.top_margin, s.bottom_margin, s.left_margin, s.right_margin = Mm(10), Mm(10), Mm(12), Mm(12)
    return d

def build_diary(tmpdir, date_str, si_path, sc_path):
    doc = make_doc()
    p0 = doc.add_paragraph("■ 건축공사 감리세부기준 [별지 제2호서식] <개정 2017. 2. 4.>")
    p0.paragraph_format.space_after = Pt(3)
    p0.runs[0].font.size = Pt(8)

    tbl = doc.add_table(rows=7 + len(check_results), cols=5)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.style = "Table Grid"
    fix_table_layout(tbl, [34, 72, 22, 29, 29])

    set_row_height(tbl.rows[0], 7)
    set_row_height(tbl.rows[1], 14)
    set_row_height(tbl.rows[2], 9)
    set_row_height(tbl.rows[3], 9)
    set_row_height(tbl.rows[4], 10)

    style_cell(tbl.cell(0, 0).merge(tbl.cell(1, 2)), "공  사  감  리  일  지", 18, True, WD_ALIGN_PARAGRAPH.CENTER)
    style_cell(tbl.cell(0, 3), "건축사보", 8.5, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(0, 4), "총괄감리책임자", 8.5, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(1, 3), f"{inspector}\n(서명 또는 인)", 8.5, False, WD_ALIGN_PARAGRAPH.CENTER, img_path=si_path)
    style_cell(tbl.cell(1, 4), f"{chief_supervisor}\n(서명 또는 인)", 8.5, False, WD_ALIGN_PARAGRAPH.CENTER, img_path=sc_path)

    style_cell(tbl.cell(2, 0), "공  사  명", 9.5, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(2, 1).merge(tbl.cell(2, 4)), site_name, 9.5, True)

    style_cell(tbl.cell(3, 0), "작  성  일", 9.5, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(3, 1).merge(tbl.cell(3, 2)), date_str, 9)
    style_cell(tbl.cell(3, 3), "날  씨", 9.5, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(3, 4), weather, 9, False, WD_ALIGN_PARAGRAPH.CENTER)

    style_cell(tbl.cell(4, 0), f"공종 및 세부공정\n({location})", 8.5, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(4, 1).merge(tbl.cell(4, 4)), "감 리 내 용 (감리항목 및 확인결과)", 9.5, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")

    r = 5
    for proc, items in check_results.items():
        lines = [f"• {x['item']} ({x['res']})" for x in items if x['res'] != "해당없음"]
        set_row_height(tbl.rows[r], 45)
        style_cell(tbl.cell(r, 0), proc, 9.5, True, WD_ALIGN_PARAGRAPH.CENTER)
        style_cell(tbl.cell(r, 1).merge(tbl.cell(r, 4)), "\n".join(lines), 8.5)
        r += 1

    set_row_height(tbl.rows[r], 22)
    style_cell(tbl.cell(r, 0), "특 기 사 항", 9.5, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(r, 1).merge(tbl.cell(r, 4)), special_note, 8.5)

    set_row_height(tbl.rows[r + 1], 24)
    style_cell(tbl.cell(r + 1, 0), "지적사항 및\n처리결과", 9, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(r + 1, 1).merge(tbl.cell(r + 1, 4)), issue_note if issue_note.strip() else "\n\n", 8.5)

    p_g = doc.add_paragraph(
        "작성방법\n"
        "1. 공종에는 주요공종 및 단위공종 그리고 해당 층수를 기재합니다.\n"
        "2. 감리항목은 공종별 감리 체크리스트를 기반으로 감리항목을 기재합니다.\n"
        "3. 감리내용에는 각 감리항목에 대하여 육안검사·입회·시험 등을 통한 감리내용을 구체적으로 기재하고, 그 결과와 업무수행 근거자료를 상세히 기재합니다.\n"
        "4. 특기사항은 특별히 명기되어 있지 아니한 내용의 발생·조치사항 등을 기재합니다.\n"
        "5. 지적사항 및 처리결과는 재시공 및 공사중지 등 구두 또는 서면에 의한 지시내용과 처리결과를 기재합니다."
    )
    p_g.paragraph_format.space_before = Pt(4)
    p_g.paragraph_format.line_spacing = 1.08
    p_g.runs[0].font.size = Pt(7.5)

    out = os.path.join(tmpdir, "diary.docx")
    doc.save(out)
    return out

def build_chk(tmpdir, idx, proc, items, short_date):
    doc = make_doc()
    p0 = doc.add_paragraph(f"■ 건축공사 감리세부기준 [별표 1] ({gubun})")
    p0.paragraph_format.space_after = Pt(3)
    p0.runs[0].font.size = Pt(8.5)

    tbl = doc.add_table(rows=6 + len(items), cols=5)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.style = "Table Grid"
    fix_table_layout(tbl, [24, 98, 16, 18, 30])

    set_row_height(tbl.rows[0], 11)
    set_row_height(tbl.rows[1], 9)
    set_row_height(tbl.rows[2], 6)
    set_row_height(tbl.rows[3], 6)

    style_cell(tbl.cell(0, 0).merge(tbl.cell(0, 3)), "공 종 별  감 리  체 크 리 스 트", 15, True, WD_ALIGN_PARAGRAPH.CENTER)
    style_cell(tbl.cell(0, 4), "문서번호 :\n", 8)

    style_cell(tbl.cell(1, 0), "공   종", 9, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(1, 1), proc, 9.5, True)
    style_cell(tbl.cell(1, 2).merge(tbl.cell(1, 3)), "부위/위치", 8.5, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(1, 4), location, 8.5, False, WD_ALIGN_PARAGRAPH.CENTER)

    style_cell(tbl.cell(2, 0).merge(tbl.cell(3, 0)), "구  분", 9, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(2, 1).merge(tbl.cell(3, 1)), "검  사  항  목", 9.5, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(2, 2).merge(tbl.cell(2, 3)), "검사결과", 8.5, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(3, 2), "적합", 8, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(3, 3), "부적합", 8, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(2, 4).merge(tbl.cell(3, 4)), "조치사항", 9, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")

    for i, x in enumerate(items, start=4):
        style_cell(tbl.cell(i, 0), x["cat"], 8, False, WD_ALIGN_PARAGRAPH.CENTER)
        style_cell(tbl.cell(i, 1), x["item"], 8.5)
        style_cell(tbl.cell(i, 2), "✓" if x["res"] == "적합" else "", 9.5, True, WD_ALIGN_PARAGRAPH.CENTER)
        style_cell(tbl.cell(i, 3), "✓" if x["res"] == "부적합" else "", 9.5, True, WD_ALIGN_PARAGRAPH.CENTER)
        style_cell(tbl.cell(i, 4), "", 8)

    r = 4 + len(items)
    set_row_height(tbl.rows[r], 9)
    set_row_height(tbl.rows[r + 1], 9)
    style_cell(tbl.cell(r, 0), "시공자점검일", 8.5, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(r, 1), short_date, 9, False, WD_ALIGN_PARAGRAPH.CENTER)
    style_cell(tbl.cell(r, 2).merge(tbl.cell(r, 4)), f"총괄 시공 책임자 : {contractor}  (서명 또는 인)", 9)
    style_cell(tbl.cell(r + 1, 0), "감리자점검일", 8.5, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(r + 1, 1), short_date, 9, False, WD_ALIGN_PARAGRAPH.CENTER)
    style_cell(tbl.cell(r + 1, 2).merge(tbl.cell(r + 1, 4)), f"건축사보(감리원) : {inspector}  (서명 또는 인)", 9)

    out = os.path.join(tmpdir, f"chk_{idx}.docx")
    doc.save(out)
    return out

def build_photos(p_list, tmpdir):
    doc = make_doc()
    p_title = doc.add_paragraph("■ 현장 감리 사진대지")
    p_title.runs[0].font.size = Pt(12)
    p_title.runs[0].bold = True

    num_rows = (len(p_list) + 1) // 2
    tbl = doc.add_table(rows=num_rows, cols=2)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.style = "Table Grid"
    fix_table_layout(tbl, [93, 93])

    for i in range(0, len(p_list), 2):
        row = tbl.rows[i // 2].cells
        for j in range(2):
            if i + j < len(p_list):
                clean_bytes, cap_txt = p_list[i + j]
                pp = os.path.join(tmpdir, f"clean_p_{i+j}.jpg")
                with open(pp, "wb") as f_out:
                    f_out.write(clean_bytes)

                set_cell_margins(row[j], 45, 45, 45, 45)
                p_img = row[j].paragraphs[0]
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img.add_run().add_picture(pp, width=Mm(82))

                p_cap = row[j].add_paragraph(f"[사진 {i+j+1}] {cap_txt}")
                p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_cap.runs[0].font.size = Pt(8.5)

    out = os.path.join(tmpdir, "photos.docx")
    doc.save(out)
    return out

if st.button("📄 통합 감리일지 PDF 생성", type="primary", use_container_width=True):
    if not selected_procs:
        st.warning("최소 1개 이상의 공종을 선택해 주세요.")
    else:
        with st.spinner("통합 PDF 문서를 생성 중입니다..."):
            with tempfile.TemporaryDirectory() as tmpdir:
                wk = ["월", "화", "수", "목", "금", "토", "일"][work_date.weekday()]
                date_str = f"{work_date.year}년 {work_date.month:02d}월 {work_date.day:02d}일 ({wk}요일)"
                short_date = f"{work_date.year}년 {work_date.month:02d}월 {work_date.day:02d}일"

                si_p, sc_p = None, None
                if stamp_insp:
                    si_p = os.path.join(tmpdir, "si.png")
                    Image.open(io.BytesIO(stamp_insp.getvalue())).save(si_p)
                if stamp_chief:
                    sc_p = os.path.join(tmpdir, "sc.png")
                    Image.open(io.BytesIO(stamp_chief.getvalue())).save(sc_p)

                docs = [build_diary(tmpdir, date_str, si_p, sc_p)]
                for idx, (proc, items) in enumerate(check_results.items()):
                    docs.append(build_chk(tmpdir, idx, proc, items, short_date))
                if photo_data_list:
                    docs.append(build_photos(photo_data_list, tmpdir))

                writer = PdfWriter()
                for dp in docs:
                    subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", tmpdir, dp], check=True)
                    writer.append(os.path.splitext(dp)[0] + ".pdf")

                final_pdf = os.path.join(tmpdir, "final.pdf")
                writer.write(final_pdf)
                with open(final_pdf, "rb") as fp:
                    st.download_button("⬇️ 완성된 통합 PDF 다운로드", fp.read(), file_name=f"[별지2]공사감리일지({work_date}).pdf", mime="application/pdf", use_container_width=True)