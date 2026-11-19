import streamlit as st
import json, os, subprocess, tempfile
from datetime import date
from docx import Document
from docx.shared import Mm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from pypdf import PdfWriter

st.set_page_config(page_title="공사감리일지 & 체크리스트", page_icon="🏗️", layout="centered")
st.title("🏗️ 공사감리일지 & 체크리스트 (v2.2)")

with open("checklist_db.json", "r", encoding="utf-8") as f:
    DB = json.load(f)

def parse_item(raw):
    if isinstance(raw, str):
        return "기본업무", raw
    if isinstance(raw, dict):
        cat = raw.get("cat") or raw.get("category") or raw.get("구분") or "기본업무"
        text = raw.get("item") or raw.get("text") or raw.get("content") or raw.get("검사항목") or ""
        if not text:
            vals = [str(v) for v in raw.values() if v]
            text = vals[-1] if vals else ""
        return str(cat), str(text)
    return "기본업무", str(raw)

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

st.subheader("3. 특기사항 및 지적사항")
special_note = st.text_area("특기사항", value="- 해당공정 인부 안전관리 철저 주문\n- 용접공정이나 기타 인화물질자재 사용시 화재 안전기준 철저 주문")
issue_note = st.text_area("지적사항 및 처리결과", value="")

st.subheader("4. 📸 현장 사진 업로드 (사진대지 자동 병합)")
photos = st.file_uploader("현장 사진 선택 (여러 장 가능)", type=["png", "jpg", "jpeg"], accept_multiple_files=True)
if photos:
    st.success(f"총 {len(photos)}장 첨부 완료")

st.subheader("5. ✍️ 결재 도장/서명 첨부")
s1, s2 = st.columns(2)
stamp_insp = s1.file_uploader("건축사보(감리원) 도장/서명", type=["png", "jpg", "jpeg"])
stamp_chief = s2.file_uploader("총괄감리책임자 도장/서명", type=["png", "jpg", "jpeg"])

def style_cell(cell, text="", size=9, bold=False, align=WD_ALIGN_PARAGRAPH.LEFT, bg=None, img_path=None):
    cell.text = ""
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    if bg:
        cell._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{bg}"/>'))
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(1)
    if text:
        run = p.add_run(text)
        run.font.name = "NanumGothic"
        run._element.rPr.rFonts.set(qn('w:eastAsia'), "NanumGothic")
        run.font.size = Pt(size)
        run.bold = bold
    if img_path and os.path.exists(img_path):
        p.add_run().add_picture(img_path, width=Mm(11))

def make_doc():
    d = Document()
    for s in d.sections:
        s.top_margin, s.bottom_margin, s.left_margin, s.right_margin = Mm(10), Mm(10), Mm(12), Mm(12)
    return d

def build_diary(tmpdir, date_str, si_path, sc_path):
    doc = make_doc()
    doc.add_paragraph("■ 건축공사 감리세부기준 [별지 제2호서식]").runs[0].font.size = Pt(8)
    tbl = doc.add_table(rows=7 + len(check_results), cols=4)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.style = "Table Grid"
    for row in tbl.rows:
        for i, w in enumerate([Mm(32), Mm(94), Mm(30), Mm(30)]):
            row.cells[i].width = w

    style_cell(tbl.cell(0, 0).merge(tbl.cell(1, 1)), "공 사 감 리 일 지", 16, True, WD_ALIGN_PARAGRAPH.CENTER)
    style_cell(tbl.cell(0, 2), "건축사보", 9, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(0, 3), "총괄감리책임자", 9, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(1, 2), f"{inspector}\n(서명 또는 인)", 8.5, False, WD_ALIGN_PARAGRAPH.CENTER, img_path=si_path)
    style_cell(tbl.cell(1, 3), f"{chief_supervisor}\n(서명 또는 인)", 8.5, False, WD_ALIGN_PARAGRAPH.CENTER, img_path=sc_path)

    style_cell(tbl.cell(2, 0), "공 사 명", 9, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(2, 1).merge(tbl.cell(2, 3)), site_name, 9.5, True)
    style_cell(tbl.cell(3, 0), "작 성 일", 9, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(3, 1), date_str, 9)
    style_cell(tbl.cell(3, 2), "날 씨", 9, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(3, 3), weather, 9, False, WD_ALIGN_PARAGRAPH.CENTER)
    style_cell(tbl.cell(4, 0), f"공종 및 세부공정\n({location})", 9, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(4, 1).merge(tbl.cell(4, 3)), "감리내용 (감리항목 및 확인결과)", 9, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")

    r = 5
    for proc, items in check_results.items():
        lines = [f"• {x['item']} ({x['res']})" for x in items if x['res'] != "해당없음"]
        style_cell(tbl.cell(r, 0), proc, 9, True, WD_ALIGN_PARAGRAPH.CENTER)
        style_cell(tbl.cell(r, 1).merge(tbl.cell(r, 3)), "\n".join(lines), 8.5)
        r += 1

    style_cell(tbl.cell(r, 0), "특 기 사 항", 9, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(r, 1).merge(tbl.cell(r, 3)), special_note, 8.5)
    style_cell(tbl.cell(r + 1, 0), "지적사항 및\n처리결과", 9, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(r + 1, 1).merge(tbl.cell(r + 1, 3)), issue_note, 8.5)

    p_g = doc.add_paragraph("작성방법: 1.주요/단위공종 및 층수 기재  2.체크리스트 기반 감리항목 기재  3.육안·입회·시험 결과 상세 기재  4.특기사항  5.지적사항 및 처리결과")
    p_g.runs[0].font.size = Pt(7.5)
    out = os.path.join(tmpdir, "diary.docx")
    doc.save(out)
    return out

def build_chk(tmpdir, idx, proc, items, short_date):
    doc = make_doc()
    doc.add_paragraph(f"■ 건축공사 감리세부기준 [별표 1] ({gubun})").runs[0].font.size = Pt(8.5)
    tbl = doc.add_table(rows=5 + len(items), cols=5)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.style = "Table Grid"
    for row in tbl.rows:
        for i, w in enumerate([Mm(24), Mm(96), Mm(16), Mm(18), Mm(32)]):
            row.cells[i].width = w

    style_cell(tbl.cell(0, 0).merge(tbl.cell(0, 3)), "공 종 별  감 리  체 크 리 스 트", 14, True, WD_ALIGN_PARAGRAPH.CENTER)
    style_cell(tbl.cell(0, 4), "문서번호 :\n", 8)
    style_cell(tbl.cell(1, 0), "공 종", 9, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(1, 1), proc, 9, True)
    style_cell(tbl.cell(1, 2).merge(tbl.cell(1, 3)), "부위/위치", 9, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(1, 4), location, 9)
    style_cell(tbl.cell(2, 0), "구 분", 9, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(2, 1), "검 사 항 목", 9, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(2, 2), "적합", 8.5, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(2, 3), "부적합", 8.5, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(2, 4), "조치사항", 9, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")

    for i, x in enumerate(items, start=3):
        style_cell(tbl.cell(i, 0), x["cat"], 8, False, WD_ALIGN_PARAGRAPH.CENTER)
        style_cell(tbl.cell(i, 1), x["item"], 8.5)
        style_cell(tbl.cell(i, 2), "✓" if x["res"] == "적합" else "", 9, True, WD_ALIGN_PARAGRAPH.CENTER)
        style_cell(tbl.cell(i, 3), "✓" if x["res"] == "부적합" else "", 9, True, WD_ALIGN_PARAGRAPH.CENTER)
        style_cell(tbl.cell(i, 4), "", 8)

    r = 3 + len(items)
    style_cell(tbl.cell(r, 0), "시공자점검일", 8.5, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(r, 1), short_date, 9, False, WD_ALIGN_PARAGRAPH.CENTER)
    style_cell(tbl.cell(r, 2).merge(tbl.cell(r, 4)), f"총괄 시공 책임자 : {contractor} (서명 또는 인)", 9)
    style_cell(tbl.cell(r + 1, 0), "감리자점검일", 8.5, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(r + 1, 1), short_date, 9, False, WD_ALIGN_PARAGRAPH.CENTER)
    style_cell(tbl.cell(r + 1, 2).merge(tbl.cell(r + 1, 4)), f"건축사보(감리원) : {inspector} (서명 또는 인)", 9)

    out = os.path.join(tmpdir, f"chk_{idx}.docx")
    doc.save(out)
    return out

def build_photos(photo_files, tmpdir):
    doc = make_doc()
    doc.add_heading("현장 감리 사진대지", level=1)
    tbl = doc.add_table(rows=0, cols=2)
    tbl.style = "Table Grid"
    for i in range(0, len(photo_files), 2):
        row = tbl.add_row().cells
        for j in range(2):
            if i + j < len(photo_files):
                pf = photo_files[i + j]
                pp = os.path.join(tmpdir, f"p_{i+j}_{pf.name}")
                with open(pp, "wb") as f:
                    f.write(pf.getbuffer())
                row[j].paragraphs[0].add_run().add_picture(pp, width=Mm(78))
                row[j].add_paragraph(f"사진 {i+j+1}. {pf.name}")
    out = os.path.join(tmpdir, "photos.docx")
    doc.save(out)
    return out

if st.button("📄 통합 감리일지 PDF 생성", type="primary", use_container_width=True):
    if not selected_procs:
        st.warning("최소 1개 이상의 공종을 선택해 주세요.")
    else:
        with st.spinner("PDF 생성 중..."):
            with tempfile.TemporaryDirectory() as tmpdir:
                wk = ["월", "화", "수", "목", "금", "토", "일"][work_date.weekday()]
                date_str = f"{work_date.year}년 {work_date.month:02d}월 {work_date.day:02d}일 ({wk}요일)"
                short_date = f"{work_date.year}년 {work_date.month:02d}월 {work_date.day:02d}일"

                si_p, sc_p = None, None
                if stamp_insp:
                    si_p = os.path.join(tmpdir, "si.png")
                    with open(si_p, "wb") as f: f.write(stamp_insp.getbuffer())
                if stamp_chief:
                    sc_p = os.path.join(tmpdir, "sc.png")
                    with open(sc_p, "wb") as f: f.write(stamp_chief.getbuffer())

                docs = [build_diary(tmpdir, date_str, si_p, sc_p)]
                for idx, (proc, items) in enumerate(check_results.items()):
                    docs.append(build_chk(tmpdir, idx, proc, items, short_date))
                if photos:
                    docs.append(build_photos(photos, tmpdir))

                writer = PdfWriter()
                for dp in docs:
                    subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", tmpdir, dp], check=True)
                    writer.append(os.path.splitext(dp)[0] + ".pdf")

                final_pdf = os.path.join(tmpdir, "final.pdf")
                writer.write(final_pdf)
                with open(final_pdf, "rb") as fp:
                    st.download_button("⬇️ 완성된 통합 PDF 다운로드", fp.read(), file_name=f"[별지2]공사감리일지({work_date}).pdf", mime="application/pdf", use_container_width=True)