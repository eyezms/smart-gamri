import streamlit as st
import json, os, subprocess, tempfile
from datetime import date
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from docx import Document
from pypdf import PdfWriter

st.set_page_config(page_title="공사감리일지 & 체크리스트", page_icon="🏗️", layout="centered")
st.title("🏗️ 공사감리일지 & 체크리스트 앱")

with open("checklist_db.json", "r", encoding="utf-8") as f:
    DB = json.load(f)

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
        for idx, item in enumerate(DB[gubun][proc]):
            col_t, col_r = st.columns([3, 1])
            col_t.write(f"**[{item['cat']}]** {item['item']}")
            res = col_r.selectbox("결과", ["적합", "부적합", "해당없음"], key=f"{proc}_{idx}", label_visibility="collapsed")
            check_results[proc].append({"cat": item["cat"], "item": item["item"], "res": res})

# 3. 특기사항 및 지적사항
st.subheader("3. 특기사항 및 지적사항")
special_note = st.text_area("특기사항", value="- 해당공정 인부 안전관리 철저 주문\n- 용접공정이나 기타 인화물질자재 사용시 화재 안전기준 철저 주문")
issue_note = st.text_area("지적사항 및 처리결과", value="")

# 4. 서명(도장) 및 현장 사진 첨부
st.subheader("4. 서명(도장) 및 현장 사진 첨부")
s1, s2 = st.columns(2)
stamp_insp = s1.file_uploader("건축사보(감리원) 도장/서명", type=["png", "jpg", "jpeg"])
stamp_chief = s2.file_uploader("총괄감리책임자 도장/서명", type=["png", "jpg", "jpeg"])
photos = st.file_uploader("현장 사진 첨부 (여러 장 선택 가능)", type=["png", "jpg", "jpeg"], accept_multiple_files=True)

def build_photo_docx(photo_files, tmpdir):
    doc = Document()
    for sec in doc.sections:
        sec.top_margin = Mm(12)
        sec.bottom_margin = Mm(12)
        sec.left_margin = Mm(12)
        sec.right_margin = Mm(12)
    doc.add_heading("현장 감리 사진대지", level=1)
    table = doc.add_table(rows=0, cols=2)
    table.style = "Table Grid"
    for i in range(0, len(photo_files), 2):
        row = table.add_row().cells
        for j in range(2):
            if i + j < len(photo_files):
                p_file = photo_files[i + j]
                p_path = os.path.join(tmpdir, f"photo_{i+j}_{p_file.name}")
                with open(p_path, "wb") as pf:
                    pf.write(p_file.getbuffer())
                run = row[j].paragraphs[0].add_run()
                run.add_picture(p_path, width=Mm(78))
                row[j].add_paragraph(f"사진 {i+j+1}. {p_file.name}")
    out_docx = os.path.join(tmpdir, "photos.docx")
    doc.save(out_docx)
    return out_docx

if st.button("📄 통합 감리일지 PDF 생성", type="primary", use_container_width=True):
    if not selected_procs:
        st.warning("최소 1개 이상의 공종을 선택해 주세요.")
    else:
        with st.spinner("PDF 문서를 생성 중입니다..."):
            with tempfile.TemporaryDirectory() as tmpdir:
                weekdays = ["월", "화", "수", "목", "금", "토", "일"]
                date_str = f"{work_date.year}년 {work_date.month:02d}월 {work_date.day:02d}일 ({weekdays[work_date.weekday()]}요일)"
                
                diary_items = []
                for proc, items in check_results.items():
                    lines = [f"• {x['item']} ({x['res']})" for x in items if x['res'] != "해당없음"]
                    diary_items.append({"proc": proc, "detail": "\n".join(lines)})

                tpl_diary = DocxTemplate("template_diary.docx")
                ctx_diary = {
                    "site_name": site_name, "date_str": date_str, "weather": weather,
                    "location": location, "inspector": inspector, "chief_supervisor": chief_supervisor,
                    "special_note": special_note, "issue_note": issue_note, "items": diary_items,
                    "stamp_insp": "", "stamp_chief": ""
                }
                if stamp_insp:
                    p_si = os.path.join(tmpdir, "si.png")
                    with open(p_si, "wb") as f_si: f_si.write(stamp_insp.getbuffer())
                    ctx_diary["stamp_insp"] = InlineImage(tpl_diary, p_si, width=Mm(12))
                if stamp_chief:
                    p_sc = os.path.join(tmpdir, "sc.png")
                    with open(p_sc, "wb") as f_sc: f_sc.write(stamp_chief.getbuffer())
                    ctx_diary["stamp_chief"] = InlineImage(tpl_diary, p_sc, width=Mm(12))

                diary_docx = os.path.join(tmpdir, "diary.docx")
                tpl_diary.render(ctx_diary)
                tpl_diary.save(diary_docx)

                docx_list = [diary_docx]
                for idx, (proc, items) in enumerate(check_results.items()):
                    tpl_chk = DocxTemplate("template_checklist.docx")
                    chk_rows = [{
                        "cat": x["cat"], "item": x["item"],
                        "ok": "✓" if x["res"] == "적합" else "",
                        "no": "✓" if x["res"] == "부적합" else "",
                        "action": ""
                    } for x in items]
                    ctx_chk = {
                        "gubun": gubun, "proc": proc, "location": location,
                        "date_str": f"{work_date.year}년 {work_date.month:02d}월 {work_date.day:02d}일",
                        "contractor": contractor, "inspector": inspector, "rows": chk_rows
                    }
                    c_docx = os.path.join(tmpdir, f"chk_{idx}.docx")
                    tpl_chk.render(ctx_chk)
                    tpl_chk.save(c_docx)
                    docx_list.append(c_docx)

                if photos:
                    docx_list.append(build_photo_docx(photos, tmpdir))

                writer = PdfWriter()
                for d_path in docx_list:
                    subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", tmpdir, d_path], check=True)
                    pdf_path = os.path.splitext(d_path)[0] + ".pdf"
                    writer.append(pdf_path)

                final_pdf = os.path.join(tmpdir, "final.pdf")
                writer.write(final_pdf)
                with open(final_pdf, "rb") as fp:
                    st.download_button("⬇️ 완성된 통합 PDF 다운로드", fp.read(), file_name=f"[별지2]공사감리일지({work_date}).pdf", mime="application/pdf", use_container_width=True)