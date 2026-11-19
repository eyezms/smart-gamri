import streamlit as st
from docxtpl import DocxTemplate, InlineImage
from docx import Document
from docx.shared import Mm, Pt
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from pypdf import PdfWriter
from datetime import date
import subprocess
import tempfile
import json
import io
import os

st.set_page_config(page_title="스마트 공사감리일지", page_icon="🏗️", layout="centered")

WEEKDAYS = ["월요일", "화요일", "수요일", "목요일", "금요일", "토요일", "일요일"]

@st.cache_data
def load_checklist_db():
    if os.path.exists("checklist_db.json"):
        with open("checklist_db.json", "r", encoding="utf-8") as f:
            return json.load(f)
    return {
        "비상주 감리": {"가설공사": [{"업무구분": "기본업무", "항목": "대지의 고저차 설계도서 확인"}]},
        "책임상주 감리": {"가설공사": [{"업무구분": "기본업무", "항목": "대지의 고저차 설계도서 확인"}]}
    }

CHECKLIST_DB = load_checklist_db()

def create_photo_sheet_docx(photos_with_captions, save_path):
    doc = Document()
    for section in doc.sections:
        section.top_margin = Mm(15)
        section.bottom_margin = Mm(15)
        section.left_margin = Mm(15)
        section.right_margin = Mm(15)

    for page_idx in range(0, len(photos_with_captions), 6):
        if page_idx > 0:
            doc.add_page_break()
            
        chunk = photos_with_captions[page_idx:page_idx + 6]
        table = doc.add_table(rows=6, cols=2)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.style = "Table Grid"

        for i, (img_bytes, caption) in enumerate(chunk):
            row_img = (i // 2) * 2
            row_cap = row_img + 1
            col = i % 2

            cell_img = table.cell(row_img, col)
            p_img = cell_img.paragraphs[0]
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.add_run().add_picture(io.BytesIO(img_bytes), width=Mm(82), height=Mm(58))

            cell_cap = table.cell(row_cap, col)
            p_cap = cell_cap.paragraphs[0]
            p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run_cap = p_cap.add_run(caption if caption else "현장 확인 사진")
            run_cap.font.name = "맑은 고딕"
            run_cap.font.size = Pt(10)

        p_footer = doc.add_paragraph("현 장 사 진")
        p_footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_footer.runs[0].font.bold = True

    doc.save(save_path)

def build_combined_pdf(diary_ctx, checklist_ctx_list, photos_list, inspector_stamp, contractor_stamp):
    with tempfile.TemporaryDirectory() as tmpdir:
        pdf_parts = []

        if os.path.exists("template_diary.docx"):
            doc_diary = DocxTemplate("template_diary.docx")
            diary_ctx["inspector_sign"] = InlineImage(doc_diary, io.BytesIO(inspector_stamp), width=Mm(16)) if inspector_stamp else "(인)"
            doc_diary.render(diary_ctx)
            
            diary_docx = os.path.join(tmpdir, "part1_diary.docx")
            doc_diary.save(diary_docx)
            subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", tmpdir, diary_docx], check=True)
            pdf_parts.append(os.path.join(tmpdir, "part1_diary.pdf"))

        if photos_list:
            photo_docx = os.path.join(tmpdir, "part2_photos.docx")
            create_photo_sheet_docx(photos_list, photo_docx)
            subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", tmpdir, photo_docx], check=True)
            pdf_parts.append(os.path.join(tmpdir, "part2_photos.pdf"))

        if os.path.exists("template_checklist.docx"):
            for idx, chk_ctx in enumerate(checklist_ctx_list):
                doc_chk = DocxTemplate("template_checklist.docx")
                chk_ctx["inspector_sign"] = InlineImage(doc_chk, io.BytesIO(inspector_stamp), width=Mm(15)) if inspector_stamp else "(인)"
                chk_ctx["contractor_sign"] = InlineImage(doc_chk, io.BytesIO(contractor_stamp), width=Mm(15)) if contractor_stamp else "(인)"
                doc_chk.render(chk_ctx)
                
                chk_docx = os.path.join(tmpdir, f"part3_chk_{idx}.docx")
                doc_chk.save(chk_docx)
                subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", tmpdir, chk_docx], check=True)
                pdf_parts.append(os.path.join(tmpdir, f"part3_chk_{idx}.pdf"))

        merger = PdfWriter()
        for pdf in pdf_parts:
            if os.path.exists(pdf):
                merger.append(pdf)
        
        final_pdf_path = os.path.join(tmpdir, "final_report.pdf")
        merger.write(final_pdf_path)
        merger.close()

        with open(final_pdf_path, "rb") as f:
            return f.read()

st.title("🏗️ 공사감리일지 & 체크리스트 앱")

st.subheader("1. 감리 구분 및 공사 개요")
supervision_mode = st.radio("감리업무 구분 선택", ["비상주 감리", "책임상주 감리"], horizontal=True)

project_name = st.text_input("공사명", value="서탄산업단지 (주)노바테크놀러지스 신축공사")
col1, col2, col3 = st.columns(3)
with col1:
    report_date = st.date_input("작성일", value=date.today())
with col2:
    weather = st.selectbox("날씨", ["맑음", "흐림", "비", "눈"])
with col3:
    floor_info = st.text_input("시공 층/위치", value="지상층 / 부지내")

col_a, col_b = st.columns(2)
with col_a:
    inspector_name = st.text_input("건축사보(감리원)", value="홍사현")
with col_b:
    contractor_name = st.text_input("총괄 시공 책임자", value="이상민")

st.subheader(f"2. [{supervision_mode}] 공종별 체크리스트")
available_processes = list(CHECKLIST_DB.get(supervision_mode, {}).keys())
selected_processes = st.multiselect("금일 감리 공종 선택 (전체 공종 탑재됨)", available_processes)

all_checklist_contexts = []
diary_summary_items = []
diary_summary_details = []

for proc in selected_processes:
    with st.expander(f"📋 [{proc}] 체크리스트 점검표 ({len(CHECKLIST_DB[supervision_mode][proc])}개 항목)", expanded=True):
        filter_opt = st.radio("표시 범위", ["전체 보기", "기본업무만", "기본외업무만"], horizontal=True, key=f"flt_{proc}")
        
        rows_for_template = []
        for idx, item in enumerate(CHECKLIST_DB[supervision_mode][proc]):
            if filter_opt == "기본업무만" and item["업무구분"] != "기본업무":
                continue
            if filter_opt == "기본외업무만" and item["업무구분"] != "기본외업무":
                continue

            c1, c2, c3 = st.columns([4, 2, 2])
            with c1:
                st.write(f"**[{item['업무구분']}]** {item['항목']}")
            with c2:
                res = st.selectbox("결과", ["적합", "부적합", "해당없음"], key=f"{supervision_mode}_{proc}_{idx}", label_visibility="collapsed")
            with c3:
                action = st.text_input("조치사항", value="해당없음" if res == "해당없음" else "", key=f"act_{supervision_mode}_{proc}_{idx}", label_visibility="collapsed", placeholder="조치사항")
            
            rows_for_template.append({
                "category": item["업무구분"],
                "item": item["항목"],
                "pass": "✓" if res == "적합" else "",
                "fail": "✓" if res == "부적합" else "",
                "action": action
            })
            
            if res in ["적합", "부적합"]:
                diary_summary_items.append(f"[{proc}] {item['업무구분']}")
                diary_summary_details.append(f"• {item['항목']} ({res})")

        all_checklist_contexts.append({
            "mode": supervision_mode,
            "process": proc,
            "floor": floor_info,
            "year": report_date.year,
            "month": f"{report_date.month:02d}",
            "day": f"{report_date.day:02d}",
            "inspector": inspector_name,
            "contractor": contractor_name,
            "rows": rows_for_template
        })

st.subheader("3. 특기사항 및 지적사항")
special_note = st.text_area("특기사항", value="- 해당공정 인부 안전관리 철저 주문\n- 용접공정이나 기타 인화물질자재 사용시 화재 안전기준 철저 주문")
issue_note = st.text_area("지적사항 및 처리결과", placeholder="지적사항 및 시공사 조치 결과를 입력하세요.")

st.subheader("4. 현장 사진 첨부 (6장당 1페이지 자동 편집)")
cam_photo = st.camera_input("카메라로 바로 찍기")
multi_photos = st.file_uploader("또는 갤러리에서 여러 장 선택", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

photo_entries = []
combined_uploads = []
if cam_photo:
    combined_uploads.append(cam_photo)
if multi_photos:
    combined_uploads.extend(multi_photos)

for idx, file in enumerate(combined_uploads):
    col_img, col_txt = st.columns([1, 2])
    with col_img:
        st.image(file, use_container_width=True)
    with col_txt:
        cap = st.text_input(f"사진 {idx+1} 설명(캡션)", placeholder="예: 규준틀 및 터파기 확인(Y4,X1)", key=f"cap_{idx}")
    photo_entries.append((file.getvalue(), cap))

st.subheader("5. 날인 도장 업로드")
col_s1, col_s2 = st.columns(2)
with col_s1:
    inspector_stamp_file = st.file_uploader("건축사보(감리원) 도장 (PNG)", type=["png", "jpg"])
    if inspector_stamp_file:
        st.image(inspector_stamp_file, width=70)
with col_s2:
    contractor_stamp_file = st.file_uploader("총괄 시공 책임자 도장 (PNG)", type=["png", "jpg"])
    if contractor_stamp_file:
        st.image(contractor_stamp_file, width=70)

st.divider()
if st.button("📄 샘플 동일 양식 PDF 일괄 생성", use_container_width=True):
    weekday_str = WEEKDAYS[report_date.weekday()]
    
    diary_context = {
        "project_name": project_name,
        "year": report_date.year,
        "month": f"{report_date.month:02d}",
        "day": f"{report_date.day:02d}",
        "weekday": weekday_str,
        "weather": weather,
        "process_list": ", ".join(selected_processes),
        "floor": floor_info,
        "inspector": inspector_name,
        "audit_items": "\n".join(dict.fromkeys(diary_summary_items)),
        "audit_details": "\n".join(diary_summary_details),
        "special_note": special_note,
        "issue_note": issue_note
    }

    insp_bytes = inspector_stamp_file.getvalue() if inspector_stamp_file else None
    cont_bytes = contractor_stamp_file.getvalue() if contractor_stamp_file else None

    with st.spinner("감리일지 + 현장사진대지 + 체크리스트를 PDF로 병합 생성 중..."):
        try:
            final_pdf = build_combined_pdf(diary_context, all_checklist_contexts, photo_entries, insp_bytes, cont_bytes)
            st.success("✅ 통합 PDF 보고서가 완성되었습니다!")
            st.download_button(
                label="📥 공사감리일지 통합 PDF 다운로드",
                data=final_pdf,
                file_name=f"[별지2]공사감리일지({report_date}).pdf",
                mime="application/pdf",
                use_container_width=True
            )
        except Exception as e:
            st.error(f"생성 중 오류가 발생했습니다: {e}")