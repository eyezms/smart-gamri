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
st.title("🏗️ 공사감리일지 & 체크리스트 (v3.7)")

with open("checklist_db.json", "r", encoding="utf-8") as f:
    DB = json.load(f)

def clean_txt(t):
    return re.sub(r'^[☑☐■□✓✔•\-\s]+', '', str(t)).strip()

def is_extra(c):
    s = str(c).replace(" ", "")
    return ("외" in s) or ("추가" in s)

# 법정 서식 [별표 1] 비상주 감리 '기본업무' 핵심 항목 판별 키워드
BASIC_KW = [
    "고저차", "벤치 마크", "벤치마크", "bm", "먹매김", "지내력", "지지층",
    "ks", "자재성능", "최종 마감", "최종마감", "최종 설치", "동결선", "터파기 깊이",
    "토질상태", "말뚝박기 관련 도서", "말뚝 시공완료", "바닥면의 레벨", "지정공사의 확인",
    "철근배근의 확인", "레디믹스트", "레미콘", "강재 규격", "설계도서와의 대조",
    "제작도면", "현치도", "앵커 볼트", "무수축 모르타르", "제조업자의 관련 자료",
    "강재의 재질 규격", "강재시험성적표", "방화셔터", "방화문", "개폐상태", "개폐장치",
    "미끄럼방지", "미끄럼 방지", "누수검사", "난간 등의 높이", "배연창 및 피난창",
    "고효율 에너지", "전기안전검사필증", "통신관련 검사필증", "소방관련 검사필증"
]

def classify_item(gubun_name, proc_name, idx, raw_cat, text_str):
    if is_extra(raw_cat):
        return "기본외업무"
    if "상주" in gubun_name:
        return "기본업무"
    # 비상주 가설공사: 앞 7건 기본업무, 이후 기본외업무
    if "가설" in proc_name:
        return "기본업무" if idx < 7 else "기본외업무"
    # 기타 비상주 공종: 법정 기준 핵심 키워드 포함 시 기본업무, 그 외 기본외업무
    return "기본업무" if any(kw in text_str.lower() for kw in BASIC_KW) else "기본외업무"

def extract_items(gubun_name, proc_name, proc_data):
    raw_list = []
    if isinstance(proc_data, dict):
        for k, v in proc_data.items():
            if isinstance(v, list):
                for item in v:
                    txt = item.get("item") or item.get("text") or item.get("검사항목") or str(item) if isinstance(item, dict) else str(item)
                    raw_list.append((str(k).strip(), clean_txt(txt)))
    elif isinstance(proc_data, list):
        for item in proc_data:
            if isinstance(item, dict):
                cat = item.get("cat") or item.get("category") or item.get("구분") or "기본업무"
                txt = item.get("item") or item.get("text") or item.get("검사항목") or item.get("content") or ""
                raw_list.append((str(cat).strip(), clean_txt(txt)))
            else:
                raw_list.append(("기본업무", clean_txt(item)))
    return [(classify_item(gubun_name, proc_name, idx, rc, rt), rt) for idx, (rc, rt) in enumerate(raw_list)]

# 1. 감리 구분 및 공사 개요
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

# 2. 공종별 체크리스트 (기본업무 / 기본외업무 구분)
st.subheader(f"2. [{gubun}] 공종별 체크리스트")
task_mode = st.radio("업무 구분 선택", ["기본업무", "기본외업무"], horizontal=True)
selected_procs = st.multiselect("금일 감리 공종 선택", list(DB[gubun].keys()))

check_results = {}
for proc in selected_procs:
    all_items = extract_items(gubun, proc, DB[gubun][proc])
    b_items = [(c, t) for c, t in all_items if not is_extra(c)]
    e_items = [(c, t) for c, t in all_items if is_extra(c)]
    show_items = b_items if task_mode == "기본업무" else all_items

    with st.expander(f"📋 {proc} (기본업무 {len(b_items)}건 / 기본외업무 {len(e_items)}건) → 출력 {len(show_items)}건", expanded=True):
        check_results[proc] = []
        for idx, (c_str, i_str) in enumerate(show_items):
            col_t, col_r = st.columns([3, 1])
            col_t.write(f"{'🔵' if is_extra(c_str) else '🟢'} **[{c_str}]** {i_str}")
            res = col_r.selectbox("결과", ["적합", "부적합", "해당없음"], key=f"{proc}_{task_mode}_{idx}", label_visibility="collapsed")
            check_results[proc].append({"cat": c_str, "item": i_str, "res": res})

# 3. 특기사항 및 지적사항
st.subheader("3. 특기사항 및 지적사항")
special_note = st.text_area("특기사항", value="- 해당공정 인부 안전관리 철저 주문\n- 용접공정이나 기타 인화물질자재 사용시 화재 안전기준 철저 주문")
issue_note = st.text_area("지적사항 및 처리결과", value="")

# 4. 현장 사진 업로드 (원래 위치로 복원)
st.subheader("4. 📸 현장 사진 업로드 (사진대지 자동 병합)")
if "photos_bin" not in st.session_state:
    st.session_state["photos_bin"] = []

up_files = st.file_uploader("앨범에서 현장 사진 선택 (여러 장 가능)", accept_multiple_files=True, key="album_up")
if up_files:
    temp_imgs = []
    for f_obj in up_files:
        try:
            im = ImageOps.exif_transpose(Image.open(io.BytesIO(f_obj.getvalue()))).convert("RGB")
            im.thumbnail((900, 900))
            out_b = io.BytesIO()
            im.save(out_b, format="JPEG", quality=82)
            temp_imgs.append((f_obj.name, out_b.getvalue()))
        except Exception as err:
            st.warning(f"사진 변환 실패({f_obj.name}): {err}")
    if temp_imgs:
        st.session_state["photos_bin"] = temp_imgs

photo_data_list = []
if st.session_state["photos_bin"]:
    c_ok, c_del = st.columns([3, 1])
    c_ok.success(f"✅ 사진 {len(st.session_state['photos_bin'])}장 첨부 완료 (PDF 맨 뒤에 병합됩니다)")
    if c_del.button("사진 초기화"):
        st.session_state["photos_bin"] = []
        st.rerun()
    cols = st.columns(2)
    for i, (fn, jb) in enumerate(st.session_state["photos_bin"]):
        with cols[i % 2]:
            st.image(jb, caption=f"사진 {i+1}", use_container_width=True)
            cap = st.text_input(f"사진 {i+1} 설명", value=f"시공 현황 ({location})", key=f"pcap_{i}")
            photo_data_list.append((jb, cap))

# 5. 결재 도장/서명 첨부
st.subheader("5. ✍️ 결재 도장/서명 첨부")
s1, s2 = st.columns(2)
stamp_insp = s1.file_uploader("건축사보(감리원) 도장/서명", type=["png", "jpg", "jpeg"])
stamp_chief = s2.file_uploader("총괄감리책임자 도장/서명", type=["png", "jpg", "jpeg"])

def fix_table(tbl, widths):
    tbl.autofit = False
    tl = OxmlElement('w:tblLayout')
    tl.set(qn('w:type'), 'fixed')
    tbl._tbl.tblPr.append(tl)
    tg = tbl._tbl.tblGrid
    for ch in list(tg): tg.remove(ch)
    for w in widths:
        gc = OxmlElement('w:gridCol')
        gc.set(qn('w:w'), str(int(w * 56.7)))
        tg.append(gc)
    for r in tbl.rows:
        for i, w in enumerate(widths):
            c = r.cells[i]
            c.width = Mm(w)
            tp = c._tc.get_or_add_tcPr()
            for old in tp.findall(qn('w:tcW')): tp.remove(old)
            tw = OxmlElement('w:tcW')
            tw.set(qn('w:w'), str(int(w * 56.7)))
            tw.set(qn('w:type'), 'dxa')
            tp.append(tw)

def set_h(row, h_mm):
    th = OxmlElement('w:trHeight')
    th.set(qn('w:val'), str(int(h_mm * 56.7)))
    th.set(qn('w:hRule'), 'atLeast')
    row._tr.get_or_add_trPr().append(th)

def style_cell(cell, txt="", sz=9, b=False, al=WD_ALIGN_PARAGRAPH.LEFT, bg=None, img=None):
    cell.text = ""
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    tp = cell._tc.get_or_add_tcPr()
    tm = OxmlElement('w:tcMar')
    for m, v in [('top', 40), ('bottom', 40), ('left', 65), ('right', 65)]:
        nd = OxmlElement(f'w:{m}')
        nd.set(qn('w:w'), str(v))
        nd.set(qn('w:type'), 'dxa')
        tm.append(nd)
    tp.append(tm)
    if bg: tp.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{bg}"/>'))
    p = cell.paragraphs[0]
    p.alignment = al
    p.paragraph_format.space_before = p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.12
    if txt:
        r = p.add_run(txt)
        r.font.name = "NanumGothic"
        r._element.rPr.rFonts.set(qn('w:eastAsia'), "NanumGothic")
        r.font.size, r.bold = Pt(sz), b
    if img and os.path.exists(img):
        p.add_run().add_picture(img, width=Mm(12))

def make_doc():
    d = Document()
    for s in d.sections:
        s.page_width, s.page_height = Mm(210), Mm(297)
        s.top_margin = s.bottom_margin = Mm(10)
        s.left_margin = s.right_margin = Mm(12)
    return d

def build_diary(tmp, d_str, si, sc):
    doc = make_doc()
    p0 = doc.add_paragraph("■ 건축공사 감리세부기준 [별지 제2호서식] <개정 2017. 2. 4.>")
    p0.paragraph_format.space_after = Pt(3)
    p0.runs[0].font.size = Pt(8)
    tbl = doc.add_table(rows=7 + len(check_results), cols=5)
    tbl.alignment, tbl.style = WD_TABLE_ALIGNMENT.CENTER, "Table Grid"
    fix_table(tbl, [34, 72, 22, 29, 29])
    for idx, h in enumerate([7, 14, 9, 9, 10]): set_h(tbl.rows[idx], h)
    style_cell(tbl.cell(0, 0).merge(tbl.cell(1, 2)), "공  사  감  리  일  지", 18, True, WD_ALIGN_PARAGRAPH.CENTER)
    style_cell(tbl.cell(0, 3), "건축사보", 8.5, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(0, 4), "총괄감리책임자", 8.5, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(1, 3), f"{inspector}\n(서명 또는 인)", 8.5, False, WD_ALIGN_PARAGRAPH.CENTER, img=si)
    style_cell(tbl.cell(1, 4), f"{chief_supervisor}\n(서명 또는 인)", 8.5, False, WD_ALIGN_PARAGRAPH.CENTER, img=sc)
    style_cell(tbl.cell(2, 0), "공  사  명", 9.5, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(2, 1).merge(tbl.cell(2, 4)), site_name, 9.5, True)
    style_cell(tbl.cell(3, 0), "작  성  일", 9.5, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(3, 1).merge(tbl.cell(3, 2)), d_str, 9)
    style_cell(tbl.cell(3, 3), "날  씨", 9.5, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(3, 4), weather, 9, False, WD_ALIGN_PARAGRAPH.CENTER)
    style_cell(tbl.cell(4, 0), f"공종 및 세부공정\n({location})", 8.5, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(4, 1).merge(tbl.cell(4, 4)), "감 리 내 용 (감리항목 및 확인결과)", 9.5, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    r = 5
    for proc, items in check_results.items():
        lines = [f"• [{x['cat']}] {x['item']} ({x['res']})" for x in items if x['res'] != "해당없음"]
        set_h(tbl.rows[r], 45)
        style_cell(tbl.cell(r, 0), proc, 9.5, True, WD_ALIGN_PARAGRAPH.CENTER)
        style_cell(tbl.cell(r, 1).merge(tbl.cell(r, 4)), "\n".join(lines), 8.5)
        r += 1
    set_h(tbl.rows[r], 22)
    style_cell(tbl.cell(r, 0), "특 기 사 항", 9.5, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(r, 1).merge(tbl.cell(r, 4)), special_note, 8.5)
    set_h(tbl.rows[r + 1], 24)
    style_cell(tbl.cell(r + 1, 0), "지적사항 및\n처리결과", 9, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(r + 1, 1).merge(tbl.cell(r + 1, 4)), issue_note if issue_note.strip() else "\n\n", 8.5)
    pg = doc.add_paragraph("작성방법: 1.주요/단위공종 및 층수 기재  2.체크리스트 기반 감리항목 기재  3.육안·입회·시험 결과 상세 기재  4.특기사항  5.지적사항 및 처리결과")
    pg.paragraph_format.space_before = Pt(4)
    pg.runs[0].font.size = Pt(7.5)
    out = os.path.join(tmp, "diary.docx")
    doc.save(out)
    return out

def build_chk(tmp, idx, proc, items, s_date):
    doc = make_doc()
    p0 = doc.add_paragraph(f"■ 건축공사 감리세부기준 [별표 1] ({gubun} - {task_mode})")
    p0.paragraph_format.space_after = Pt(3)
    p0.runs[0].font.size = Pt(8.5)
    tbl = doc.add_table(rows=6 + len(items), cols=5)
    tbl.alignment, tbl.style = WD_TABLE_ALIGNMENT.CENTER, "Table Grid"
    fix_table(tbl, [24, 98, 16, 18, 30])
    for i, h in enumerate([11, 9, 6, 6]): set_h(tbl.rows[i], h)
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
    set_h(tbl.rows[r], 9); set_h(tbl.rows[r + 1], 9)
    style_cell(tbl.cell(r, 0), "시공자점검일", 8.5, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(r, 1), s_date, 9, False, WD_ALIGN_PARAGRAPH.CENTER)
    style_cell(tbl.cell(r, 2).merge(tbl.cell(r, 4)), f"총괄 시공 책임자 : {contractor}  (서명 또는 인)", 9)
    style_cell(tbl.cell(r + 1, 0), "감리자점검일", 8.5, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(r + 1, 1), s_date, 9, False, WD_ALIGN_PARAGRAPH.CENTER)
    style_cell(tbl.cell(r + 1, 2).merge(tbl.cell(r + 1, 4)), f"건축사보(감리원) : {inspector}  (서명 또는 인)", 9)
    out = os.path.join(tmp, f"chk_{idx}.docx")
    doc.save(out)
    return out

def build_photos(plist, tmp):
    doc = make_doc()
    pt = doc.add_paragraph("■ 현장 감리 사진대지")
    pt.runs[0].font.size, pt.runs[0].bold = Pt(12), True
    tbl = doc.add_table(rows=(len(plist) + 1) // 2, cols=2)
    tbl.alignment, tbl.style = WD_TABLE_ALIGNMENT.CENTER, "Table Grid"
    fix_table(tbl, [93, 93])
    for i in range(0, len(plist), 2):
        row = tbl.rows[i // 2].cells
        for j in range(2):
            if i + j < len(plist):
                cb, cap = plist[i + j]
                pp = os.path.join(tmp, f"p_{i+j}.jpg")
                with open(pp, "wb") as fo: fo.write(cb)
                pi = row[j].paragraphs[0]
                pi.alignment = WD_ALIGN_PARAGRAPH.CENTER
                pi.add_run().add_picture(pp, width=Mm(82))
                pc = row[j].add_paragraph(f"[사진 {i+j+1}] {cap}")
                pc.alignment = WD_ALIGN_PARAGRAPH.CENTER
                pc.runs[0].font.size = Pt(8.5)
    out = os.path.join(tmp, "photos.docx")
    doc.save(out)
    return out

if st.button("📄 통합 감리일지 PDF 생성", type="primary", use_container_width=True):
    if not selected_procs:
        st.warning("최소 1개 이상의 공종을 선택해 주세요.")
    else:
        with st.spinner("통합 PDF 생성 중..."):
            with tempfile.TemporaryDirectory() as tmp:
                wk = ["월", "화", "수", "목", "금", "토", "일"][work_date.weekday()]
                d_str = f"{work_date.year}년 {work_date.month:02d}월 {work_date.day:02d}일 ({wk}요일)"
                s_date = f"{work_date.year}년 {work_date.month:02d}월 {work_date.day:02d}일"
                si, sc = None, None
                if stamp_insp:
                    si = os.path.join(tmp, "si.png")
                    Image.open(io.BytesIO(stamp_insp.getvalue())).save(si)
                if stamp_chief:
                    sc = os.path.join(tmp, "sc.png")
                    Image.open(io.BytesIO(stamp_chief.getvalue())).save(sc)
                docs = [build_diary(tmp, d_str, si, sc)]
                for idx, (proc, items) in enumerate(check_results.items()):
                    if items: docs.append(build_chk(tmp, idx, proc, items, s_date))
                if photo_data_list:
                    docs.append(build_photos(photo_data_list, tmp))
                writer = PdfWriter()
                for dp in docs:
                    subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", tmp, dp], check=True)
                    writer.append(os.path.splitext(dp)[0] + ".pdf")
                f_pdf = os.path.join(tmp, "final.pdf")
                writer.write(f_pdf)
                with open(f_pdf, "rb") as fp:
                    st.download_button("⬇️ 완성된 통합 PDF 다운로드", fp.read(), file_name=f"[별지2]공사감리일지({work_date}).pdf", mime="application/pdf", use_container_width=True)