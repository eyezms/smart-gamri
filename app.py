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
st.title("🏗️ 공사감리일지 & 체크리스트 (v5.0)")

with open("checklist_db.json", "r", encoding="utf-8") as f:
    DB = json.load(f)

st.subheader("1. 감리 구분 및 공사 개요")
gubun = st.radio("감리업무 구분", ["비상주 감리", "책임상주 감리"], horizontal=True)
site_name = st.text_input("공사명", value="서탄산업단지 (주)노바테크놀러지스 신축공사")

c1, c2, c3, c4 = st.columns(4)
serial_no = c1.text_input("일련번호", value="", placeholder="예: 01")
work_date = c2.date_input("작성일", value=date.today())
weather = c3.selectbox("날씨", ["맑음", "구름많음", "흐림", "비", "눈"])
floor_txt = c4.text_input("층 / 부위", value="지상층")

c5, c6, c7, c8 = st.columns(4)
loc_txt = c5.text_input("위치", value="부지내")
chief_sup = c6.text_input("총괄감리책임자", value="", placeholder="성명 입력")
inspector = c7.text_input("건축사보(감리원)", value="", placeholder="성명 입력")
contractor = c8.text_input("총괄 시공 책임자", value="", placeholder="성명 입력")
site_mgr = st.text_input("공종별 시공 관리자 (선택)", value="", placeholder="성명 입력")

st.subheader(f"2. [{gubun}] 공종별 체크리스트 (PDF 출력 시 기본업무+기본외업무 전체 출력)")
view_filter = st.radio("화면 입력 필터 (PDF는 항상 전체 항목 출력)", ["기본업무만 보기", "전체 항목 보기 (기본+기본외)"], horizontal=True)
selected_procs = st.multiselect("금일 감리 공종 선택", list(DB[gubun].keys()))

full_check_data = {}
for proc in selected_procs:
    all_items = DB[gubun][proc]
    b_cnt = sum(1 for x in all_items if x["cat"] == "기본업무")
    e_cnt = sum(1 for x in all_items if x["cat"] != "기본업무")

    with st.expander(f"📋 {proc} (기본업무 {b_cnt}건 / 기본외업무 {e_cnt}건 → PDF 전체 {len(all_items)}건 출력)", expanded=True):
        full_check_data[proc] = []
        for idx, x in enumerate(all_items):
            is_basic = (x["cat"] == "기본업무")
            # 화면 필터링만 적용, 데이터는 전체 보존
            if view_filter == "기본업무만 보기" and not is_basic:
                res = st.session_state.get(f"chk_{proc}_{idx}", "미선택")
                note_val = ""
            else:
                col_t, col_r = st.columns([3, 1])
                badge = "🟢" if is_basic else "🔵"
                sub_tag = f" | {x['sub']}" if x.get("sub") else ""
                col_t.write(f"{badge} **[{x['cat']}{sub_tag}]** {x['item']}")
                default_idx = 1 if is_basic else 0  # 기본업무는 '적합', 기본외업무는 '미선택(공란)' 기본값
                res = col_r.selectbox(
                    "결과", ["미선택", "적합", "부적합", "해당없음"],
                    index=default_idx, key=f"chk_{proc}_{idx}", label_visibility="collapsed"
                )
                note_val = "해당없음" if res == "해당없음" else ""
            full_check_data[proc].append({
                "cat": x["cat"], "sub": x.get("sub", "-"), "item": x["item"],
                "res": res, "note": note_val
            })

st.subheader("3. 특기사항 및 지적사항")
special_note = st.text_area("특기사항", value="-해당공정 인부 안전관리 철저 주문\n-용접공정이나 기타 인화물질자재 사용시 화재 안전기준 철저 주문\n-건물인접지 낙하물이나 피해 방지 및 민원철저 대비 주문")
issue_note = st.text_area("지적사항 및 처리결과", value="")

st.subheader("4. 📸 현장사진 업로드 (2페이지부터 2열×3행 자동 배치)")
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
    c_ok.success(f"✅ 사진 {len(st.session_state['photos_bin'])}장 첨부 완료")
    if c_del.button("사진 초기화"):
        st.session_state["photos_bin"] = []
        st.rerun()
    cols = st.columns(2)
    for i, (fn, jb) in enumerate(st.session_state["photos_bin"]):
        with cols[i % 2]:
            st.image(jb, caption=f"사진 {i+1}", use_container_width=True)
            cap = st.text_input(f"사진 {i+1} 내용", value="현장사진", key=f"pcap_{i}")
            photo_data_list.append((jb, cap))

st.subheader("5. ✍️ 결재 도장/서명 첨부")
s1, s2 = st.columns(2)
stamp_chief = s1.file_uploader("총괄감리책임자 도장/서명", type=["png", "jpg", "jpeg"])
stamp_insp = s2.file_uploader("건축사보(감리원) 도장/서명", type=["png", "jpg", "jpeg"])

# --- 워드 표 XML 정밀 제어 유틸리티 ---
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

def set_h(row, h_mm, exact=False):
    th = OxmlElement('w:trHeight')
    th.set(qn('w:val'), str(int(h_mm * 56.7)))
    th.set(qn('w:hRule'), 'exact' if exact else 'atLeast')
    row._tr.get_or_add_trPr().append(th)

def style_cell(cell, txt="", sz=8.5, b=False, al=WD_ALIGN_PARAGRAPH.LEFT, bg=None, img=None, pad=(35, 35, 55, 55)):
    cell.text = ""
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    tp = cell._tc.get_or_add_tcPr()
    tm = OxmlElement('w:tcMar')
    for m, v in zip(['top', 'bottom', 'left', 'right'], pad):
        nd = OxmlElement(f'w:{m}')
        nd.set(qn('w:w'), str(v))
        nd.set(qn('w:type'), 'dxa')
        tm.append(nd)
    tp.append(tm)
    if bg: tp.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{bg}"/>'))
    p = cell.paragraphs[0]
    p.alignment = al
    p.paragraph_format.space_before = p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.1
    if txt:
        r = p.add_run(txt)
        r.font.name = "NanumGothic"
        r._element.rPr.rFonts.set(qn('w:eastAsia'), "NanumGothic")
        r.font.size, r.bold = Pt(sz), b
    if img and os.path.exists(img):
        p.add_run().add_picture(img, width=Mm(11))

def make_doc():
    d = Document()
    for s in d.sections:
        s.page_width, s.page_height = Mm(210), Mm(297)
        s.top_margin = s.bottom_margin = Mm(10)
        s.left_margin = s.right_margin = Mm(12)
    return d

def split_proc_label(proc_str):
    m = re.match(r'^(.+?)\s*\((.+?)\)$', proc_str)
    return (m.group(1).strip(), m.group(2).strip()) if m else (proc_str, "-")

# [1페이지] 별지 제2호서식 공사감리일지 (업로드 PDF 1p와 동일 양식)
def build_diary(tmp, d_str, sc, si):
    doc = make_doc()
    p0 = doc.add_paragraph("■ 건축공사 감리세부기준 [별지 제2호서식] <개정 2017. 2. 4.>")
    p0.paragraph_format.space_after = Pt(2)
    p0.runs[0].font.size = Pt(8)

    # 감리일지 본문 행 구성: 적합/부적합 선택된 항목만 공종-감리항목별로 그룹핑
    diary_rows = []
    for proc, items in full_check_data.items():
        main_p, _ = split_proc_label(proc)
        checked = [x for x in items if x["res"] in ["적합", "부적합"]]
        if not checked:
            continue
        sub_groups = {}
        for x in checked:
            sub_groups.setdefault(x["sub"], []).append(x["item"])
        first = True
        for sub_name, item_txts in sub_groups.items():
            diary_rows.append((main_p if first else "", sub_name, "\n".join(item_txts), first))
            first = False

    if not diary_rows:
        diary_rows.append(("-", "-", "-", True))

    # 6열 그리드: [일련번호/공종(34), 감리항목(30), 내용좌(56), 날씨/결재1(33), 결재2(33)] = 186mm
    total_rows = 6 + len(diary_rows)
    tbl = doc.add_table(rows=total_rows, cols=5)
    tbl.alignment, tbl.style = WD_TABLE_ALIGNMENT.CENTER, "Table Grid"
    fix_table(tbl, [34, 30, 56, 33, 33])

    set_h(tbl.rows[0], 7)
    set_h(tbl.rows[1], 14)
    set_h(tbl.rows[2], 9)
    set_h(tbl.rows[3], 9)

    # 상단 헤더: 일련번호 | 공사감리일지 | 총괄감리책임자 | 건축사보
    style_cell(tbl.cell(0, 0), "일련번호", 8.5, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(1, 0), serial_no, 9, False, WD_ALIGN_PARAGRAPH.CENTER)
    style_cell(tbl.cell(0, 1).merge(tbl.cell(1, 2)), "공 사 감 리 일 지", 18, True, WD_ALIGN_PARAGRAPH.CENTER)
    style_cell(tbl.cell(0, 3), "총괄감리책임자", 8.5, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(0, 4), "건축사보", 8.5, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(1, 3), f"{chief_sup} (서명 또는 인)", 8, False, WD_ALIGN_PARAGRAPH.CENTER, img=sc)
    style_cell(tbl.cell(1, 4), f"{inspector} (서명 또는 인)", 8, False, WD_ALIGN_PARAGRAPH.CENTER, img=si)

    # 공사명 | 작성일 및 날씨
    style_cell(tbl.cell(2, 0), "공 사 명", 9, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(2, 1).merge(tbl.cell(2, 2)), site_name, 9, True)
    style_cell(tbl.cell(2, 3).merge(tbl.cell(2, 4)), f"{d_str}  날씨 : {weather}", 8.5, False, WD_ALIGN_PARAGRAPH.CENTER)

    # 공종 및 세부공정 | 감리 항목 | 감리내용
    style_cell(tbl.cell(3, 0), f"공종 및 세부공정\n({floor_txt})", 8.5, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(3, 1), "감리 항목", 9, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(3, 2).merge(tbl.cell(3, 4)), "감 리 내 용", 9, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")

    r = 4
    merge_start = None
    for idx_r, (p_lbl, s_lbl, c_txt, is_first) in enumerate(diary_rows):
        set_h(tbl.rows[r], 16)
        if is_first:
            if merge_start is not None and r - 1 > merge_start:
                tbl.cell(merge_start, 0).merge(tbl.cell(r - 1, 0))
            merge_start = r
            style_cell(tbl.cell(r, 0), p_lbl, 9, True, WD_ALIGN_PARAGRAPH.CENTER)
        style_cell(tbl.cell(r, 1), s_lbl, 8.5, False, WD_ALIGN_PARAGRAPH.CENTER)
        style_cell(tbl.cell(r, 2).merge(tbl.cell(r, 4)), c_txt, 8.5)
        r += 1
    if merge_start is not None and r - 1 > merge_start:
        tbl.cell(merge_start, 0).merge(tbl.cell(r - 1, 0))

    set_h(tbl.rows[r], 22)
    style_cell(tbl.cell(r, 0), "특기사항", 9, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(r, 1).merge(tbl.cell(r, 4)), special_note, 8.5)

    set_h(tbl.rows[r + 1], 22)
    style_cell(tbl.cell(r + 1, 0), "지적사항 및\n처리결과", 8.5, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(r + 1, 1).merge(tbl.cell(r + 1, 4)), issue_note if issue_note.strip() else "\n", 8.5)

    pg = doc.add_paragraph(
        "작성방법\n"
        "1. 공종에는 주요공종 및 단위공종 그리고 해당 층수를 기재합니다.\n"
        "2. 감리항목은 공종별 감리 체크리스트를 기반으로 감리항목을 기재합니다.\n"
        "3. 감리내용에는 각 감리항목에 대하여 육안검사·입회·시험 등을 통한 감리내용을 구체적으로 기재하고, 그 결과와 업무수행 근거자료를 상세히 기재합니다.\n"
        "4. 특기사항은 특별히 명기되어 있지 아니한 내용의 발생·조치사항 등을 기재합니다.\n"
        "5. 지적사항 및 처리결과는 재시공 및 공사중지 등 구두 또는 서면에 의한 지시내용과 처리결과를 기재합니다."
    )
    pg.paragraph_format.space_before = Pt(3)
    pg.paragraph_format.space_after = Pt(1)
    pg.runs[0].font.size = Pt(7.5)

    pf = doc.add_paragraph("210mm×297mm[일반용지 60g/m²)]")
    pf.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    pf.runs[0].font.size = Pt(7)

    out = os.path.join(tmp, "diary.docx")
    doc.save(out)
    return out

# [2~3페이지] 현장사진 (업로드 PDF 2~3p와 동일: 2열 × 3행 = 페이지당 6장)
def build_photos(plist, tmp):
    doc = make_doc()
    pt = doc.add_paragraph("현 장 사 진")
    pt.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pt.paragraph_format.space_after = Pt(4)
    pt.runs[0].font.size, pt.runs[0].bold = Pt(14), True

    # 6장 단위로 페이지 분할 (2열 x 3세트(사진행+설명행))
    for page_start in range(0, len(plist), 6):
        if page_start > 0:
            doc.add_page_break()
        chunk = plist[page_start:page_start + 6]
        num_pairs = (len(chunk) + 1) // 2
        tbl = doc.add_table(rows=num_pairs * 2, cols=2)
        tbl.alignment, tbl.style = WD_TABLE_ALIGNMENT.CENTER, "Table Grid"
        fix_table(tbl, [93, 93])

        for pair_idx in range(num_pairs):
            img_row = tbl.rows[pair_idx * 2]
            cap_row = tbl.rows[pair_idx * 2 + 1]
            set_h(img_row, 68, exact=True)
            set_h(cap_row, 10, exact=True)
            for col_j in range(2):
                item_idx = pair_idx * 2 + col_j
                if item_idx < len(chunk):
                    cb, cap = chunk[item_idx]
                    pp = os.path.join(tmp, f"p_{page_start + item_idx}.jpg")
                    with open(pp, "wb") as fo: fo.write(cb)
                    style_cell(img_row.cells[col_j], "", 8, False, WD_ALIGN_PARAGRAPH.CENTER, img=None, pad=(20, 20, 20, 20))
                    p_img = img_row.cells[col_j].paragraphs[0]
                    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_img.add_run().add_picture(pp, width=Mm(84), height=Mm(62))
                    style_cell(cap_row.cells[col_j], cap, 8.5, False, WD_ALIGN_PARAGRAPH.CENTER, "F9F9F9", pad=(15, 15, 30, 30))
                else:
                    style_cell(img_row.cells[col_j], "")
                    style_cell(cap_row.cells[col_j], "")

    out = os.path.join(tmp, "photos.docx")
    doc.save(out)
    return out

# [4~5페이지] 공종별 감리 체크리스트 (업로드 PDF 4~5p와 100% 동일: 7열 표 + 기본/기본외 전체 출력)
def build_chk(tmp, idx, proc, items, s_date):
    doc = make_doc()
    main_p, sub_p = split_proc_label(proc)

    # 7열 그리드: [구분(14), 세부구분(20), 검사항목(76), 기준참고(22), 적합(11), 부적합(13), 조치사항(30)] = 186mm
    tbl = doc.add_table(rows=8 + len(items), cols=7)
    tbl.alignment, tbl.style = WD_TABLE_ALIGNMENT.CENTER, "Table Grid"
    fix_table(tbl, [14, 20, 76, 22, 11, 13, 30])

    set_h(tbl.rows[0], 11)
    set_h(tbl.rows[1], 8)
    set_h(tbl.rows[2], 8)
    set_h(tbl.rows[3], 6)
    set_h(tbl.rows[4], 6)

    # Row 0: 제목 | 문서번호
    style_cell(tbl.cell(0, 0).merge(tbl.cell(0, 5)), "공 종 별  감 리  체 크 리 스 트", 15, True, WD_ALIGN_PARAGRAPH.CENTER)
    style_cell(tbl.cell(0, 6), "문서번호\n", 8, False, WD_ALIGN_PARAGRAPH.LEFT)

    # Row 1: 공종 | [공종] | 세부공종 | [세부공종]
    style_cell(tbl.cell(1, 0).merge(tbl.cell(1, 1)), "공  종", 8.5, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(1, 2), main_p, 9, True)
    style_cell(tbl.cell(1, 3), "세부공종", 8.5, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(1, 4).merge(tbl.cell(1, 6)), sub_p, 8.5, False, WD_ALIGN_PARAGRAPH.CENTER)

    # Row 2: 부위 | 층 | [층] | 위치 | [위치]
    style_cell(tbl.cell(2, 0).merge(tbl.cell(2, 1)), "부  위", 8.5, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(2, 2), f"층 : {floor_txt}", 8.5)
    style_cell(tbl.cell(2, 3), "위  치", 8.5, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(2, 4).merge(tbl.cell(2, 6)), loc_txt, 8.5, False, WD_ALIGN_PARAGRAPH.CENTER)

    # Row 3~4: 2단 헤더
    style_cell(tbl.cell(3, 0).merge(tbl.cell(4, 0)), "구분", 8.5, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(3, 1).merge(tbl.cell(4, 1)), "세부구분", 8, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(3, 2).merge(tbl.cell(4, 2)), "검  사  항  목", 9, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(3, 3).merge(tbl.cell(4, 3)), "기준,참고사항", 7.5, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(3, 4).merge(tbl.cell(3, 5)), "검사결과", 8, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(4, 4), "적합", 7.5, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(4, 5), "부적합", 7.5, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(3, 6).merge(tbl.cell(4, 6)), "조치사항", 8.5, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")

    # 데이터 행 채우기 + 구분(col 0) & 세부구분(col 1) 세로 병합
    start_r = 5
    for i, x in enumerate(items):
        r = start_r + i
        style_cell(tbl.cell(r, 0), x["cat"], 7.5, True, WD_ALIGN_PARAGRAPH.CENTER, pad=(25, 25, 25, 25))
        style_cell(tbl.cell(r, 1), x["sub"], 7.5, False, WD_ALIGN_PARAGRAPH.CENTER, pad=(25, 25, 25, 25))
        style_cell(tbl.cell(r, 2), x["item"], 8, False, WD_ALIGN_PARAGRAPH.LEFT, pad=(25, 25, 45, 45))
        style_cell(tbl.cell(r, 3), "", 7.5, False, WD_ALIGN_PARAGRAPH.CENTER)
        style_cell(tbl.cell(r, 4), "✓" if x["res"] == "적합" else "", 9, True, WD_ALIGN_PARAGRAPH.CENTER)
        style_cell(tbl.cell(r, 5), "✓" if x["res"] == "부적합" else "", 9, True, WD_ALIGN_PARAGRAPH.CENTER)
        style_cell(tbl.cell(r, 6), x.get("note", ""), 7.5, False, WD_ALIGN_PARAGRAPH.CENTER)

    # 구분(col 0) 세로 병합
    c0_start = start_r
    for i in range(1, len(items) + 1):
        if i == len(items) or items[i]["cat"] != items[c0_start - start_r]["cat"]:
            if start_r + i - 1 > c0_start:
                tbl.cell(c0_start, 0).merge(tbl.cell(start_r + i - 1, 0))
            c0_start = start_r + i

    # 세부구분(col 1) 세로 병합
    c1_start = start_r
    for i in range(1, len(items) + 1):
        if i == len(items) or (items[i]["cat"], items[i]["sub"]) != (items[c1_start - start_r]["cat"], items[c1_start - start_r]["sub"]):
            if start_r + i - 1 > c1_start:
                tbl.cell(c1_start, 1).merge(tbl.cell(start_r + i - 1, 1))
            c1_start = start_r + i

    # 하단 서명란 (시공자점검일 / 감리자점검일 / 첨부자료)
    r = start_r + len(items)
    set_h(tbl.rows[r], 12)
    set_h(tbl.rows[r + 1], 12)
    set_h(tbl.rows[r + 2], 7)

    style_cell(tbl.cell(r, 0).merge(tbl.cell(r, 1)), "시공자점검일", 8, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(r, 2), s_date, 8.5, False, WD_ALIGN_PARAGRAPH.CENTER)
    style_cell(tbl.cell(r, 3).merge(tbl.cell(r, 6)), f"총괄 시공 책임자(또는 현장관리인) : {contractor} (인)\n공종별 시공 관리자 : {site_mgr} (인)", 8)

    style_cell(tbl.cell(r + 1, 0).merge(tbl.cell(r + 1, 1)), "감리자점검일", 8, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(r + 1, 2), s_date, 8.5, False, WD_ALIGN_PARAGRAPH.CENTER)
    style_cell(tbl.cell(r + 1, 3).merge(tbl.cell(r + 1, 6)), f"총괄 감리 책임자 : {chief_sup} (인)\n건축사보(공종별 감리 책임자) : {inspector} (인)", 8)

    style_cell(tbl.cell(r + 2, 0).merge(tbl.cell(r + 2, 1)), "첨부자료", 8, True, WD_ALIGN_PARAGRAPH.CENTER, "F2F2F2")
    style_cell(tbl.cell(r + 2, 2).merge(tbl.cell(r + 2, 6)), "", 8)

    out = os.path.join(tmp, f"chk_{idx}.docx")
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
                sc, si = None, None
                if stamp_chief:
                    sc = os.path.join(tmp, "sc.png")
                    Image.open(io.BytesIO(stamp_chief.getvalue())).save(sc)
                if stamp_insp:
                    si = os.path.join(tmp, "si.png")
                    Image.open(io.BytesIO(stamp_insp.getvalue())).save(si)

                # 업로드 PDF와 동일한 페이지 순서: 1.감리일지 -> 2.현장사진 -> 3.공종별 체크리스트
                docs = [build_diary(tmp, d_str, sc, si)]
                if photo_data_list:
                    docs.append(build_photos(photo_data_list, tmp))
                for idx, (proc, items) in enumerate(full_check_data.items()):
                    if items:
                        docs.append(build_chk(tmp, idx, proc, items, s_date))

                writer = PdfWriter()
                for dp in docs:
                    subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", tmp, dp], check=True)
                    writer.append(os.path.splitext(dp)[0] + ".pdf")
                f_pdf = os.path.join(tmp, "final.pdf")
                writer.write(f_pdf)
                with open(f_pdf, "rb") as fp:
                    st.download_button("⬇️ 완성된 통합 PDF 다운로드", fp.read(), file_name=f"[별지 2] 공사감리일지({work_date}).pdf", mime="application/pdf", use_container_width=True)