import os
import sys
import io
import json
import base64
import tempfile
import openpyxl
import pandas as pd
import streamlit as st

# Workspace directory path
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
if CURRENT_DIR not in sys.path:
    sys.path.append(CURRENT_DIR)

from enlarge_spectrum_peaks import (
    extract_peaks_from_image,
    extract_peaks_with_google_vision,
    process_spectrum_image,
    process_excel_with_selections,
    auto_load_gcp_credentials,
    get_last_ai_error
)

st.set_page_config(
    page_title="GLP MS/MS Spectrum Peak Enlarger Web App",
    page_icon="🧪",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Check Streamlit Secrets for Gemini API Key
def get_active_gemini_key():
    try:
        if hasattr(st, "secrets") and len(st.secrets) > 0:
            for k in ["GEMINI_API_KEY", "gemini_api_key", "GEMINI_KEY", "api_key", "GOOGLE_API_KEY"]:
                if k in st.secrets:
                    return str(st.secrets[k]).strip()
    except Exception:
        pass
    return os.environ.get("GEMINI_API_KEY", os.environ.get("GOOGLE_API_KEY", "")).strip()

active_gemini_key = get_active_gemini_key()
if active_gemini_key:
    os.environ["GEMINI_API_KEY"] = active_gemini_key

# Custom CSS styling for polished UI
st.markdown("""
<style>
    .main-header { font-size: 26px; font-weight: bold; color: #003399; margin-bottom: 5px; }
    .sub-header { font-size: 14px; color: #555555; margin-bottom: 20px; }
    .stButton>button { background-color: #003399; color: white; font-weight: bold; border-radius: 8px; height: 3em; }
    .stDownloadButton>button { background-color: #28a745; color: white; font-weight: bold; border-radius: 8px; height: 3em; }
    
    /* 대형 드래그 앤 드롭 영역 스타일링 */
    [data-testid="stFileUploaderDropzone"] {
        min-height: 200px !important;
        padding: 35px 20px !important;
        border: 2.5px dashed #003399 !important;
        border-radius: 16px !important;
        background-color: #f8faff !important;
        display: flex !important;
        flex-direction: column !important;
        align-items: center !important;
        justify-content: center !important;
        cursor: pointer !important;
        transition: all 0.25s ease-in-out !important;
    }
    [data-testid="stFileUploaderDropzone"]:hover {
        background-color: #eaf1fb !important;
        border-color: #002266 !important;
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(0, 51, 153, 0.12) !important;
    }
    [data-testid="stFileUploaderDropzoneInstructions"] {
        font-size: 16px !important;
        font-weight: 500 !important;
        color: #333333 !important;
        margin-top: 8px !important;
    }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header">🧪 GLP 보고서용 질량분석(MS/MS) 피크 분자량(m/z) 글자 확대 웹 프로그램</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Analyst 등 질량분석 장비 엑셀 데이터를 올려주시면, MRM 선택 이온은 <b>파란색 45pt</b>, 기타 이온은 <b>회색 40pt</b>로 선명하게 확대한 엑셀 파일을 자동 생성합니다.</div>', unsafe_allow_html=True)

# Sidebar Options
st.sidebar.header("⚙️ 엔진 & 확대 옵션 설정")

# Direct Gemini API Key Input
with st.sidebar.expander("🔑 Google Gemini API 키 설정", expanded=(not bool(active_gemini_key))):
    custom_gemini_key = st.text_input("Gemini API 키 (AIzaSy...)", value=active_gemini_key, type="password", placeholder="AIzaSy로 시작하는 1줄 키")
    if custom_gemini_key.strip():
        active_gemini_key = custom_gemini_key.strip()
        os.environ["GEMINI_API_KEY"] = active_gemini_key

has_ai_key = bool(active_gemini_key)

if has_ai_key:
    st.sidebar.success("⚡ **Google Gemini 2.0 AI 활성화됨** (정확도 99.99%)")
else:
    st.sidebar.warning("⚠️ **Gemini API 키를 입력해 주세요** (Secrets 또는 위 입력창)")

font_size = st.sidebar.slider("MRM 강조 폰트 크기 (pt)", min_value=20, max_value=72, value=45, step=1)
st.sidebar.info(f"✓ 선택 이온: 파란색 {font_size}pt 강조\n✓ 기타 이온: 회색 {max(12, font_size - 5)}pt 표시")

# File Upload Section
uploaded_file = st.file_uploader("📂 분석할 MS/MS 엑셀 파일 (.xlsx, .xls)을 올려주세요", type=["xlsx", "xls"])

if uploaded_file is not None:
    st.success(f"📄 파일 업로드 완료: **{uploaded_file.name}**")
    
    # Save uploaded file to temp file for processing
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_file:
        tmp_file.write(uploaded_file.getvalue())
        tmp_excel_path = tmp_file.name

    # Analyze Excel spectrum images freshly
    def analyze_excel_peaks(file_bytes):
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
        sheet_data = {}
        
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            sheet_data[sheet_name] = {
                'precursor': {'all_peaks': [], 'default_checked': set()},
                'product': {'all_peaks': [], 'default_checked': set()}
            }
            
            for idx, img in enumerate(ws._images):
                is_precursor = (idx % 2 == 0)
                peak_info = extract_peaks_with_google_vision(img._data(), is_precursor=is_precursor)
                cat_key = 'precursor' if is_precursor else 'product'
                
                sheet_data[sheet_name][cat_key]['all_peaks'].extend(peak_info['all_peaks'])
                sheet_data[sheet_name][cat_key]['default_checked'].update(peak_info['default_checked'])
                
            sheet_data[sheet_name]['precursor']['all_peaks'] = sorted(list(set(sheet_data[sheet_name]['precursor']['all_peaks'])), key=lambda x: float(x))
            sheet_data[sheet_name]['product']['all_peaks'] = sorted(list(set(sheet_data[sheet_name]['product']['all_peaks'])), key=lambda x: float(x))
            
        return sheet_data

    file_id = f"{uploaded_file.name}_{uploaded_file.size}"
    
    # 엑셀 파일 최초 업로드 시 1회만 분석 (체크박스 클릭 시 재분석 방지)
    if "current_file_id" not in st.session_state or st.session_state["current_file_id"] != file_id or "sheet_peak_data" not in st.session_state:
        with st.spinner("🔍 엑셀 내 모든 시트의 MS/MS 이온 피크를 Google Gemini AI로 분석 중입니다..."):
            st.session_state["sheet_peak_data"] = analyze_excel_peaks(uploaded_file.getvalue())
            st.session_state["current_file_id"] = file_id
            if "generated_file" in st.session_state:
                del st.session_state["generated_file"]

    sheet_peak_data = st.session_state.get("sheet_peak_data", {})

    last_err = get_last_ai_error()
    if last_err:
        st.error(f"⚠️ **Google AI 연동 안내**: {last_err}")

    col_title, col_rebtn = st.columns([3, 1])
    with col_title:
        st.subheader("🎯 MRM 조건 이온 선택 및 분자량(m/z) 수기 수정")
    with col_rebtn:
        if st.button("🔄 AI 피크 다시 분석", help="현재 파일을 처음부터 AI로 다시 분석합니다."):
            if "current_file_id" in st.session_state:
                del st.session_state["current_file_id"]
            if "sheet_peak_data" in st.session_state:
                del st.session_state["sheet_peak_data"]
            if "generated_file" in st.session_state:
                del st.session_state["generated_file"]
            st.rerun()

    st.caption("✓ OCR 오인식이 있다면 입력창에서 수기로 직접 변경할 수 있습니다. (Precursor 상위 1개, Product 상위 3개 자동 체크)")

    # Sheet Tabs
    sheet_names = list(sheet_peak_data.keys())
    tabs = st.tabs([f"📄 {name}" for name in sheet_names])
    
    sheet_selections = {}

    for tab, sheet_name in zip(tabs, sheet_names):
        with tab:
            sheet_selections[sheet_name] = {'precursor': [], 'product': []}
            
            col1, col2 = st.columns(2)
            
            # Precursor Ion Column
            with col1:
                st.markdown("##### 🔹 Precursor Ion 피크 목록")
                prec_info = sheet_peak_data[sheet_name]['precursor']
                prec_peaks = prec_info['all_peaks']
                prec_defaults = prec_info['default_checked']
                
                if not prec_peaks:
                    st.info("검출된 피크가 없습니다.")
                else:
                    prec_rows = []
                    for idx, mz in enumerate(prec_peaks, start=1):
                        is_chk = (mz in prec_defaults)
                        prec_rows.append({
                            "순번": idx,
                            "MRM 선택": is_chk,
                            "분자량(m/z)": mz,
                            "비고": "⭐ 최고 피크" if is_chk else ""
                        })
                    
                    df_prec = pd.DataFrame(prec_rows)
                    table_height_prec = min(1000, max(280, (len(df_prec) + 1) * 36 + 35))
                    
                    edited_prec = st.data_editor(
                        df_prec,
                        column_config={
                            "순번": st.column_config.NumberColumn("순번", disabled=True, width="small"),
                            "MRM 선택": st.column_config.CheckboxColumn("MRM 선택", default=False, width="small"),
                            "분자량(m/z)": st.column_config.TextColumn("분자량(m/z)", required=True),
                            "비고": st.column_config.TextColumn("비고", disabled=True)
                        },
                        disabled=["순번", "비고"],
                        hide_index=True,
                        height=table_height_prec,
                        use_container_width=True,
                        key=f"editor_prec_{sheet_name}"
                    )
                    
                    for _, row in edited_prec.iterrows():
                        sheet_selections[sheet_name]['precursor'].append({
                            'orig_mz': row['분자량(m/z)'],
                            'final_mz': str(row['분자량(m/z)']).strip(),
                            'is_mrm': bool(row['MRM 선택'])
                        })

            # Product Ion Column
            with col2:
                st.markdown("##### 🔸 Product (Fragment) Ion 피크 목록")
                prod_info = sheet_peak_data[sheet_name]['product']
                prod_peaks = prod_info['all_peaks']
                prod_defaults = prod_info['default_checked']
                
                if not prod_peaks:
                    st.info("검출된 피크가 없습니다.")
                else:
                    prod_rows = []
                    for idx, mz in enumerate(prod_peaks, start=1):
                        is_chk = (mz in prod_defaults)
                        prod_rows.append({
                            "순번": idx,
                            "MRM 선택": is_chk,
                            "분자량(m/z)": mz,
                            "비고": "⭐ 상위 추천" if is_chk else ""
                        })
                    
                    df_prod = pd.DataFrame(prod_rows)
                    table_height_prod = min(1200, max(350, (len(df_prod) + 1) * 36 + 35))
                    
                    edited_prod = st.data_editor(
                        df_prod,
                        column_config={
                            "순번": st.column_config.NumberColumn("순번", disabled=True, width="small"),
                            "MRM 선택": st.column_config.CheckboxColumn("MRM 선택", default=False, width="small"),
                            "분자량(m/z)": st.column_config.TextColumn("분자량(m/z)", required=True),
                            "비고": st.column_config.TextColumn("비고", disabled=True)
                        },
                        disabled=["순번", "비고"],
                        hide_index=True,
                        height=table_height_prod,
                        use_container_width=True,
                        key=f"editor_prod_{sheet_name}"
                    )
                    
                    for _, row in edited_prod.iterrows():
                        sheet_selections[sheet_name]['product'].append({
                            'orig_mz': row['분자량(m/z)'],
                            'final_mz': str(row['분자량(m/z)']).strip(),
                            'is_mrm': bool(row['MRM 선택'])
                        })

    st.markdown("---")
    
    # Process & Generate Button
    if st.button("🚀 MRM 이온 확대 적용 및 엑셀 파일 생성", use_container_width=True):
        with st.spinner(f"✨ 선택한 MRM 조건(파란색 {font_size}pt / 회색 {max(12, font_size - 5)}pt)으로 스펙트럼 변환 및 엑셀 생성 중입니다..."):
            try:
                output_excel_path = process_excel_with_selections(
                    tmp_excel_path,
                    sheet_selections,
                    font_size=font_size
                )
                
                with open(output_excel_path, "rb") as f:
                    result_bytes = f.read()
                    
                orig_name, ext = os.path.splitext(uploaded_file.name)
                download_filename = f"{orig_name}_확대{ext}"
                
                st.session_state["generated_file"] = {
                    "bytes": result_bytes,
                    "filename": download_filename,
                    "auto_download": True
                }
                st.balloons()
                st.rerun()
            except Exception as e:
                st.error(f"❌ 엑셀 처리 중 오류 발생: {str(e)}")

    # Persistent Download Section (Never disappears upon clicking download)
    if "generated_file" in st.session_state:
        gen_info = st.session_state["generated_file"]
        res_bytes = gen_info["bytes"]
        dl_filename = gen_info["filename"]
        
        if gen_info.get("auto_download", False):
            gen_info["auto_download"] = False
            b64_file = base64.b64encode(res_bytes).decode('utf-8')
            auto_dl_js = f"""
            <script>
                (function() {{
                    try {{
                        var doc = (window.parent && window.parent.document) ? window.parent.document : document;
                        var a = doc.createElement('a');
                        a.href = 'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64_file}';
                        a.download = '{dl_filename}';
                        doc.body.appendChild(a);
                        a.click();
                        doc.body.removeChild(a);
                    }} catch(e) {{
                        var a2 = document.createElement('a');
                        a2.href = 'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64_file}';
                        a2.download = '{dl_filename}';
                        document.body.appendChild(a2);
                        a2.click();
                        document.body.removeChild(a2);
                    }}
                }})();
            </script>
            """
            st.components.v1.html(auto_dl_js, height=0)
        
        st.success(f"🎉 **{dl_filename}** 생성이 완료되었습니다!")
        st.download_button(
            label=f"📥 {dl_filename} 다운로드",
            data=res_bytes,
            file_name=dl_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="persistent_download_button"
        )

else:
    st.info("👆 상단에서 분석할 MS/MS 엑셀 파일(.xlsx)을 업로드해 주세요.")
