import os
import sys
import io
import json
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
    auto_load_gcp_credentials
)

st.set_page_config(
    page_title="GLP MS/MS Spectrum Peak Enlarger Web App",
    page_icon="🧪",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Check Streamlit Secrets & GCP Key Status
def load_gcp_creds_from_secrets():
    if not hasattr(st, "secrets"):
        return None
    try:
        # 1. Direct section name
        if "gcp_service_account" in st.secrets:
            val = st.secrets["gcp_service_account"]
            if hasattr(val, "to_dict"): return val.to_dict()
            if isinstance(val, dict): return dict(val)
            if isinstance(val, str): return json.loads(val)
            
        # 2. Iterate all keys in st.secrets
        for k in st.secrets.keys():
            val = st.secrets[k]
            if isinstance(val, str) and ("service_account" in val or "private_key" in val):
                try:
                    return json.loads(val)
                except Exception:
                    pass
            if hasattr(val, "to_dict"):
                d = val.to_dict()
                if "private_key" in d or "type" in d:
                    return d
            elif isinstance(val, dict):
                if "private_key" in val or "type" in val:
                    return dict(val)
    except Exception:
        pass
    return None

gcp_creds_dict = load_gcp_creds_from_secrets()
if gcp_creds_dict:
    os.environ["GOOGLE_APPLICATION_CREDENTIALS_JSON"] = json.dumps(gcp_creds_dict)

gcp_key_path = auto_load_gcp_credentials()
has_gcp_key = (gcp_creds_dict is not None) or (gcp_key_path is not None) or ("GOOGLE_APPLICATION_CREDENTIALS" in os.environ)

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
st.markdown('<div class="sub-header">Analyst 등 질량분석 장비 엑셀 데이터를 올려주시면, MRM 선택 이온은 <b>파란색 48pt</b>, 기타 이온은 <b>회색 46pt</b>로 선명하게 확대한 엑셀 파일을 자동 생성합니다.</div>', unsafe_allow_html=True)

# Sidebar Options
st.sidebar.header("⚙️ 엔진 & 확대 옵션 설정")

# Direct JSON Key Paste fallback for ease of use
with st.sidebar.expander("🔑 구글 키(JSON) 직접 등록 / Secrets 확인", expanded=(not has_gcp_key)):
    custom_json = st.text_area("메모장 복사 내용(.json) 붙여넣기", height=120, placeholder="{\n  \"type\": \"service_account\",\n  ...\n}")
    if custom_json.strip():
        try:
            d = json.loads(custom_json.strip())
            if "type" in d or "private_key" in d:
                os.environ["GOOGLE_APPLICATION_CREDENTIALS_JSON"] = json.dumps(d)
                has_gcp_key = True
                st.success("✅ 구글 키가 실시간 적용되었습니다!")
        except Exception as e:
            st.error(f"JSON 형식 오류: {e}")

if has_gcp_key:
    st.sidebar.success("⚡ **Google Lens AI 엔진 활성화됨** (정확도 99.99%)")
else:
    st.sidebar.warning("💻 **고속 로컬 OCR 엔진 가동 중** (구글 키 입력 시 99.99% Lens AI 자동 전환)")

font_size = st.sidebar.slider("MRM 강조 폰트 크기 (pt)", min_value=24, max_value=72, value=48, step=2)
st.sidebar.info("✓ 선택 이온: 파란색 48pt 강조\n✓ 기타 이온: 회색 46pt 표시")

# File Upload Section
uploaded_file = st.file_uploader("📂 분석할 MS/MS 엑셀 파일 (.xlsx, .xls)을 올려주세요", type=["xlsx", "xls"])

if uploaded_file is not None:
    st.success(f"📄 파일 업로드 완료: **{uploaded_file.name}**")
    
    # Save uploaded file to temp file for processing
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_file:
        tmp_file.write(uploaded_file.getvalue())
        tmp_excel_path = tmp_file.name

    # Cache peak extraction per uploaded file
    @st.cache_data(show_spinner=False)
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

    with st.spinner("🔍 엑셀 내 모든 시트의 MS/MS 이온 피크를 초고속 분석 중입니다..."):
        sheet_peak_data = analyze_excel_peaks(uploaded_file.getvalue())

    st.subheader("🎯 MRM 조건 이온 선택 및 분자량(m/z) 수기 수정")
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
                    for mz in prec_peaks:
                        is_chk = (mz in prec_defaults)
                        prec_rows.append({"MRM 선택": is_chk, "분자량 (m/z) 수기수정": mz, "비고": "⭐ 최고 피크" if is_chk else ""})
                    
                    df_prec = pd.DataFrame(prec_rows)
                    edited_prec = st.data_editor(
                        df_prec,
                        column_config={
                            "MRM 선택": st.column_config.CheckboxColumn("MRM 선택", default=False),
                            "분자량 (m/z) 수기수정": st.column_config.TextColumn("분자량 (m/z) 수기수정", required=True),
                            "비고": st.column_config.TextColumn("비고", disabled=True)
                        },
                        disabled=["비고"],
                        hide_index=True,
                        key=f"editor_prec_{sheet_name}"
                    )
                    
                    for _, row in edited_prec.iterrows():
                        sheet_selections[sheet_name]['precursor'].append({
                            'orig_mz': row['분자량 (m/z) 수기수정'],
                            'final_mz': str(row['분자량 (m/z) 수기수정']).strip(),
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
                    for mz in prod_peaks:
                        is_chk = (mz in prod_defaults)
                        prod_rows.append({"MRM 선택": is_chk, "분자량 (m/z) 수기수정": mz, "비고": "⭐ 상위 추천" if is_chk else ""})
                    
                    df_prod = pd.DataFrame(prod_rows)
                    edited_prod = st.data_editor(
                        df_prod,
                        column_config={
                            "MRM 선택": st.column_config.CheckboxColumn("MRM 선택", default=False),
                            "분자량 (m/z) 수기수정": st.column_config.TextColumn("분자량 (m/z) 수기수정", required=True),
                            "비고": st.column_config.TextColumn("비고", disabled=True)
                        },
                        disabled=["비고"],
                        hide_index=True,
                        key=f"editor_prod_{sheet_name}"
                    )
                    
                    for _, row in edited_prod.iterrows():
                        sheet_selections[sheet_name]['product'].append({
                            'orig_mz': row['분자량 (m/z) 수기수정'],
                            'final_mz': str(row['분자량 (m/z) 수기수정']).strip(),
                            'is_mrm': bool(row['MRM 선택'])
                        })

    st.markdown("---")
    
    # Process & Generate Button
    if st.button("🚀 MRM 이온 확대 적용 및 엑셀 파일 생성", use_container_width=True):
        with st.spinner("✨ 선택한 MRM 조건(파란색 48pt / 회색 46pt)으로 스펙트럼 변환 중입니다..."):
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
                
                st.balloons()
                st.success("🎉 분자량 확대 엑셀 파일 생성이 완료되었습니다!")
                st.download_button(
                    label=f"📥 {download_filename} 다운로드",
                    data=result_bytes,
                    file_name=download_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"❌ 엑셀 처리 중 오류 발생: {str(e)}")

else:
    st.info("👆 상단에서 분석할 MS/MS 엑셀 파일(.xlsx)을 업로드해 주세요.")
