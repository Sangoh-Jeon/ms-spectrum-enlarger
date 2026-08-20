# 🧪 GLP MS/MS Spectrum Peak Enlarger (질량분석 피크 분자량 확대 프로그램)

[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://share.streamlit.io)
[![Python 3.9+](https://img.shields.io/badge/python-3.9+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

GLP(우수시험검사기준) 보고서 및 질량분석(LC-MS/MS, GC-MS/MS) 시험 자료 작성 시, 스펙트럼 이미지 내 작게 표시된 **전구이온(Precursor Ion) 및 조각이온(Product Ion)의 분자량($m/z$) 텍스트를 고화질로 확대/강조**해 주는 도구입니다.

---

## 📌 주요 기능 (Key Features)

1. **자동 피크 분자량($m/z$) 인식 및 선명한 확대**
   - 질량분석 스펙트럼 이미지에서 기존의 작고 흐릿한 $m/z$ 수치를 인식 후 **대형 폰트(기본 48pt)** 로 재배치
   - 타겟 MRM 이온은 **진한 파란색(Blue, 볼드)**, 기타 주요 이온은 **회색(Gray)** 으로 구분 강조
2. **이중 OCR 엔진 지원 (Dual OCR Engine)**
   - **Google Cloud Vision AI (Lens)**: 99.99% 초고정밀 인식
   - **Local EasyOCR / OpenCV Engine**: API 키 없이도 로컬 환경에서 즉시 구동 가능
3. **엑셀 원본 서식 완벽 유지 (Excel Report Compatible)**
   - 엑셀 내 삽입된 스펙트럼 이미지만을 자동으로 탐지하여 처리된 고화질 이미지로 1:1 교체
   - 기존 보고서 양식, 표, 텍스트 서식 완벽 보존
4. **웹 UI (Streamlit Web App) 및 로컬 GUI / EXE 지원**
   - 직관적인 웹 인터페이스(`app.py`)를 통해 파일 업로드/다운로드 가능
   - 독립 실행형 데스크톱 실행파일(`.exe`) 빌드 지원

---

## 🚀 빠른 시작 (Quick Start)

### 1. 저장소 복제 (Clone Repository)
```bash
git clone https://github.com/Sangoh-Jeon/ms-spectrum-enlarger.git
cd ms-spectrum-enlarger
```

### 2. 가상환경 생성 및 패키지 설치
```bash
python -m venv venv
# Windows
venv\Scripts\activate
# Mac/Linux
source venv/bin/activate

pip install -r requirements.txt
```

### 3. Streamlit 웹 애플리케이션 실행
```bash
streamlit run app.py
```
브라우저에서 `http://localhost:8501`로 자동 접속됩니다.

---

## ☁️ Streamlit Cloud 배포 가이드 (Deployment)

1. [Streamlit Community Cloud](https://share.streamlit.io/)에 접속하여 GitHub 계정으로 로그인합니다.
2. **"New app"** 클릭 후 다음을 설정합니다:
   - **Repository**: `Sangoh-Jeon/ms-spectrum-enlarger`
   - **Branch**: `main`
   - **Main file path**: `app.py`
3. *(선택사항)* Google Cloud Vision API를 사용하려면:
   - App Settings -> **Secrets**에 서비스 계정 키 JSON 등록:
     ```toml
     [gcp_service_account]
     type = "service_account"
     project_id = "your-project-id"
     private_key_id = "..."
     private_key = "..."
     client_email = "..."
     ...
     ```
4. **Deploy** 버튼을 클릭하면 수 분 내에 웹 서비스가 공개됩니다.

---

## 📂 파일 구조 (File Structure)

```text
├── app.py                     # Streamlit 기반 웹 애플리케이션
├── enlarge_spectrum_peaks.py  # 핵심 이미지/엑셀 처리 및 OCR 엔진 모듈
├── requirements.txt           # 의존성 패키지 목록
└── README.md                  # 프로젝트 설명서
```

---

## 🛠 기술 스택 (Tech Stack)

- **Language**: Python 3.9+
- **Frontend / Web**: Streamlit
- **Image Processing**: OpenCV, Pillow (PIL), NumPy
- **OCR Engine**: Google Cloud Vision API, EasyOCR
- **Spreadsheet**: openpyxl, pandas

---

## 📄 License
This project is licensed under the MIT License.
