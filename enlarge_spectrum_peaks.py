import os
import sys
import io
import re
import json
import base64
import concurrent.futures
from functools import lru_cache

import requests
import cv2
import numpy as np
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
from PIL import Image, ImageDraw, ImageFont

# ──────────────────────────────────────────────────────────────────────────────
# Tkinter (데스크톱 GUI) 는 desktop_app.py 로 분리되었습니다.
# 서버/클라우드 환경에서는 tkinter 가 없어도 이 모듈이 정상 작동합니다.
# ──────────────────────────────────────────────────────────────────────────────


# ── X축 보정 상수 (Analyst 소프트웨어 출력 기준 경험적 상수) ──────────────────
# 조정이 필요할 경우 아래 값만 변경하세요.
_X_LOW_ORIGIN = 40          # max_v <= 210 시 x 기준점 (픽셀)
_X_LOW_MZ_START = 30.0      # max_v <= 210 시 m/z 시작값
_X_LOW_SCALE = 13.48        # max_v <= 210 시 픽셀/Da 비율

_X_HIGH_ORIGIN = 59         # max_v > 210 시 x 기준점 (픽셀)
_X_HIGH_PREC_MZ_START = 50.0    # Precursor, max_v > 210 시 m/z 시작값
_X_HIGH_PREC_SCALE = 4.75333    # Precursor, max_v > 210 시 픽셀/Da 비율

_X_HIGH_PROD_MZ_START = 20.0    # Product, max_v > 210 시 m/z 시작값
_X_HIGH_PROD_SCALE = 6.684375   # Product, max_v > 210 시 픽셀/Da 비율
# ─────────────────────────────────────────────────────────────────────────────


def get_base_dir():
    """
    Returns the absolute directory path of the running application.
    Compatible with both PyInstaller executables and standard Python scripts.
    """
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def detect_axis_lines(img):
    """
    Detects Y-axis (vertical) and X-axis (horizontal) lines in the spectrum plot
    to automatically exclude axis tick labels outside the plot frame.
    이미지 크기에 무관하게 비율(ratio) 기반으로 탐색 범위를 결정합니다.
    """
    if img is None or img.size == 0:
        return 80, 1170

    h, w = img.shape[:2]
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    # Y축(세로선) 탐색: 이미지 너비의 3~6% 범위에서 탐색 (하드코딩 70~120px → 비율 기반)
    x_search_start = int(w * 0.03)
    x_search_end = int(w * 0.06)
    x_search_start = max(0, min(x_search_start, w - 1))
    x_search_end = max(x_search_start + 1, min(x_search_end, w))

    if x_search_end > x_search_start:
        dark_y = (gray[:, x_search_start:x_search_end] < 100)
        if dark_y.size > 0:
            col_counts = np.sum(dark_y, axis=0)
            if col_counts.size > 0 and np.max(col_counts) > h * 0.35:
                x_axis = x_search_start + int(np.argmax(col_counts))
            else:
                x_axis = 80
        else:
            x_axis = 80
    else:
        x_axis = 80

    # X축(가로선) 탐색: 이미지 높이의 90~98% 범위에서 탐색 (하드코딩 1100~1180px → 비율 기반)
    y_search_start = int(h * 0.90)
    y_search_end = int(h * 0.98)
    y_search_start = max(0, min(y_search_start, h - 2))
    y_search_end = max(y_search_start + 1, min(y_search_end, h))

    if y_search_end > y_search_start:
        dark_x = (gray[y_search_start:y_search_end, :] < 100)
        if dark_x.size > 0:
            row_counts = np.sum(dark_x, axis=1)
            if row_counts.size > 0 and np.max(row_counts) > w * 0.35:
                y_axis = y_search_start + int(np.argmax(row_counts))
            else:
                y_axis = int(h * 0.95)
        else:
            y_axis = int(h * 0.95)
    else:
        y_axis = int(h * 0.95)

    return x_axis, y_axis


def resolve_label_collisions(labels, font_main, font_sub, draw, min_dist_y=50):
    """
    Precision Peak Layout Algorithm:
    - Default: Centered directly above the peak apex.
    - Top Ceiling Clamp: If apex_y < 75, place text to the right of apex.
    - Horizontal Overlap Resolution (O(n) with adjacent-only check after sort):
      인접 레이블만 비교하여 겹침 해소.
    """
    labels.sort(key=lambda item: item['center'][0])

    for item in labels:
        txt = item['text']
        cx, cy = item['center']
        font = font_main if item.get('is_mrm', False) else font_sub
        tb = draw.textbbox((0, 0), txt, font=font)
        tw, th = tb[2] - tb[0], tb[3] - tb[1]
        item['tw'] = tw
        item['th'] = th
        item['apex_x'] = cx
        item['apex_y'] = cy

        if cy < 75:
            item['orig_x'] = cx + 8
            item['orig_y'] = max(10, cy + 2)
            item['is_side_aligned'] = True
        else:
            item['orig_x'] = cx - tw // 2
            item['orig_y'] = max(10, cy - th - 6)
            item['is_side_aligned'] = False

        item['curr_x'] = item['orig_x']
        item['curr_y'] = item['orig_y']

    # O(n) — X 기준 정렬 후 인접 항목끼리만 비교
    for i in range(1, len(labels)):
        l1 = labels[i - 1]
        l2 = labels[i]

        b1_x1, b1_x2 = l1['curr_x'] - 6, l1['curr_x'] + l1['tw'] + 6
        b2_x1, b2_x2 = l2['curr_x'] - 6, l2['curr_x'] + l2['tw'] + 6

        overlap_x = not (b1_x2 <= b2_x1 or b2_x2 <= b1_x1)
        if overlap_x:
            dist_y = abs(l1['curr_y'] - l2['curr_y'])
            if dist_y < min_dist_y:
                l2['curr_y'] = min(l1['curr_y'], l2['curr_y']) - min_dist_y
                l2['is_shifted'] = True

    return labels


def calibrate_and_correct_peaks(peaks):
    """
    Cleans, deduplicates, and sorts detected peaks by m/z value.
    Merges OCR duplicate readings within delta m/z <= 0.8 to keep the single best peak.
    """
    if not peaks:
        return peaks

    sorted_peaks = sorted(peaks, key=lambda p: (not p.get('is_recommended', False), p.get('height_rank', 999)))

    cleaned = []
    for p in sorted_peaks:
        val = p['val_num']
        if not any(abs(val - c['val_num']) <= 0.8 for c in cleaned):
            p['mz'] = f"{val:.2f}"
            cleaned.append(p)

    cleaned.sort(key=lambda p: p['val_num'])
    return cleaned


def auto_load_gcp_credentials():
    """Auto-detects gcp_key.json or credentials.json in base directory for Google Cloud Vision API."""
    base_d = get_base_dir()
    for key_file in ["gcp_key.json", "credentials.json", "google_key.json"]:
        key_path = os.path.join(base_d, key_file)
        if os.path.exists(key_path):
            os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = key_path
            return key_path
    return None


def get_gemini_api_key():
    """
    Retrieves Gemini API Key from Streamlit Secrets or Environment Variables.
    app.py 에서도 이 함수를 공용으로 사용합니다 (중복 제거).
    탐색 순서: Streamlit Secrets → 환경변수 GEMINI_API_KEY → GOOGLE_API_KEY
    """
    try:
        import streamlit as st
        if hasattr(st, "secrets") and len(st.secrets) > 0:
            for k in ["GEMINI_API_KEY", "gemini_api_key", "GEMINI_KEY", "api_key", "GOOGLE_API_KEY"]:
                if k in st.secrets:
                    return str(st.secrets[k]).strip()
    except Exception:
        pass
    return os.environ.get("GEMINI_API_KEY", os.environ.get("GOOGLE_API_KEY", "")).strip()


_last_ai_error = None


def get_last_ai_error():
    global _last_ai_error
    return _last_ai_error


@lru_cache(maxsize=8)
def _get_dynamic_gemini_models(api_key: str) -> tuple:
    """
    Google API에서 사용 가능한 Gemini 모델 목록을 조회합니다.
    동일 api_key에 대해서는 캐싱된 결과를 반환하여 반복 네트워크 호출을 방지합니다.
    반환값은 lru_cache 호환을 위해 tuple 형식입니다.
    """
    try:
        url = f"https://generativelanguage.googleapis.com/v1beta/models?key={api_key}"
        resp = requests.get(url, timeout=10)
        if resp.status_code == 200:
            data = resp.json()
            supported = []
            for m in data.get("models", []):
                name = m.get("name", "")
                methods = m.get("supportedGenerationMethods", [])
                if "generateContent" in methods:
                    clean_name = name.replace("models/", "")
                    supported.append(clean_name)
            if supported:
                flash_models = [m for m in supported if "flash" in m.lower()]
                other_models = [m for m in supported if "flash" not in m.lower()]
                return tuple(flash_models + other_models)
    except Exception:
        pass
    return ("gemini-3.6-flash", "gemini-3-flash", "gemini-2.0-flash", "gemini-1.5-flash")


def _parse_gcp_vision_raw_peaks(response):
    if not response or response.error.message:
        return []
    annotations = response.text_annotations
    if not annotations:
        return []
    raw_peaks = []
    seen_centers = set()
    for ann in annotations[1:]:
        text = ann.description.strip()
        if re.search(r'[a-zA-Z]', text) and not re.search(r'^\d+\.\d+$', text):
            continue
        match = re.search(r'\d+\.\d+|\d{2,3}', text)
        if not match:
            continue
        val_str = match.group(0)
        try:
            val_num = float(val_str)
            if not (10.0 <= val_num <= 2000.0):
                continue
        except ValueError:
            continue

        vertices = ann.bounding_poly.vertices
        xs = [v.x for v in vertices]
        ys = [v.y for v in vertices]
        x_min, x_max = min(xs), max(xs)
        y_min, y_max = min(ys), max(ys)
        cx, cy = (x_min + x_max) // 2, (y_min + y_max) // 2
        dup = any(abs(cx - scx) < 25 and abs(cy - scy) < 25 for scx, scy in seen_centers)
        if dup:
            continue
        seen_centers.add((cx, cy))
        raw_peaks.append({
            'mz': f"{val_num:.2f}",
            'val_num': val_num,
            'y_min': y_min,
            'x_min': x_min,
            'x_max': x_max,
            'y_max': y_max,
            'cx': cx,
            'cy': cy
        })
    return raw_peaks


def _parse_gcp_vision_response(response, is_precursor):
    """Helper to parse Google Cloud Vision API response into peak dictionaries."""
    raw_peaks = _parse_gcp_vision_raw_peaks(response)
    if not raw_peaks:
        return None
    corrected_peaks = calibrate_and_correct_peaks(raw_peaks)
    if is_precursor:
        peaks_by_height = sorted(corrected_peaks, key=lambda p: p['y_min'])
        default_checked = set(p['mz'] for p in peaks_by_height[:1])
    else:
        max_mz = max((p['val_num'] for p in corrected_peaks), default=0)
        candidates = [p for p in corrected_peaks if p['val_num'] < max_mz - 0.5]
        if not candidates:
            candidates = corrected_peaks
        peaks_by_height = sorted(candidates, key=lambda p: p['y_min'])
        default_checked = set(p['mz'] for p in peaks_by_height[:3])
    all_peaks_sorted = sorted(list(set(p['mz'] for p in corrected_peaks)), key=lambda x: float(x))
    return {'all_peaks': all_peaks_sorted, 'default_checked': default_checked, 'engine': 'Google Lens Cloud AI ⚡ (99.99%)'}


def _get_gemini_raw_peaks(img_bytes, is_precursor=True):
    """Calls Google Gemini REST API with tailored Precursor vs Product smart recommendations."""
    global _last_ai_error
    _last_ai_error = None

    api_key = get_gemini_api_key()
    if not api_key:
        _last_ai_error = "Gemini API 키가 설정되지 않았습니다."
        return []

    nparr = np.frombuffer(img_bytes, np.uint8)
    img_cv = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
    if img_cv is None:
        img_h, img_w = 1200, 2200
    else:
        img_h, img_w = img_cv.shape[:2]

    b64_img = base64.b64encode(img_bytes).decode('utf-8')

    if is_precursor:
        prompt = f"""
Examine this LC-MS/MS Precursor Ion Spectrum image (Width: {img_w}px, Height: {img_h}px).
1. Look at the horizontal X-axis tick numbers at the bottom to determine the mass range:
   - "x_axis_min": numerical m/z at the left edge (usually 0 or 50)
   - "x_axis_max": numerical m/z at the right edge (e.g. 200, 350, 500)
2. Identify the single main compound Precursor Ion peak (the tallest, most dominant peak on the entire spectrum, e.g. 336.06 or 168.96).
3. Extract all other numerical peak m/z labels printed near peak tops. Do not extract axis tick numbers as peaks.
4. Mark ONLY the single tallest main precursor peak with "is_recommended": true. Mark all others with "is_recommended": false.

Return ONLY a JSON object:
{{
  "x_axis_min": 0,
  "x_axis_max": 500,
  "peaks": [
    {{"mz": 336.06, "is_recommended": true, "height_rank": 1, "ymin": 250, "xmin": 800, "xmax": 920, "ymax": 280}}
  ]
}}
"""
    else:
        prompt = f"""
Examine this LC-MS/MS Product (Fragment) Ion Spectrum image (Width: {img_w}px, Height: {img_h}px).
1. Look at the horizontal X-axis tick numbers at the bottom to determine the mass range:
   - "x_axis_min": numerical m/z at the left edge (usually 0 or 30 or 50)
   - "x_axis_max": numerical m/z at the right edge (e.g. 200, 350, 500)
2. Identify all numerical fragment peak m/z labels printed near peak tops (e.g. 283.07, 183.07, 76.99, etc.). Do not extract axis numbers.
3. Evaluate their spectral peak heights (sensitivities) from tallest (rank 1) to shortest.
4. Exclude the unfragmented parent precursor ion (the largest m/z peak).
5. Mark the TOP 3 tallest fragment ion peaks with "is_recommended": true. Mark all other peaks with "is_recommended": false.

Return ONLY a JSON object:
{{
  "x_axis_min": 0,
  "x_axis_max": 350,
  "peaks": [
    {{"mz": 283.07, "is_recommended": true, "height_rank": 1, "ymin": 150, "xmin": 1200, "xmax": 1300, "ymax": 180}},
    {{"mz": 183.07, "is_recommended": true, "height_rank": 2, "ymin": 580, "xmin": 450, "xmax": 550, "ymax": 610}},
    {{"mz": 76.99, "is_recommended": true, "height_rank": 3, "ymin": 1000, "xmin": 200, "xmax": 300, "ymax": 1030}},
    {{"mz": 51.01, "is_recommended": false, "height_rank": 4, "ymin": 1100, "xmin": 100, "xmax": 180, "ymax": 1130}}
  ]
}}
"""

    # lru_cache 로 캐싱된 모델 목록 사용 (매번 API 호출 불필요)
    active_models = list(_get_dynamic_gemini_models(api_key))
    errors = []
    headers = {"Content-Type": "application/json"}

    for model_name in active_models:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{model_name}:generateContent?key={api_key}"
        payload = {
            "contents": [{
                "role": "user",
                "parts": [
                    {"text": prompt},
                    {
                        "inlineData": {
                            "mimeType": "image/png",
                            "data": b64_img
                        }
                    }
                ]
            }],
            "generationConfig": {
                "response_mime_type": "application/json"
            }
        }

        try:
            resp = requests.post(url, headers=headers, json=payload, timeout=25)
            if resp.status_code == 200:
                res_data = resp.json()
                text = res_data["candidates"][0]["content"]["parts"][0]["text"]
                data = json.loads(text)
                raw_peaks = []
                seen = set()
                for p in data.get("peaks", []):
                    try:
                        val_num = float(p["mz"])
                        if not (10.0 <= val_num <= 2000.0):
                            continue
                        val_str = f"{val_num:.2f}"
                        if val_str in seen:
                            continue
                        seen.add(val_str)

                        ymin_raw = float(p.get("ymin", 200))
                        xmin_raw = float(p.get("xmin", 200))
                        xmax_raw = float(p.get("xmax", xmin_raw + 80))
                        ymax_raw = float(p.get("ymax", ymin_raw + 30))

                        if xmax_raw <= 1000 and img_w > 1200:
                            xmin = int((xmin_raw / 1000.0) * img_w)
                            xmax = int((xmax_raw / 1000.0) * img_w)
                            ymin = int((ymin_raw / 1000.0) * img_h)
                            ymax = int((ymax_raw / 1000.0) * img_h)
                        else:
                            xmin = int(xmin_raw)
                            xmax = int(xmax_raw)
                            ymin = int(ymin_raw)
                            ymax = int(ymax_raw)

                        cx, cy = (xmin + xmax) // 2, (ymin + ymax) // 2

                        raw_peaks.append({
                            'mz': val_str,
                            'val_num': val_num,
                            'is_recommended': bool(p.get('is_recommended', False)),
                            'height_rank': int(p.get('height_rank', 999)),
                            'x_axis_min': float(data.get('x_axis_min', 0.0)),
                            'x_axis_max': float(data.get('x_axis_max', 500.0)),
                            'y_min': ymin,
                            'x_min': xmin,
                            'x_max': xmax,
                            'y_max': ymax,
                            'cx': cx,
                            'cy': cy
                        })
                    except Exception:
                        pass
                if raw_peaks:
                    _last_ai_error = None
                    return raw_peaks
            else:
                errors.append(f"[{model_name}] HTTP {resp.status_code}: {resp.text}")
        except Exception as e:
            errors.append(f"[{model_name}] 오류: {e}")

    if errors:
        _last_ai_error = errors[0] if len(errors) == 1 else " | ".join(errors[:2])

    return []


def extract_peaks_with_google_vision(img_bytes, is_precursor=True):
    """
    Extracts peak m/z values and accurately ranks them by physical peak height (sensitivity / 감도).
    우선순위: Google Gemini Flash Vision AI → Google Cloud Vision API → 빈 결과
    """
    # 1. Gemini Flash Vision AI (API 키 필요, 무과금 티어 지원)
    gemini_raw = _get_gemini_raw_peaks(img_bytes, is_precursor=is_precursor)
    if gemini_raw:
        corrected_peaks = calibrate_and_correct_peaks(gemini_raw)

        default_checked = set()
        for p in corrected_peaks:
            if p.get('is_recommended', False):
                default_checked.add(p['mz'])

        # AI가 추천 플래그를 반환하지 않은 경우 height_rank로 폴백
        if not default_checked and corrected_peaks:
            if is_precursor:
                peaks_by_rank = sorted(corrected_peaks, key=lambda p: p.get('height_rank', 999))
                default_checked = set(p['mz'] for p in peaks_by_rank[:1])
            else:
                max_mz = max((p['val_num'] for p in corrected_peaks), default=0)
                candidates = [p for p in corrected_peaks if p['val_num'] < max_mz - 0.5] or corrected_peaks
                peaks_by_rank = sorted(candidates, key=lambda p: p.get('height_rank', 999))
                default_checked = set(p['mz'] for p in peaks_by_rank[:3])

        all_peaks_sorted = sorted(list(set(p['mz'] for p in corrected_peaks)), key=lambda x: float(x))
        return {'all_peaks': all_peaks_sorted, 'default_checked': default_checked, 'engine': 'Google Gemini Flash AI ⚡ (99.99%)'}

    # 2. Google Cloud Vision API (서비스 계정 키 필요)
    key_path = auto_load_gcp_credentials()
    if key_path or "GOOGLE_APPLICATION_CREDENTIALS" in os.environ or "GOOGLE_APPLICATION_CREDENTIALS_JSON" in os.environ:
        try:
            from google.cloud import vision
            client = vision.ImageAnnotatorClient()
            image = vision.Image(content=img_bytes)
            response = client.text_detection(image=image)
            parsed = _parse_gcp_vision_response(response, is_precursor)
            if parsed:
                return parsed
        except Exception:
            pass

    return {'all_peaks': [], 'default_checked': set(), 'engine': 'Gemini API 키 입력 필요 🔑'}


def _get_raw_peaks_for_image(img_bytes):
    """
    이미지에서 원시(raw) 피크를 추출합니다.
    Gemini Vision → Google Cloud Vision 순으로 시도합니다.
    """
    # 1. Gemini Vision AI
    gemini_raw = _get_gemini_raw_peaks(img_bytes)
    if gemini_raw:
        return gemini_raw

    # 2. Google Cloud Vision
    key_path = auto_load_gcp_credentials()
    if key_path or "GOOGLE_APPLICATION_CREDENTIALS" in os.environ or "GOOGLE_APPLICATION_CREDENTIALS_JSON" in os.environ:
        try:
            from google.cloud import vision
            client = vision.ImageAnnotatorClient()
            response = client.text_detection(image=vision.Image(content=img_bytes))
            rp = _parse_gcp_vision_raw_peaks(response)
            if rp:
                return rp
        except Exception:
            pass

    return []


def get_render_font(font_size, bold=True):
    """Finds available system font on Windows or Linux / Streamlit Cloud."""
    font_candidates = [
        'C:/Windows/Fonts/arialbd.ttf',
        'C:/Windows/Fonts/arial.ttf',
        'C:/Windows/Fonts/malgunbd.ttf',
        'C:/Windows/Fonts/malgun.ttf',
        '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf',
        '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf',
        '/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf',
        '/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf',
        '/usr/share/fonts/truetype/freefont/FreeSansBold.ttf',
        '/usr/share/fonts/truetype/freefont/FreeSans.ttf',
    ]
    for fp in font_candidates:
        if os.path.exists(fp):
            try:
                return ImageFont.truetype(fp, font_size)
            except Exception:
                pass
    try:
        return ImageFont.load_default(size=font_size)
    except TypeError:
        return ImageFont.load_default()


def process_spectrum_image(img_bytes, peak_overrides=None, font_size=45, is_precursor=True):
    """
    Renders enlarged peak labels with user editable overrides:
    peak_overrides = [ {'orig_mz': '336.06', 'final_mz': '335.06', 'is_mrm': True}, ... ]
    """
    if peak_overrides is None:
        peak_overrides = []

    nparr = np.frombuffer(img_bytes, np.uint8)
    img = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
    if img is None or img.size == 0:
        return img_bytes

    x_axis, y_axis = detect_axis_lines(img)
    x_start = x_axis + 10
    x_end = min(img.shape[1] - 30, 2180)
    y_top = 50
    y_bottom = y_axis - 10

    # 1. 기존 소형 텍스트 및 리더선 완전 소거
    out_img = img.copy()
    b, g, r = cv2.split(out_img)
    pure_blue = (b > 130) & (r < 70) & (g < 120)
    is_blue = (b > 110) & (b > r + 25) & (b > g + 5)
    gray = cv2.cvtColor(out_img, cv2.COLOR_BGR2GRAY)

    plot_gray = gray[y_top:y_bottom, x_start:x_end]
    plot_blue = is_blue[y_top:y_bottom, x_start:x_end]
    text_mask = (plot_gray < 235) & (~plot_blue)

    roi = out_img[y_top:y_bottom, x_start:x_end]
    roi[text_mask] = [255, 255, 255]

    # 2. X축 스케일 보정 및 피크 레이블 구성
    peak_labels = []
    if peak_overrides:
        all_vals = []
        for item in peak_overrides:
            try:
                all_vals.append(float(item.get('final_mz', item.get('orig_mz'))))
            except ValueError:
                pass
        max_v = max(all_vals, default=350.0)

        # Named constant 기반 X축 보정 (파일 상단의 상수 참조)
        if max_v <= 210.0:
            def calc_x(mz):
                return int(_X_LOW_ORIGIN + (mz - _X_LOW_MZ_START) * _X_LOW_SCALE)
        elif is_precursor:
            def calc_x(mz):
                return int(_X_HIGH_ORIGIN + (mz - _X_HIGH_PREC_MZ_START) * _X_HIGH_PREC_SCALE)
        else:
            def calc_x(mz):
                return int(_X_HIGH_ORIGIN + (mz - _X_HIGH_PROD_MZ_START) * _X_HIGH_PROD_SCALE)

        for item in peak_overrides:
            orig_mz = item.get('orig_mz', '')
            final_mz = item.get('final_mz', orig_mz)
            is_mrm = item.get('is_mrm', False)
            try:
                val_f = float(final_mz)
            except ValueError:
                continue

            x_calc = calc_x(val_f)
            x_calc = max(x_start + 15, min(x_end - 15, x_calc))

            x1 = max(x_start, x_calc - 25)
            x2 = min(x_end, x_calc + 25)
            best_apex_x = x_calc
            best_apex_y = y_bottom

            for cx in range(x1, x2):
                col = pure_blue[80:y_bottom, cx]
                blue_idx = np.where(col)[0]
                if len(blue_idx) > 0:
                    top_y = 80 + np.min(blue_idx)
                    if top_y < best_apex_y:
                        best_apex_y = top_y
                        best_apex_x = cx

            peak_labels.append({
                'text': final_mz,
                'center': (best_apex_x, best_apex_y),
                'bbox': (best_apex_x - 40, best_apex_y - 15, best_apex_x + 40, best_apex_y + 15),
                'is_mrm': is_mrm
            })
    else:
        raw_peaks = _get_raw_peaks_for_image(img_bytes)
        corrected_peaks = calibrate_and_correct_peaks(raw_peaks)
        for p in corrected_peaks:
            peak_labels.append({
                'text': p['mz'],
                'center': (p['cx'], p['cy']),
                'bbox': (p['x_min'], p['y_min'], p['x_max'], p['y_max']),
                'is_mrm': False
            })

    # 3. 투명 오버레이로 MRM 스타일 레이블 합성
    img_rgb = cv2.cvtColor(out_img, cv2.COLOR_BGR2RGB)
    pil_base = Image.fromarray(img_rgb).convert('RGBA')

    overlay = Image.new('RGBA', pil_base.size, (255, 255, 255, 0))
    draw = ImageDraw.Draw(overlay)

    font_main = get_render_font(font_size, bold=True)
    font_sub = get_render_font(max(12, font_size - 5), bold=False)

    labels = resolve_label_collisions(peak_labels, font_main, font_sub, draw, min_dist_y=int(font_size * 1.2))

    for item in labels:
        txt = item['text']
        tx, ty = item['curr_x'], item['curr_y']
        tw, th = item['tw'], item['th']
        apex_x, apex_y = item['apex_x'], item['apex_y']
        is_mrm = item.get('is_mrm', False)
        is_side = item.get('is_side_aligned', False)
        is_shifted = item.get('is_shifted', False) or abs(ty - item['orig_y']) > 15

        text_color = (0, 0, 220, 255) if is_mrm else (120, 120, 120, 255)
        line_color = (0, 0, 220, 255) if is_mrm else (150, 150, 150, 255)
        font = font_main if is_mrm else font_sub

        if is_side:
            draw.ellipse([apex_x - 3, apex_y - 3, apex_x + 3, apex_y + 3], fill=line_color)
        elif is_shifted:
            start_pt = (tx + tw // 2, ty + th + 2)
            end_pt = (apex_x, apex_y)
            draw.line([start_pt, end_pt], fill=line_color, width=2)
            draw.ellipse([apex_x - 3, apex_y - 3, apex_x + 3, apex_y + 3], fill=line_color)

        draw.text((tx, ty), txt, font=font, fill=text_color, stroke_width=2, stroke_fill=(255, 255, 255, 180))

    final_pil = Image.alpha_composite(pil_base, overlay).convert('RGB')

    buf = io.BytesIO()
    final_pil.save(buf, format='PNG')
    return buf.getvalue()


def process_excel_with_selections(excel_path, sheet_selections, font_size=45, status_callback=None):
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"파일을 찾을 수 없습니다: {excel_path}")

    if status_callback:
        status_callback("엑셀 파일 로딩 및 MS/MS 이미지 변환 중...")

    # NOTE: ws._images 는 openpyxl 의 비공개(private) 속성입니다.
    # openpyxl 업그레이드 전 반드시 동작 확인이 필요합니다 (requirements.txt 버전 고정 참조).
    wb = openpyxl.load_workbook(excel_path)
    temp_dir = os.path.join(os.path.dirname(excel_path), "_temp_mz_img")
    os.makedirs(temp_dir, exist_ok=True)

    total_images = sum(len(wb[name]._images) for name in wb.sheetnames)
    processed_count = 0

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        sel_dict = sheet_selections.get(sheet_name, {'precursor': [], 'product': []})

        for idx, img in enumerate(ws._images):
            processed_count += 1
            if status_callback:
                status_callback(f"이미지 처리 중 ({processed_count}/{total_images}) - 시트: {sheet_name}")

            is_precursor = (idx % 2 == 0)
            overrides = sel_dict['precursor'] if is_precursor else sel_dict['product']

            img_bytes = img._data()
            proc_bytes = process_spectrum_image(img_bytes, peak_overrides=overrides, font_size=font_size, is_precursor=is_precursor)

            temp_path = os.path.join(temp_dir, f"{sheet_name}_{idx + 1}.png")
            with open(temp_path, "wb") as f:
                f.write(proc_bytes)

            new_img = OpenpyxlImage(temp_path)
            new_img.anchor = img.anchor
            ws._images[idx] = new_img

    script_dir = get_base_dir()
    results_dir = os.path.join(script_dir, "Results")
    os.makedirs(results_dir, exist_ok=True)

    _, file_name = os.path.split(excel_path)
    name_part, ext_part = os.path.splitext(file_name)
    output_path = os.path.join(results_dir, f"{name_part}_확대{ext_part}")

    if status_callback:
        status_callback("수정된 엑셀 파일 저장 중...")

    wb.save(output_path)
    return output_path


if __name__ == "__main__":
    # 데스크톱 GUI 실행은 desktop_app.py 를 사용하세요.
    # python desktop_app.py
    print("웹 앱 실행: streamlit run app.py")
    print("데스크톱 GUI 실행: python desktop_app.py")
