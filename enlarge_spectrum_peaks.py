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
# ──────────────────────────────────────────────────────────────────────────────

# ── X축 고정 상수 (axis_ticks 보간 실패 시 폴백 전용) ─────────────────────────
_X_LOW_ORIGIN       = 40
_X_LOW_MZ_START     = 30.0
_X_LOW_SCALE        = 13.48

_X_HIGH_ORIGIN          = 59
_X_HIGH_PREC_MZ_START   = 50.0
_X_HIGH_PREC_SCALE      = 4.75333
_X_HIGH_PROD_MZ_START   = 20.0
_X_HIGH_PROD_SCALE      = 6.684375
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
    Detects Y-axis (vertical) and X-axis (horizontal) lines in the spectrum plot.
    이미지 크기 비율 기반으로 탐색 범위를 결정합니다.
    """
    if img is None or img.size == 0:
        return 80, 1170

    h, w = img.shape[:2]
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    x_search_start = max(0, min(int(w * 0.03), w - 1))
    x_search_end   = max(x_search_start + 1, min(int(w * 0.06), w))

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

    y_search_start = max(0, min(int(h * 0.90), h - 2))
    y_search_end   = max(y_search_start + 1, min(int(h * 0.98), h))

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


def _linear_x_from_ticks(mz, axis_ticks, x_start, x_end):
    """
    두 개 이상의 axis_tick (m/z, pixel_x) 을 이용해 선형 보간으로 픽셀 X 위치를 반환합니다.
    axis_ticks가 없거나 부족하면 None을 반환 → 호출자에서 고정 상수 폴백 사용.
    """
    if not axis_ticks or len(axis_ticks) < 2:
        return None

    ticks = sorted(axis_ticks, key=lambda t: t['mz'])
    t1, t2 = ticks[0], ticks[-1]
    dmz = t2['mz'] - t1['mz']
    if abs(dmz) < 0.01:
        return None

    scale = (t2['pixel_x'] - t1['pixel_x']) / dmz
    x = int(t1['pixel_x'] + (mz - t1['mz']) * scale)
    return max(x_start + 5, min(x_end - 5, x))


def resolve_label_collisions(labels, font_main, font_sub, draw, min_dist_y=50):
    """
    Precision Peak Layout Algorithm.
    X 기준 정렬 후 인접 항목끼리만 비교하는 O(n) 충돌 해소.
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
    Merges OCR duplicate readings within delta m/z <= 0.8.
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
    """Auto-detects gcp_key.json or credentials.json in base directory."""
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
    탐색 순서: Streamlit Secrets → GEMINI_API_KEY → GOOGLE_API_KEY
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
    사용 가능한 Gemini 모델 목록을 조회합니다.
    동일 api_key에 대해 캐싱하여 반복 호출을 방지합니다.
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
                    supported.append(name.replace("models/", ""))
            if supported:
                flash_models = [m for m in supported if "flash" in m.lower()]
                other_models = [m for m in supported if "flash" not in m.lower()]
                return tuple(flash_models + other_models)
    except Exception:
        pass
    return ("gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash")


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
        if any(abs(cx - scx) < 25 and abs(cy - scy) < 25 for scx, scy in seen_centers):
            continue
        seen_centers.add((cx, cy))
        raw_peaks.append({
            'mz': f"{val_num:.2f}",
            'val_num': val_num,
            'y_min': y_min, 'x_min': x_min,
            'x_max': x_max, 'y_max': y_max,
            'cx': cx, 'cy': cy
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
        candidates = [p for p in corrected_peaks if p['val_num'] < max_mz - 0.5] or corrected_peaks
        peaks_by_height = sorted(candidates, key=lambda p: p['y_min'])
        default_checked = set(p['mz'] for p in peaks_by_height[:3])
    all_peaks_sorted = sorted(list(set(p['mz'] for p in corrected_peaks)), key=lambda x: float(x))
    return {'all_peaks': all_peaks_sorted, 'default_checked': default_checked,
            'axis_ticks': [], 'engine': 'Google Lens Cloud AI ⚡ (99.99%)'}


def _parse_axis_ticks(data, img_w, img_h):
    """
    Gemini 응답의 axis_ticks 를 파싱하여 {'mz': float, 'pixel_x': int} 목록을 반환합니다.
    AI가 반환한 좌표가 정규화 좌표(0~1000)인 경우 실제 픽셀로 변환합니다.
    """
    axis_ticks = []
    raw_ticks = data.get("axis_ticks", [])

    # axis_ticks 가 없으면 x_axis_min/max 와 plot 경계로 대체 시도
    if not raw_ticks:
        x_min_mz = data.get("x_axis_min")
        x_max_mz = data.get("x_axis_max")
        px_left  = data.get("plot_left_pixel")
        px_right = data.get("plot_right_pixel")
        if all(v is not None for v in [x_min_mz, x_max_mz, px_left, px_right]):
            raw_ticks = [
                {"mz": x_min_mz, "pixel_x": px_left},
                {"mz": x_max_mz, "pixel_x": px_right},
            ]

    for t in raw_ticks:
        try:
            mz_t  = float(t["mz"])
            px_t  = float(t["pixel_x"])
            # AI가 정규화 좌표(≤1000)를 반환했고 실제 이미지가 더 크면 스케일 변환
            if px_t <= 1000 and img_w > 1200:
                px_t = (px_t / 1000.0) * img_w
            axis_ticks.append({'mz': mz_t, 'pixel_x': int(px_t)})
        except (KeyError, ValueError, TypeError):
            pass

    return axis_ticks


def _get_gemini_raw_peaks(img_bytes, is_precursor=True):
    """
    Google Gemini REST API 로 피크 m/z 와 X축 눈금 픽셀 위치를 함께 추출합니다.
    반환값: {'peaks': [...], 'axis_ticks': [...]}
    """
    global _last_ai_error
    _last_ai_error = None

    api_key = get_gemini_api_key()
    if not api_key:
        _last_ai_error = "Gemini API 키가 설정되지 않았습니다."
        return {'peaks': [], 'axis_ticks': []}

    nparr = np.frombuffer(img_bytes, np.uint8)
    img_cv = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
    img_h, img_w = (img_cv.shape[:2] if img_cv is not None else (1200, 2200))

    b64_img = base64.b64encode(img_bytes).decode('utf-8')

    # ── 공통 axis_ticks 요청 문구 ─────────────────────────────────────────────
    axis_tick_instruction = """
IMPORTANT — X-axis calibration (required for accurate label placement):
Identify 2-3 clearly labeled X-axis tick marks.
For each, record:
  - "mz": the m/z number printed below the tick
  - "pixel_x": the pixel X coordinate of that tick mark in the image (use actual image pixels, not normalized)
Also record the pixel X of the leftmost plot border and rightmost plot border:
  - "plot_left_pixel": pixel X of the Y-axis (left plot border)
  - "plot_right_pixel": pixel X of the rightmost axis border
"""

    if is_precursor:
        prompt = f"""
Examine this LC-MS/MS Precursor Ion Spectrum image (Width: {img_w}px, Height: {img_h}px).

1. Look at the horizontal X-axis tick numbers at the bottom to determine the mass range:
   - "x_axis_min": numerical m/z at the left edge
   - "x_axis_max": numerical m/z at the right edge

2. Identify the single main compound Precursor Ion peak (the tallest, most dominant peak).

3. Extract all other numerical peak m/z labels printed near peak tops.
   Do NOT extract axis tick numbers as peaks.

4. Mark ONLY the single tallest main precursor peak with "is_recommended": true.

{axis_tick_instruction}

Return ONLY a JSON object:
{{
  "x_axis_min": 332,
  "x_axis_max": 338,
  "plot_left_pixel": 85,
  "plot_right_pixel": 2170,
  "axis_ticks": [
    {{"mz": 332.0, "pixel_x": 85}},
    {{"mz": 335.0, "pixel_x": 1128}},
    {{"mz": 338.0, "pixel_x": 2170}}
  ],
  "peaks": [
    {{"mz": 335.06, "is_recommended": true, "height_rank": 1, "ymin": 50, "xmin": 1100, "xmax": 1200, "ymax": 80}}
  ]
}}
"""
    else:
        prompt = f"""
Examine this LC-MS/MS Product (Fragment) Ion Spectrum image (Width: {img_w}px, Height: {img_h}px).

1. Look at the horizontal X-axis tick numbers at the bottom to determine the mass range:
   - "x_axis_min": numerical m/z at the left edge
   - "x_axis_max": numerical m/z at the right edge

2. Identify all numerical fragment peak m/z labels printed near peak tops.
   Do NOT extract axis numbers.

3. Evaluate spectral peak heights from tallest (rank 1) to shortest.

4. Exclude the unfragmented parent precursor ion (the largest m/z peak).

5. Mark the TOP 3 tallest fragment ion peaks with "is_recommended": true.

{axis_tick_instruction}

Return ONLY a JSON object:
{{
  "x_axis_min": 20,
  "x_axis_max": 350,
  "plot_left_pixel": 85,
  "plot_right_pixel": 2170,
  "axis_ticks": [
    {{"mz": 50.0,  "pixel_x": 340}},
    {{"mz": 150.0, "pixel_x": 920}},
    {{"mz": 300.0, "pixel_x": 1800}}
  ],
  "peaks": [
    {{"mz": 183.07, "is_recommended": true,  "height_rank": 1, "ymin": 50,  "xmin": 1100, "xmax": 1200, "ymax": 80}},
    {{"mz": 76.99,  "is_recommended": true,  "height_rank": 2, "ymin": 200, "xmin": 400,  "xmax": 490,  "ymax": 230}},
    {{"mz": 51.01,  "is_recommended": false, "height_rank": 3, "ymin": 500, "xmin": 300,  "xmax": 380,  "ymax": 530}}
  ]
}}
"""

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
                    {"inlineData": {"mimeType": "image/png", "data": b64_img}}
                ]
            }],
            "generationConfig": {"response_mime_type": "application/json"}
        }

        try:
            resp = requests.post(url, headers=headers, json=payload, timeout=30)
            if resp.status_code == 200:
                res_data = resp.json()
                text = res_data["candidates"][0]["content"]["parts"][0]["text"]
                data = json.loads(text)

                # ── 피크 파싱 ──────────────────────────────────────────────
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
                            xmin, xmax = int(xmin_raw), int(xmax_raw)
                            ymin, ymax = int(ymin_raw), int(ymax_raw)

                        cx, cy = (xmin + xmax) // 2, (ymin + ymax) // 2
                        raw_peaks.append({
                            'mz': val_str,
                            'val_num': val_num,
                            'is_recommended': bool(p.get('is_recommended', False)),
                            'height_rank': int(p.get('height_rank', 999)),
                            'x_axis_min': float(data.get('x_axis_min', 0.0)),
                            'x_axis_max': float(data.get('x_axis_max', 500.0)),
                            'y_min': ymin, 'x_min': xmin,
                            'x_max': xmax, 'y_max': ymax,
                            'cx': cx, 'cy': cy
                        })
                    except Exception:
                        pass

                # ── axis_ticks 파싱 ────────────────────────────────────────
                axis_ticks = _parse_axis_ticks(data, img_w, img_h)

                if raw_peaks:
                    _last_ai_error = None
                    return {'peaks': raw_peaks, 'axis_ticks': axis_ticks}
            else:
                errors.append(f"[{model_name}] HTTP {resp.status_code}: {resp.text[:200]}")
        except Exception as e:
            errors.append(f"[{model_name}] 오류: {e}")

    if errors:
        _last_ai_error = errors[0] if len(errors) == 1 else " | ".join(errors[:2])

    return {'peaks': [], 'axis_ticks': []}


def extract_peaks_with_google_vision(img_bytes, is_precursor=True):
    """
    피크 m/z 값 및 X축 보정 눈금을 추출합니다.
    우선순위: Gemini Flash AI → Google Cloud Vision API
    반환값: {'all_peaks': [...], 'default_checked': set(), 'axis_ticks': [...], 'engine': '...'}
    """
    # 1. Gemini Flash Vision AI
    gemini_result = _get_gemini_raw_peaks(img_bytes, is_precursor=is_precursor)
    gemini_raw    = gemini_result['peaks']
    axis_ticks    = gemini_result['axis_ticks']

    if gemini_raw:
        corrected_peaks = calibrate_and_correct_peaks(gemini_raw)

        default_checked = {p['mz'] for p in corrected_peaks if p.get('is_recommended', False)}

        if not default_checked and corrected_peaks:
            if is_precursor:
                peaks_by_rank = sorted(corrected_peaks, key=lambda p: p.get('height_rank', 999))
                default_checked = {p['mz'] for p in peaks_by_rank[:1]}
            else:
                max_mz = max((p['val_num'] for p in corrected_peaks), default=0)
                candidates = [p for p in corrected_peaks if p['val_num'] < max_mz - 0.5] or corrected_peaks
                peaks_by_rank = sorted(candidates, key=lambda p: p.get('height_rank', 999))
                default_checked = {p['mz'] for p in peaks_by_rank[:3]}

        all_peaks_sorted = sorted({p['mz'] for p in corrected_peaks}, key=lambda x: float(x))
        return {
            'all_peaks': all_peaks_sorted,
            'default_checked': default_checked,
            'axis_ticks': axis_ticks,
            'engine': 'Google Gemini Flash AI ⚡ (99.99%)'
        }

    # 2. Google Cloud Vision API
    key_path = auto_load_gcp_credentials()
    if key_path or "GOOGLE_APPLICATION_CREDENTIALS" in os.environ:
        try:
            from google.cloud import vision
            client   = vision.ImageAnnotatorClient()
            response = client.text_detection(image=vision.Image(content=img_bytes))
            parsed   = _parse_gcp_vision_response(response, is_precursor)
            if parsed:
                return parsed
        except Exception:
            pass

    return {'all_peaks': [], 'default_checked': set(), 'axis_ticks': [], 'engine': 'Gemini API 키 입력 필요 🔑'}


def _get_raw_peaks_for_image(img_bytes):
    """
    오버라이드 없이 이미지에서 직접 피크를 추출합니다 (raw peaks only).
    """
    gemini_result = _get_gemini_raw_peaks(img_bytes)
    if gemini_result['peaks']:
        return gemini_result['peaks']

    key_path = auto_load_gcp_credentials()
    if key_path or "GOOGLE_APPLICATION_CREDENTIALS" in os.environ:
        try:
            from google.cloud import vision
            client   = vision.ImageAnnotatorClient()
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


def process_spectrum_image(img_bytes, peak_overrides=None, font_size=45, is_precursor=True, axis_ticks=None):
    """
    Renders enlarged peak labels with user editable overrides.

    peak_overrides : [{'orig_mz': '336.06', 'final_mz': '335.06', 'is_mrm': True}, ...]
    axis_ticks     : [{'mz': 332.0, 'pixel_x': 85}, {'mz': 338.0, 'pixel_x': 2170}]
                     Gemini 가 반환한 X축 눈금 보정 데이터. 있으면 선형 보간으로 정확한
                     픽셀 X 위치를 계산합니다. 없으면 고정 상수 폴백을 사용합니다.
    """
    if peak_overrides is None:
        peak_overrides = []
    if axis_ticks is None:
        axis_ticks = []

    nparr = np.frombuffer(img_bytes, np.uint8)
    img   = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
    if img is None or img.size == 0:
        return img_bytes

    x_axis, y_axis = detect_axis_lines(img)
    x_start  = x_axis + 10
    x_end    = min(img.shape[1] - 30, 2180)
    y_top    = 50
    y_bottom = y_axis - 10

    # 1. 기존 소형 텍스트 완전 소거
    out_img  = img.copy()
    b, g, r  = cv2.split(out_img)
    pure_blue = (b > 130) & (r < 70) & (g < 120)
    is_blue   = (b > 110) & (b > r + 25) & (b > g + 5)
    gray      = cv2.cvtColor(out_img, cv2.COLOR_BGR2GRAY)

    plot_gray = gray[y_top:y_bottom, x_start:x_end]
    plot_blue = is_blue[y_top:y_bottom, x_start:x_end]
    text_mask = (plot_gray < 235) & (~plot_blue)

    roi = out_img[y_top:y_bottom, x_start:x_end]
    roi[text_mask] = [255, 255, 255]

    # 2. X 위치 계산 함수 결정
    #    우선순위: ① axis_ticks 선형 보간 → ② 고정 상수 폴백
    has_ticks = len(axis_ticks) >= 2

    if not has_ticks:
        # 고정 상수 폴백 (기존 로직)
        all_vals = []
        for item in peak_overrides:
            try:
                all_vals.append(float(item.get('final_mz', item.get('orig_mz'))))
            except ValueError:
                pass
        max_v = max(all_vals, default=350.0)

        if max_v <= 210.0:
            def _fallback_x(mz):
                return int(_X_LOW_ORIGIN + (mz - _X_LOW_MZ_START) * _X_LOW_SCALE)
        elif is_precursor:
            def _fallback_x(mz):
                return int(_X_HIGH_ORIGIN + (mz - _X_HIGH_PREC_MZ_START) * _X_HIGH_PREC_SCALE)
        else:
            def _fallback_x(mz):
                return int(_X_HIGH_ORIGIN + (mz - _X_HIGH_PROD_MZ_START) * _X_HIGH_PROD_SCALE)

    def calc_x(mz):
        if has_ticks:
            result = _linear_x_from_ticks(mz, axis_ticks, x_start, x_end)
            if result is not None:
                return result
        return max(x_start + 15, min(x_end - 15, _fallback_x(mz) if not has_ticks else x_start + 15))

    # 3. 피크 레이블 구성
    peak_labels = []
    if peak_overrides:
        for item in peak_overrides:
            orig_mz  = item.get('orig_mz', '')
            final_mz = item.get('final_mz', orig_mz)
            is_mrm   = item.get('is_mrm', False)
            try:
                val_f = float(final_mz)
            except ValueError:
                continue

            x_calc = calc_x(val_f)
            x_calc = max(x_start + 15, min(x_end - 15, x_calc))

            # 계산된 X 주변에서 실제 피크 정점(파란 선 최고점) 탐색
            scan_radius = 60 if has_ticks else 25
            x1 = max(x_start, x_calc - scan_radius)
            x2 = min(x_end, x_calc + scan_radius)
            best_apex_x = x_calc
            best_apex_y = y_bottom

            for cx in range(x1, x2):
                col       = pure_blue[80:y_bottom, cx]
                blue_idx  = np.where(col)[0]
                if len(blue_idx) > 0:
                    top_y = 80 + np.min(blue_idx)
                    if top_y < best_apex_y:
                        best_apex_y = top_y
                        best_apex_x = cx

            peak_labels.append({
                'text':   final_mz,
                'center': (best_apex_x, best_apex_y),
                'bbox':   (best_apex_x - 40, best_apex_y - 15, best_apex_x + 40, best_apex_y + 15),
                'is_mrm': is_mrm
            })
    else:
        raw_peaks       = _get_raw_peaks_for_image(img_bytes)
        corrected_peaks = calibrate_and_correct_peaks(raw_peaks)
        for p in corrected_peaks:
            peak_labels.append({
                'text':   p['mz'],
                'center': (p['cx'], p['cy']),
                'bbox':   (p['x_min'], p['y_min'], p['x_max'], p['y_max']),
                'is_mrm': False
            })

    # 4. 투명 오버레이로 레이블 합성
    img_rgb  = cv2.cvtColor(out_img, cv2.COLOR_BGR2RGB)
    pil_base = Image.fromarray(img_rgb).convert('RGBA')
    overlay  = Image.new('RGBA', pil_base.size, (255, 255, 255, 0))
    draw     = ImageDraw.Draw(overlay)

    font_main = get_render_font(font_size, bold=True)
    font_sub  = get_render_font(max(12, font_size - 5), bold=False)

    labels = resolve_label_collisions(peak_labels, font_main, font_sub, draw, min_dist_y=int(font_size * 1.2))

    for item in labels:
        txt    = item['text']
        tx, ty = item['curr_x'], item['curr_y']
        tw, th = item['tw'], item['th']
        apex_x, apex_y = item['apex_x'], item['apex_y']
        is_mrm     = item.get('is_mrm', False)
        is_side    = item.get('is_side_aligned', False)
        is_shifted = item.get('is_shifted', False) or abs(ty - item['orig_y']) > 15

        text_color = (0, 0, 220, 255)   if is_mrm else (120, 120, 120, 255)
        line_color = (0, 0, 220, 255)   if is_mrm else (150, 150, 150, 255)
        font       = font_main          if is_mrm else font_sub

        if is_side:
            draw.ellipse([apex_x - 3, apex_y - 3, apex_x + 3, apex_y + 3], fill=line_color)
        elif is_shifted:
            draw.line([(tx + tw // 2, ty + th + 2), (apex_x, apex_y)], fill=line_color, width=2)
            draw.ellipse([apex_x - 3, apex_y - 3, apex_x + 3, apex_y + 3], fill=line_color)

        draw.text((tx, ty), txt, font=font, fill=text_color, stroke_width=2, stroke_fill=(255, 255, 255, 180))

    final_pil = Image.alpha_composite(pil_base, overlay).convert('RGB')
    buf = io.BytesIO()
    final_pil.save(buf, format='PNG')
    return buf.getvalue()


def process_excel_with_selections(excel_path, sheet_selections, font_size=45, status_callback=None):
    """
    엑셀 내 모든 스펙트럼 이미지를 처리하여 새 엑셀 파일을 생성합니다.
    sheet_selections 구조:
      {
        sheet_name: {
          'precursor': [...peak dicts...],
          'product':   [...peak dicts...],
          'axis_ticks': {
            'precursor': [...tick dicts...],
            'product':   [...tick dicts...]
          }
        }
      }
    """
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"파일을 찾을 수 없습니다: {excel_path}")

    if status_callback:
        status_callback("엑셀 파일 로딩 및 MS/MS 이미지 변환 중...")

    # NOTE: ws._images 는 openpyxl 비공개 속성. 업그레이드 전 테스트 필요.
    wb       = openpyxl.load_workbook(excel_path)
    temp_dir = os.path.join(os.path.dirname(excel_path), "_temp_mz_img")
    os.makedirs(temp_dir, exist_ok=True)

    total_images    = sum(len(wb[name]._images) for name in wb.sheetnames)
    processed_count = 0

    for sheet_name in wb.sheetnames:
        ws       = wb[sheet_name]
        sel_dict = sheet_selections.get(sheet_name, {'precursor': [], 'product': [], 'axis_ticks': {}})

        ticks_dict = sel_dict.get('axis_ticks', {})

        for idx, img in enumerate(ws._images):
            processed_count += 1
            if status_callback:
                status_callback(f"이미지 처리 중 ({processed_count}/{total_images}) - 시트: {sheet_name}")

            ion_type   = 'precursor' if idx % 2 == 0 else 'product'
            is_prec    = (ion_type == 'precursor')
            overrides  = sel_dict.get(ion_type, [])
            axis_ticks = ticks_dict.get(ion_type, [])

            img_bytes  = img._data()
            proc_bytes = process_spectrum_image(
                img_bytes,
                peak_overrides=overrides,
                font_size=font_size,
                is_precursor=is_prec,
                axis_ticks=axis_ticks      # ← 선형 보간 보정 데이터 전달
            )

            temp_path = os.path.join(temp_dir, f"{sheet_name}_{idx + 1}.png")
            with open(temp_path, "wb") as f:
                f.write(proc_bytes)

            new_img        = OpenpyxlImage(temp_path)
            new_img.anchor = img.anchor
            ws._images[idx] = new_img

    script_dir  = get_base_dir()
    results_dir = os.path.join(script_dir, "Results")
    os.makedirs(results_dir, exist_ok=True)

    _, file_name          = os.path.split(excel_path)
    name_part, ext_part   = os.path.splitext(file_name)
    output_path           = os.path.join(results_dir, f"{name_part}_확대{ext_part}")

    if status_callback:
        status_callback("수정된 엑셀 파일 저장 중...")

    wb.save(output_path)
    return output_path


if __name__ == "__main__":
    print("웹 앱 실행: streamlit run app.py")
    print("데스크톱 GUI 실행: python desktop_app.py")
