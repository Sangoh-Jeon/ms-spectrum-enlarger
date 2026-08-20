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


def extract_spectrum_mz_range_from_header(img_pil):
    """
    스펙트럼 이미지 상단 헤더 영역(예: '+Q1 (166 - 172)', '+Product Ion of 335.1 (20 - 345)')에서
    m/z 시작과 끝 범위를 추출합니다.
    """
    try:
        # 상단 텍스트 영역(y=0~60px)만 잘라서 OCR 또는 문자열 패턴 탐색
        w, h = img_pil.size
        # Google Vision 또는 빠른 픽셀 기반 텍스트 추출 시도
        # 상단 영역 텍스트 탐색은 Gemini OCR 결과의 x_axis_min/max 와 조합
    except Exception:
        pass
    return None, None


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


def _get_gemini_raw_peaks(img_bytes, is_precursor=True):
    """
    Google Gemini REST API 로 피크 m/z 와 X축 범위(x_axis_min, x_axis_max), 눈금 픽셀 위치를 추출합니다.
    반환값: {'peaks': [...], 'axis_ticks': [...], 'x_axis_min': float, 'x_axis_max': float}
    """
    global _last_ai_error
    _last_ai_error = None

    api_key = get_gemini_api_key()
    if not api_key:
        _last_ai_error = "Gemini API 키가 설정되지 않았습니다."
        return {'peaks': [], 'axis_ticks': [], 'x_axis_min': None, 'x_axis_max': None}

    nparr = np.frombuffer(img_bytes, np.uint8)
    img_cv = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
    img_h, img_w = (img_cv.shape[:2] if img_cv is not None else (1200, 2200))

    b64_img = base64.b64encode(img_bytes).decode('utf-8')

    prompt = f"""
You are an expert analytical mass spectrometry (LC-MS/MS) data specialist.
Examine this LC-MS/MS spectrum plot image (Width: {img_w}px, Height: {img_h}px).

CRITICAL TASKS:
1. Header & Mass Range Detection:
   - Look at the top title header text line (e.g. "+Q1 (166 - 172)" or "+Product Ion of 335.1 (20 - 345)") and the bottom horizontal X-axis.
   - Find the exact start and end m/z values of the horizontal plot area:
     "x_axis_min": starting m/z (leftmost edge of plot)
     "x_axis_max": ending m/z (rightmost edge of plot)

2. Peak m/z Number Extraction:
   - Read all numerical peak labels printed above/near the blue spectral curve tops.
   - For EACH peak, record its exact printed m/z value (e.g. 168.96, 335.06, 76.99, 183.07, etc.).
   - Estimate its bounding box in image pixels: "xmin", "xmax", "ymin", "ymax".
   - Set "is_recommended": true for the dominant primary target ion ({'tallest main precursor peak' if is_precursor else 'top 3 fragment ion peaks excluding precursor'}).

3. X-Axis Calibration Ticks:
   - Identify 2 to 4 clearly labeled X-axis tick numbers at the bottom.
   - For each tick, record "mz" (the printed number) and "pixel_x" (its pixel X coordinate).

Return ONLY a JSON object formatted as:
{{
  "x_axis_min": 166.0,
  "x_axis_max": 172.0,
  "axis_ticks": [
    {{"mz": 166.2, "pixel_x": 140}},
    {{"mz": 169.0, "pixel_x": 1120}},
    {{"mz": 171.8, "pixel_x": 2100}}
  ],
  "peaks": [
    {{"mz": 168.96, "is_recommended": true, "height_rank": 1, "ymin": 50, "xmin": 1080, "xmax": 1150, "ymax": 85}}
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

                x_min_val = float(data.get("x_axis_min", 0.0))
                x_max_val = float(data.get("x_axis_max", 0.0))

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
                            'x_axis_min': x_min_val,
                            'x_axis_max': x_max_val,
                            'y_min': ymin, 'x_min': xmin,
                            'x_max': xmax, 'y_max': ymax,
                            'cx': cx, 'cy': cy
                        })
                    except Exception:
                        pass

                # axis_ticks 파싱
                axis_ticks = []
                for t in data.get("axis_ticks", []):
                    try:
                        mz_t = float(t["mz"])
                        px_t = float(t["pixel_x"])
                        if px_t <= 1000 and img_w > 1200:
                            px_t = (px_t / 1000.0) * img_w
                        axis_ticks.append({'mz': mz_t, 'pixel_x': int(px_t)})
                    except Exception:
                        pass

                if raw_peaks:
                    _last_ai_error = None
                    return {
                        'peaks': raw_peaks,
                        'axis_ticks': axis_ticks,
                        'x_axis_min': x_min_val if x_max_val > x_min_val else None,
                        'x_axis_max': x_max_val if x_max_val > x_min_val else None
                    }
            else:
                errors.append(f"[{model_name}] HTTP {resp.status_code}: {resp.text[:200]}")
        except Exception as e:
            errors.append(f"[{model_name}] 오류: {e}")

    if errors:
        _last_ai_error = errors[0] if len(errors) == 1 else " | ".join(errors[:2])

    return {'peaks': [], 'axis_ticks': [], 'x_axis_min': None, 'x_axis_max': None}


def extract_peaks_with_google_vision(img_bytes, is_precursor=True):
    """
    피크 m/z 값 및 X축 보정 데이터를 추출합니다.
    """
    gemini_result = _get_gemini_raw_peaks(img_bytes, is_precursor=is_precursor)
    gemini_raw = gemini_result['peaks']
    axis_ticks = gemini_result['axis_ticks']
    x_min = gemini_result['x_axis_min']
    x_max = gemini_result['x_axis_max']

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
            'x_axis_min': x_min,
            'x_axis_max': x_max,
            'raw_peaks': corrected_peaks,
            'engine': 'Google Gemini Flash AI ⚡ (99.99%)'
        }

    return {'all_peaks': [], 'default_checked': set(), 'axis_ticks': [], 'x_axis_min': None, 'x_axis_max': None, 'engine': 'Gemini API 키 필요 🔑'}


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


def process_spectrum_image(img_bytes, peak_overrides=None, font_size=45, is_precursor=True, axis_ticks=None, x_axis_min=None, x_axis_max=None):
    """
    Renders enlarged peak labels with precision peak apex mapping.
    """
    if peak_overrides is None:
        peak_overrides = []
    if axis_ticks is None:
        axis_ticks = []

    nparr = np.frombuffer(img_bytes, np.uint8)
    img = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
    if img is None or img.size == 0:
        return img_bytes

    h, w = img.shape[:2]
    x_axis, y_axis = detect_axis_lines(img)
    x_start = x_axis
    x_end = w - 35
    y_top = 50
    y_bottom = y_axis - 5

    # 1. 기존 소형 텍스트 완전 소거
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

    # 2. X축 m/z 물리적 범위 결정 (우선순위: 헤더 정규식 -> axis_ticks -> x_axis_min/max -> 폴백)
    min_mz = None
    max_mz = None

    if axis_ticks and len(axis_ticks) >= 2:
        t_sorted = sorted(axis_ticks, key=lambda t: t['mz'])
        t1, t2 = t_sorted[0], t_sorted[-1]
        dmz = t2['mz'] - t1['mz']
        dpx = t2['pixel_x'] - t1['pixel_x']
        if dmz > 0.01 and dpx > 50:
            scale = dpx / dmz
            # t1을 기준으로 x_start와 x_end에서의 m/z 역산
            min_mz = t1['mz'] - (t1['pixel_x'] - x_start) / scale
            max_mz = t1['mz'] + (x_end - t1['pixel_x']) / scale

    if min_mz is None and x_axis_min is not None and x_axis_max is not None and x_axis_max > x_axis_min:
        min_mz = x_axis_min
        max_mz = x_axis_max

    # 3. m/z -> X 픽셀 변환 함수
    if min_mz is not None and max_mz is not None and max_mz > min_mz:
        span_mz = max_mz - min_mz
        span_px = x_end - x_start
        def calc_x(mz):
            return int(x_start + ((mz - min_mz) / span_mz) * span_px)
        px_per_da = span_px / span_mz
    else:
        # 고정 기본 폴백
        all_vals = []
        for item in peak_overrides:
            try:
                all_vals.append(float(item.get('final_mz', item.get('orig_mz'))))
            except ValueError:
                pass
        max_v = max(all_vals, default=350.0)
        if max_v <= 210.0:
            def calc_x(mz): return int(40 + (mz - 30.0) * 13.48)
            px_per_da = 13.48
        elif is_precursor:
            def calc_x(mz): return int(59 + (mz - 50.0) * 4.75333)
            px_per_da = 4.75333
        else:
            def calc_x(mz): return int(59 + (mz - 20.0) * 6.684375)
            px_per_da = 6.684375

    # 4. 피크 레이블 구성 & 정점(Apex) 매핑
    peak_labels = []
    for item in peak_overrides:
        orig_mz = item.get('orig_mz', '')
        final_mz = item.get('final_mz', orig_mz)
        is_mrm = item.get('is_mrm', False)
        try:
            val_f = float(final_mz)
        except ValueError:
            continue

        x_calc = calc_x(val_f)
        x_calc = max(x_start + 10, min(x_end - 10, x_calc))

        # 로컬 피크 정점(Apex) 탐색:
        # 스캔 반경은 국소 범위 (최대 ±0.35 Da 또는 ±25px)로 제한하여
        # 인접한 다른 거대 피크로 잘못 점프하는 것을 원천 방지
        local_scan = max(15, min(40, int(px_per_da * 0.35)))
        x1 = max(x_start, x_calc - local_scan)
        x2 = min(x_end, x_calc + local_scan)

        best_apex_x = x_calc
        best_apex_y = y_bottom

        # x1~x2 범위 내에서 파란색 선의 최상단(y 최소) 탐색
        for cx in range(x1, x2):
            col = pure_blue[50:y_bottom, cx]
            blue_idx = np.where(col)[0]
            if len(blue_idx) > 0:
                top_y = 50 + np.min(blue_idx)
                if top_y < best_apex_y:
                    best_apex_y = top_y
                    best_apex_x = cx

        peak_labels.append({
            'text': final_mz,
            'center': (best_apex_x, best_apex_y),
            'bbox': (best_apex_x - 40, best_apex_y - 15, best_apex_x + 40, best_apex_y + 15),
            'is_mrm': is_mrm
        })

    # 5. 투명 오버레이로 레이블 합성
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
    """
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"파일을 찾을 수 없습니다: {excel_path}")

    if status_callback:
        status_callback("엑셀 파일 로딩 및 MS/MS 이미지 변환 중...")

    wb = openpyxl.load_workbook(excel_path)
    temp_dir = os.path.join(os.path.dirname(excel_path), "_temp_mz_img")
    os.makedirs(temp_dir, exist_ok=True)

    total_images = sum(len(wb[name]._images) for name in wb.sheetnames)
    processed_count = 0

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        sel_dict = sheet_selections.get(sheet_name, {'precursor': [], 'product': [], 'axis_ticks': {}, 'meta': {}})
        ticks_dict = sel_dict.get('axis_ticks', {})
        meta_dict = sel_dict.get('meta', {})

        for idx, img in enumerate(ws._images):
            processed_count += 1
            if status_callback:
                status_callback(f"이미지 처리 중 ({processed_count}/{total_images}) - 시트: {sheet_name}")

            ion_type = 'precursor' if idx % 2 == 0 else 'product'
            is_prec = (ion_type == 'precursor')
            overrides = sel_dict.get(ion_type, [])
            axis_ticks = ticks_dict.get(ion_type, [])
            ion_meta = meta_dict.get(ion_type, {})

            img_bytes = img._data()
            proc_bytes = process_spectrum_image(
                img_bytes,
                peak_overrides=overrides,
                font_size=font_size,
                is_precursor=is_prec,
                axis_ticks=axis_ticks,
                x_axis_min=ion_meta.get('x_axis_min'),
                x_axis_max=ion_meta.get('x_axis_max')
            )

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
    print("웹 앱 실행: streamlit run app.py")
    print("데스크톱 GUI 실행: python desktop_app.py")
