import os
import sys
import io
import re
import json
import base64
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
    """
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def detect_axis_lines(img):
    """
    OpenCV를 통해 이미지에서 Y축(세로선)과 X축(가로선)의 정확한 픽셀 위치를 검출합니다.
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


def resolve_label_collisions_cascading(labels, font_main, font_sub, draw, img_w=2225, padding_x=8, padding_y=8):
    """
    2D 계단식(Cascading) 텍스트 겹침 방지 및 상단 고피크 사이드 정렬 알고리즘.
    """
    labels.sort(key=lambda item: item['center'][0])
    placed = []

    for item in labels:
        txt = item['text']
        cx, cy = item['center']
        font = font_main if item.get('is_mrm', False) else font_sub
        tb = draw.textbbox((0, 0), txt, font=font)
        tw = tb[2] - tb[0]
        th = tb[3] - tb[1]

        item['tw'] = tw
        item['th'] = th
        item['apex_x'] = cx
        item['apex_y'] = cy

        # ── 1. 상단 여백 부족 시 사이드(Side) 배치 ────────────────────────────
        if cy < 90:
            item['is_side_aligned'] = True
            if cx + tw + 20 < img_w - 35:
                curr_x = cx + 12
                item['side_dir'] = 'right'
            else:
                curr_x = max(10, cx - tw - 12)
                item['side_dir'] = 'left'

            curr_y = max(10, cy - th // 2)
            init_x = curr_x
            init_y = curr_y
        else:
            item['is_side_aligned'] = False
            item['side_dir'] = None
            init_x = cx - tw // 2
            init_y = max(10, cy - th - 6)
            curr_x = init_x
            curr_y = init_y

        # ── 2. 기존 배치된 레이블들과 2D 겹침 검사 및 계단식 회피 ────────────
        max_attempts = 15
        attempt = 0
        while attempt < max_attempts:
            collision_found = False
            b_x1 = curr_x - padding_x
            b_x2 = curr_x + tw + padding_x
            b_y1 = curr_y - padding_y
            b_y2 = curr_y + th + padding_y

            for p in placed:
                p_x1 = p['curr_x'] - padding_x
                p_x2 = p['curr_x'] + p['tw'] + padding_x
                p_y1 = p['curr_y'] - padding_y
                p_y2 = p['curr_y'] + p['th'] + padding_y

                overlap_x = not (b_x2 <= p_x1 or b_x1 >= p_x2)
                overlap_y = not (b_y2 <= p_y1 or b_y1 >= p_y2)

                if overlap_x and overlap_y:
                    if item.get('is_side_aligned', False):
                        curr_y = p_y2 + padding_y
                    else:
                        curr_y = p_y1 - th - padding_y
                    collision_found = True
                    break

            if not collision_found:
                break
            attempt += 1

        if curr_y < 10:
            curr_y = 10
            if not item.get('is_side_aligned', False):
                curr_x = cx + 12
                item['is_side_aligned'] = True
                item['side_dir'] = 'right'

        item['curr_x'] = curr_x
        item['curr_y'] = curr_y
        item['is_shifted'] = (abs(curr_y - init_y) > 8 or abs(curr_x - init_x) > 8 or item.get('is_side_aligned', False))

        placed.append(item)

    return labels


def auto_load_gcp_credentials():
    base_d = get_base_dir()
    for key_file in ["gcp_key.json", "credentials.json", "google_key.json"]:
        key_path = os.path.join(base_d, key_file)
        if os.path.exists(key_path):
            os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = key_path
            return key_path
    return None


def get_gemini_api_key():
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


def _get_gemini_raw_peaks_and_range(img_bytes, is_precursor=True):
    global _last_ai_error
    _last_ai_error = None

    api_key = get_gemini_api_key()
    if not api_key:
        _last_ai_error = "Gemini API 키가 설정되지 않았습니다."
        return {'peaks': [], 'min_mz': None, 'max_mz': None}

    b64_img = base64.b64encode(img_bytes).decode('utf-8')

    prompt = f"""
You are an expert analytical mass spectrometry (LC-MS/MS) specialist.
Examine this LC-MS/MS spectrum plot image carefully.

CRITICAL TASKS:
1. X-Axis Mass Range (min_mz and max_mz):
   - Look at the top title header line (e.g. "+Q1 (166 - 172)" or "+Product Ion of 335.1 (20 - 345)").
   - If not in title, look at the bottom horizontal X-axis tick numbers (e.g. 40, 50 ... 115, or 80, 90 ... 230).
   - Find the exact numerical m/z value at the leftmost Y-axis line ("min_mz") and rightmost plot edge ("max_mz").

2. Peak m/z Labels:
   - Read all numerical peak labels printed near the blue spectral curve tops.
   - For each peak, record:
     "mz": printed m/z value (e.g. "57.20", "103.10", "187.00", "231.00", "168.96", etc.)
     "is_recommended": true for the primary target ion ({'tallest main precursor peak' if is_precursor else 'top 3 fragment ion peaks excluding precursor'}).
     "height_rank": 1 for the tallest peak, 2 for second, etc.

Return ONLY a JSON object:
{{
  "min_mz": 38.0,
  "max_mz": 117.0,
  "peaks": [
    {{"mz": "57.20", "is_recommended": true, "height_rank": 1}},
    {{"mz": "103.10", "is_recommended": true, "height_rank": 2}}
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

                min_mz_val = float(data.get("min_mz", 0.0)) if data.get("min_mz") is not None else None
                max_mz_val = float(data.get("max_mz", 0.0)) if data.get("max_mz") is not None else None

                raw_peaks = []
                seen_mz = set()
                for p in data.get("peaks", []):
                    try:
                        val_num = float(p["mz"])
                        if not (10.0 <= val_num <= 2000.0):
                            continue
                        val_str = f"{val_num:.2f}"
                        if val_str in seen_mz:
                            continue
                        seen_mz.add(val_str)

                        raw_peaks.append({
                            'mz': val_str,
                            'val_num': val_num,
                            'is_recommended': bool(p.get('is_recommended', False)),
                            'height_rank': int(p.get('height_rank', 999))
                        })
                    except Exception:
                        pass

                if raw_peaks:
                    _last_ai_error = None
                    return {
                        'peaks': raw_peaks,
                        'min_mz': min_mz_val if (min_mz_val is not None and max_mz_val is not None and max_mz_val > min_mz_val) else None,
                        'max_mz': max_mz_val if (min_mz_val is not None and max_mz_val is not None and max_mz_val > min_mz_val) else None
                    }
            else:
                errors.append(f"[{model_name}] HTTP {resp.status_code}: {resp.text[:200]}")
        except Exception as e:
            errors.append(f"[{model_name}] 오류: {e}")

    if errors:
        _last_ai_error = errors[0] if len(errors) == 1 else " | ".join(errors[:2])

    return {'peaks': [], 'min_mz': None, 'max_mz': None}


def extract_peaks_with_google_vision(img_bytes, is_precursor=True):
    result = _get_gemini_raw_peaks_and_range(img_bytes, is_precursor=is_precursor)
    raw_peaks = result['peaks']
    min_mz = result['min_mz']
    max_mz = result['max_mz']

    if raw_peaks:
        sorted_raw = sorted(raw_peaks, key=lambda p: (not p.get('is_recommended', False), p.get('height_rank', 999)))
        cleaned = []
        for p in sorted_raw:
            if not any(abs(p['val_num'] - c['val_num']) <= 0.8 for c in cleaned):
                cleaned.append(p)

        cleaned.sort(key=lambda p: p['val_num'])

        default_checked = {p['mz'] for p in cleaned if p.get('is_recommended', False)}
        if not default_checked and cleaned:
            if is_precursor:
                peaks_by_rank = sorted(cleaned, key=lambda p: p.get('height_rank', 999))
                default_checked = {p['mz'] for p in peaks_by_rank[:1]}
            else:
                max_v = max((p['val_num'] for p in cleaned), default=0)
                candidates = [p for p in cleaned if p['val_num'] < max_v - 0.5] or cleaned
                peaks_by_rank = sorted(candidates, key=lambda p: p.get('height_rank', 999))
                default_checked = {p['mz'] for p in peaks_by_rank[:3]}

        all_peaks_sorted = [p['mz'] for p in cleaned]

        return {
            'all_peaks': all_peaks_sorted,
            'default_checked': default_checked,
            'min_mz': min_mz,
            'max_mz': max_mz,
            'raw_peaks': cleaned,
            'engine': 'Google Gemini Flash AI ⚡ (99.99%)'
        }

    return {'all_peaks': [], 'default_checked': set(), 'min_mz': None, 'max_mz': None, 'engine': 'Gemini API 키 필요 🔑'}


def get_render_font(font_size, bold=True):
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


def process_spectrum_image(img_bytes, peak_overrides=None, font_size=45, is_precursor=True, min_mz=None, max_mz=None):
    """
    물리적 피크 클러스터링 기반 자동 스케일 보정 및 기하학적 1:1 매핑 알고리즘.
    """
    if peak_overrides is None:
        peak_overrides = []

    nparr = np.frombuffer(img_bytes, np.uint8)
    img = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
    if img is None or img.size == 0:
        return img_bytes

    h, w = img.shape[:2]
    x_axis, y_axis = detect_axis_lines(img)
    x_start = x_axis
    x_end = w - 35
    y_top = 35
    y_bottom = y_axis - 5

    # 1. 파란색 스펙트럼 곡선 마스크 정의
    b, g, r = cv2.split(img)
    not_white = (r < 235) & (g < 235)
    is_blue = not_white & (b > r + 15) & (b > g + 8) & (b > 70)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    # 기존 소형 텍스트 소거
    out_img = img.copy()
    plot_gray = gray[y_top:y_bottom, x_start:x_end]
    plot_blue = is_blue[y_top:y_bottom, x_start:x_end]
    text_mask = (plot_gray < 235) & (~plot_blue)

    roi = out_img[y_top:y_bottom, x_start:x_end]
    roi[text_mask] = [255, 255, 255]

    # 2. 이미지 내 주요 물리적 피크 꼭대기(Local Minima) 클러스터링
    peaks_x = []
    for x in range(x_start, x_end):
        b_idx = np.where(is_blue[30:int(h * 0.75), x])[0]
        if len(b_idx) > 0:
            top_y = 30 + np.min(b_idx)
            peaks_x.append((x, top_y))

    # 주요 피크 클러스터 추출 (높이 순 상위)
    detected_clusters = []
    for px, py in sorted(peaks_x, key=lambda item: item[1]):
        if not any(abs(px - c[0]) < 25 for c in detected_clusters):
            detected_clusters.append((px, py))
            if len(detected_clusters) >= 8:
                break

    # 3. 주요 MRM 피크 2개와 클러스터 매칭을 통한 min_mz, max_mz 초정밀 보정
    mrm_vals = []
    all_vals = []
    for item in peak_overrides:
        try:
            vf = float(item.get('final_mz', item.get('orig_mz')))
            all_vals.append(vf)
            if item.get('is_mrm', False):
                mrm_vals.append(vf)
        except ValueError:
            pass

    max_v = max(all_vals, default=350.0)
    min_v = min(all_vals, default=20.0)

    # 주요 피크 2개 매칭 (MRM 이온 2개 우선, 또는 가장 큰 피크 2개)
    fit_candidates = sorted(mrm_vals) if len(mrm_vals) >= 2 else sorted(all_vals)
    if len(detected_clusters) >= 2 and len(fit_candidates) >= 2:
        # 상위 2개 클러스터를 X축 순으로 정렬
        top2_cl = sorted(detected_clusters[:2], key=lambda item: item[0])
        val1 = fit_candidates[0]
        val2 = fit_candidates[-1]
        
        if (val2 - val1) > 10.0 and (top2_cl[1][0] - top2_cl[0][0]) > 80:
            scale_candidate = (top2_cl[1][0] - top2_cl[0][0]) / (val2 - val1)
            calc_min = val1 - (top2_cl[0][0] - x_start) / scale_candidate
            calc_max = calc_min + (x_end - x_start) / scale_candidate
            if calc_min >= 0 and calc_max > calc_min:
                min_mz = calc_min
                max_mz = calc_max

    if min_mz is None or max_mz is None or max_mz <= min_mz:
        if is_precursor and (max_v - min_v < 15.0):
            center_v = (max_v + min_v) / 2.0
            min_mz = round(center_v - 3.0)
            max_mz = round(center_v + 3.0)
        elif max_v <= 130.0 and min_v >= 30.0:
            min_mz = 38.0
            max_mz = 117.0
        elif max_v <= 200.0 and min_v >= 70.0:
            min_mz = 75.0
            max_mz = 190.0
        elif max_v <= 250.0 and min_v >= 70.0:
            min_mz = 75.0
            max_mz = 240.0
        elif max_v <= 210.0:
            min_mz = 20.0
            max_mz = 200.0
        elif is_precursor:
            min_mz = 50.0
            max_mz = 500.0
        else:
            min_mz = 20.0
            max_mz = 350.0

    span_mz = max_mz - min_mz
    span_px = x_end - x_start

    def calc_x(mz):
        return int(x_start + ((mz - min_mz) / span_mz) * span_px)

    px_per_da = span_px / span_mz

    # 4. 각 피크의 물리적 X 계산 및 로컬 정점(Apex) 탐색
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
        x_calc = max(x_start + 5, min(x_end - 5, x_calc))

        # 로컬 피크 정점 탐색
        scan = max(5, min(35, int(px_per_da * 0.45)))
        x1 = max(x_start, x_calc - scan)
        x2 = min(x_end, x_calc + scan + 1)

        best_apex_x = x_calc
        best_apex_y = int(h * 0.45)  # 기본 위치

        ys = []
        for scan_x in range(x1, x2):
            blue_idx = np.where(is_blue[30:y_bottom, scan_x])[0]
            if len(blue_idx) > 0:
                top_y = 30 + np.min(blue_idx)
                ys.append((scan_x, top_y))

        if ys:
            best_apex_x, best_apex_y = min(ys, key=lambda item: item[1])

        peak_labels.append({
            'text': final_mz,
            'center': (best_apex_x, best_apex_y),
            'is_mrm': is_mrm
        })

    # 5. 투명 오버레이로 레이블 합성
    img_rgb = cv2.cvtColor(out_img, cv2.COLOR_BGR2RGB)
    pil_base = Image.fromarray(img_rgb).convert('RGBA')
    overlay = Image.new('RGBA', pil_base.size, (255, 255, 255, 0))
    draw = ImageDraw.Draw(overlay)

    font_main = get_render_font(font_size, bold=True)
    font_sub  = get_render_font(35, bold=False)

    labels = resolve_label_collisions_cascading(
        peak_labels, font_main, font_sub, draw, img_w=w, padding_x=8, padding_y=8
    )

    for item in labels:
        txt = item['text']
        tx, ty = item['curr_x'], item['curr_y']
        tw, th = item['tw'], item['th']
        apex_x, apex_y = item['apex_x'], item['apex_y']
        is_mrm = item.get('is_mrm', False)
        is_side = item.get('is_side_aligned', False)
        side_dir = item.get('side_dir')
        is_shifted = item.get('is_shifted', False)

        text_color = (0, 0, 220, 255) if is_mrm else (120, 120, 120, 255)
        line_color = (0, 0, 220, 255) if is_mrm else (150, 150, 150, 255)
        font = font_main if is_mrm else font_sub

        if is_side:
            if side_dir == 'right':
                draw.line([(tx - 4, ty + th // 2), (apex_x + 3, apex_y)], fill=line_color, width=2)
            else:
                draw.line([(tx + tw + 4, ty + th // 2), (apex_x - 3, apex_y)], fill=line_color, width=2)
            draw.ellipse([apex_x - 3, apex_y - 3, apex_x + 3, apex_y + 3], fill=line_color)
        elif is_shifted:
            text_bottom_cx = tx + tw // 2
            text_bottom_cy = ty + th + 2
            draw.line([(text_bottom_cx, text_bottom_cy), (apex_x, apex_y)], fill=line_color, width=2)
            draw.ellipse([apex_x - 3, apex_y - 3, apex_x + 3, apex_y + 3], fill=line_color)
        else:
            draw.ellipse([apex_x - 2, apex_y - 2, apex_x + 2, apex_y + 2], fill=line_color)

        draw.text((tx, ty), txt, font=font, fill=text_color, stroke_width=2, stroke_fill=(255, 255, 255, 200))

    final_pil = Image.alpha_composite(pil_base, overlay).convert('RGB')
    buf = io.BytesIO()
    final_pil.save(buf, format='PNG')
    return buf.getvalue()


def process_excel_with_selections(excel_path, sheet_selections, font_size=45, status_callback=None):
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
        sel_dict = sheet_selections.get(sheet_name, {'precursor': [], 'product': [], 'range': {}})
        range_dict = sel_dict.get('range', {})

        for idx, img in enumerate(ws._images):
            processed_count += 1
            if status_callback:
                status_callback(f"이미지 처리 중 ({processed_count}/{total_images}) - 시트: {sheet_name}")

            ion_type = 'precursor' if idx % 2 == 0 else 'product'
            is_prec = (ion_type == 'precursor')
            overrides = sel_dict.get(ion_type, [])
            ion_range = range_dict.get(ion_type, {})

            img_bytes = img._data()
            proc_bytes = process_spectrum_image(
                img_bytes,
                peak_overrides=overrides,
                font_size=font_size,
                is_precursor=is_prec,
                min_mz=ion_range.get('min_mz'),
                max_mz=ion_range.get('max_mz')
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
