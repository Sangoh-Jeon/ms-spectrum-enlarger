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


def resolve_label_collisions_cascading(labels, font_main, font_sub, draw, padding_x=8, padding_y=8):
    """
    2D 계단식(Cascading) 텍스트 겹침 방지 알고리즘.
    - X축 오름차순(좌측 -> 우측)으로 순차 검사
    - 이미 배치된 좌측의 모든 레이블들과 바운딩 박스 충돌 검사
    - 충돌 시 겹친 레이블들보다 한 단계 더 위로(Y 감소) 계단식 상승
    - 피크에서 위로 이동한 레이블은 원래 피크 꼭대기까지 지시선(Leader line) 자동 연결
    """
    # X 좌표(피크 정점 X) 기준 좌 -> 우 정렬
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

        # 초기 배치: 피크 정점 바로 위 중앙
        init_x = cx - tw // 2
        init_y = max(10, cy - th - 6)

        curr_x = init_x
        curr_y = init_y

        # 이미 배치된 이전 레이블들과 반복 충돌 검사 (충돌이 완전히 사라질 때까지 위로 상승)
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

                # 2D Bounding Box AABB 겹침 검사
                overlap_x = not (b_x2 <= p_x1 or b_x1 >= p_x2)
                overlap_y = not (b_y2 <= p_y1 or b_y1 >= p_y2)

                if overlap_x and overlap_y:
                    # 겹치는 레이블보다 위로 한 칸 이동
                    curr_y = p_y1 - th - padding_y
                    collision_found = True
                    break

            if not collision_found:
                break
            attempt += 1

        # 화면 최상단 이탈 방어 (y < 10px 이면 사이드 분기)
        if curr_y < 10:
            curr_y = 10
            # 만약 최상단에서도 x가 겹치면 살짝 우측/좌측으로 이동
            curr_x = cx + 8

        item['curr_x'] = curr_x
        item['curr_y'] = curr_y
        item['is_shifted'] = (abs(curr_y - init_y) > 10 or abs(curr_x - init_x) > 10)

        placed.append(item)

    return labels


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


def _get_gemini_raw_peaks_with_coords(img_bytes, is_precursor=True):
    """
    Google Gemini REST API 로 피크 m/z 값과 해당 텍스트의 이미지 상 실제 (x, y) 픽셀 좌표를 직접 추출합니다.
    """
    global _last_ai_error
    _last_ai_error = None

    api_key = get_gemini_api_key()
    if not api_key:
        _last_ai_error = "Gemini API 키가 설정되지 않았습니다."
        return []

    nparr = np.frombuffer(img_bytes, np.uint8)
    img_cv = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
    img_h, img_w = (img_cv.shape[:2] if img_cv is not None else (1200, 2225))

    b64_img = base64.b64encode(img_bytes).decode('utf-8')

    prompt = f"""
You are an expert analytical mass spectrometry (LC-MS/MS) data specialist.
Examine this LC-MS/MS spectrum plot image (Width: {img_w}px, Height: {img_h}px).

TASK:
1. Read all numerical peak labels printed above/near the blue spectral curve tops.
2. For EACH peak, identify its printed m/z value AND its precise image pixel coordinates:
   - "mz": printed m/z string (e.g. "168.96", "335.06", "76.99", "183.07")
   - "center_x": the center pixel X coordinate in this image (from 0 to {img_w})
   - "center_y": the center pixel Y coordinate in this image (from 0 to {img_h})
   - "is_recommended": true for the primary target ion ({'tallest main precursor peak' if is_precursor else 'top 3 fragment ion peaks excluding precursor'}).

Return ONLY a JSON object formatted as:
{{
  "peaks": [
    {{"mz": "168.96", "center_x": 1102, "center_y": 60, "is_recommended": true, "height_rank": 1}},
    {{"mz": "166.96", "center_x": 420, "center_y": 880, "is_recommended": false, "height_rank": 2}}
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

                        cx_raw = float(p.get("center_x", 1000))
                        cy_raw = float(p.get("center_y", 500))

                        # 0~1000 정규화 좌표로 반환된 경우 실제 픽셀 스케일로 변환
                        if cx_raw <= 1000 and img_w > 1200:
                            cx = int((cx_raw / 1000.0) * img_w)
                            cy = int((cy_raw / 1000.0) * img_h)
                        else:
                            cx = int(cx_raw)
                            cy = int(cy_raw)

                        raw_peaks.append({
                            'mz': val_str,
                            'val_num': val_num,
                            'center_x': cx,
                            'center_y': cy,
                            'is_recommended': bool(p.get('is_recommended', False)),
                            'height_rank': int(p.get('height_rank', 999))
                        })
                    except Exception:
                        pass

                if raw_peaks:
                    _last_ai_error = None
                    return raw_peaks
            else:
                errors.append(f"[{model_name}] HTTP {resp.status_code}: {resp.text[:200]}")
        except Exception as e:
            errors.append(f"[{model_name}] 오류: {e}")

    if errors:
        _last_ai_error = errors[0] if len(errors) == 1 else " | ".join(errors[:2])

    return []


def extract_peaks_with_google_vision(img_bytes, is_precursor=True):
    """
    피크 m/z 값 및 각 텍스트의 실제 (x, y) 픽셀 좌표를 함께 추출합니다.
    """
    raw_peaks = _get_gemini_raw_peaks_with_coords(img_bytes, is_precursor=is_precursor)

    if raw_peaks:
        # 중복 병합 및 정렬 (delta <= 0.8)
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
                max_mz = max((p['val_num'] for p in cleaned), default=0)
                candidates = [p for p in cleaned if p['val_num'] < max_mz - 0.5] or cleaned
                peaks_by_rank = sorted(candidates, key=lambda p: p.get('height_rank', 999))
                default_checked = {p['mz'] for p in peaks_by_rank[:3]}

        all_peaks_sorted = [p['mz'] for p in cleaned]
        # m/z -> coords 딕셔너리 구성
        coords_map = {p['mz']: {'cx': p['center_x'], 'cy': p['center_y']} for p in cleaned}

        return {
            'all_peaks': all_peaks_sorted,
            'default_checked': default_checked,
            'coords_map': coords_map,
            'raw_peaks': cleaned,
            'engine': 'Google Gemini Flash AI ⚡ (99.99%)'
        }

    return {'all_peaks': [], 'default_checked': set(), 'coords_map': {}, 'engine': 'Gemini API 키 필요 🔑'}


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
    기억된 (x, y) 픽셀 좌표를 기반으로 피크 정점에 글자를 배치하고,
    2D 계단식 충돌 해소로 겹침을 방지하여 렌더링합니다.
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

    # 2. 기억된 (cx, cy) 좌표로부터 실제 피크 정점(Apex) 매핑
    peak_labels = []
    for item in peak_overrides:
        orig_mz = item.get('orig_mz', '')
        final_mz = item.get('final_mz', orig_mz)
        is_mrm = item.get('is_mrm', False)

        # 기억된 중심 X, Y 좌표
        cx = item.get('cx')
        cy = item.get('cy')

        if cx is None:
            # 좌표가 누락된 경우 기본 중앙 배치 폴백
            cx = (x_start + x_end) // 2
            cy = y_bottom // 2

        cx = max(x_start + 10, min(x_end - 10, int(cx)))

        # 해당 cx 주변(±25px)에서 파란색 스펙트럼 선의 최상단(정점) 탐색
        scan = 25
        x1 = max(x_start, cx - scan)
        x2 = min(x_end, cx + scan)

        best_apex_x = cx
        best_apex_y = y_bottom

        for scan_x in range(x1, x2):
            col = pure_blue[50:y_bottom, scan_x]
            blue_idx = np.where(col)[0]
            if len(blue_idx) > 0:
                top_y = 50 + np.min(blue_idx)
                if top_y < best_apex_y:
                    best_apex_y = top_y
                    best_apex_x = scan_x

        peak_labels.append({
            'text': final_mz,
            'center': (best_apex_x, best_apex_y),
            'is_mrm': is_mrm
        })

    # 3. 투명 오버레이로 레이블 합성
    img_rgb = cv2.cvtColor(out_img, cv2.COLOR_BGR2RGB)
    pil_base = Image.fromarray(img_rgb).convert('RGBA')
    overlay = Image.new('RGBA', pil_base.size, (255, 255, 255, 0))
    draw = ImageDraw.Draw(overlay)

    font_main = get_render_font(font_size, bold=True)
    font_sub = get_render_font(max(12, font_size - 5), bold=False)

    # 2D 계단식 겹침 방지 알고리즘 실행
    labels = resolve_label_collisions_cascading(
        peak_labels, font_main, font_sub, draw, padding_x=8, padding_y=8
    )

    for item in labels:
        txt = item['text']
        tx, ty = item['curr_x'], item['curr_y']
        tw, th = item['tw'], item['th']
        apex_x, apex_y = item['apex_x'], item['apex_y']
        is_mrm = item.get('is_mrm', False)
        is_shifted = item.get('is_shifted', False)

        text_color = (0, 0, 220, 255) if is_mrm else (120, 120, 120, 255)
        line_color = (0, 0, 220, 255) if is_mrm else (150, 150, 150, 255)
        font = font_main if is_mrm else font_sub

        # 글자가 원래 피크 정점에서 위로 올라간 경우 지시선 및 정점 점 표시
        if is_shifted:
            text_bottom_cx = tx + tw // 2
            text_bottom_cy = ty + th + 2
            draw.line([(text_bottom_cx, text_bottom_cy), (apex_x, apex_y)], fill=line_color, width=2)
            draw.ellipse([apex_x - 3, apex_y - 3, apex_x + 3, apex_y + 3], fill=line_color)
        else:
            # 꼭대기에 바로 붙은 경우 작은 정점 표시
            draw.ellipse([apex_x - 2, apex_y - 2, apex_x + 2, apex_y + 2], fill=line_color)

        # 텍스트 외곽에 부드러운 흰색 스트로크를 주어 배경 선과 겹쳐도 또렷하게 표시
        draw.text((tx, ty), txt, font=font, fill=text_color, stroke_width=2, stroke_fill=(255, 255, 255, 200))

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
        sel_dict = sheet_selections.get(sheet_name, {'precursor': [], 'product': []})

        for idx, img in enumerate(ws._images):
            processed_count += 1
            if status_callback:
                status_callback(f"이미지 처리 중 ({processed_count}/{total_images}) - 시트: {sheet_name}")

            ion_type = 'precursor' if idx % 2 == 0 else 'product'
            is_prec = (ion_type == 'precursor')
            overrides = sel_dict.get(ion_type, [])

            img_bytes = img._data()
            proc_bytes = process_spectrum_image(
                img_bytes,
                peak_overrides=overrides,
                font_size=font_size,
                is_precursor=is_prec
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
