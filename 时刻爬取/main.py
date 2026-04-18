import os
import re
import requests
import pandas as pd
from io import BytesIO
from PIL import Image
import pytesseract
import platform
from datetime import datetime, timedelta
import cv2
import numpy as np
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font

# 1. Tesseract 路径配置
if platform.system() == "Windows":
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# 2. 目录配置
base_dir = "."
image_dir = os.path.join(base_dir, "data", "raw_images")
excel_dir = os.path.join(base_dir, "data", "metadata")
excel_path = os.path.join(excel_dir, "表.xlsx")

os.makedirs(image_dir, exist_ok=True)
os.makedirs(excel_dir, exist_ok=True)

# 3. 读取现有表格 (用于防重复检查)
existing_times = set()
if os.path.exists(excel_path):
    try:
        df = pd.read_excel(excel_path)
        if '时间' in df.columns:
            existing_times = set(df['时间'].astype(str).tolist())
    except Exception as e:
        print(f"读取旧表格失败: {e}")

def get_real_time_from_image(img_bytes):
    """裁剪图片右上角并使用 OpenCV 增强识别时间"""
    try:
        img = Image.open(BytesIO(img_bytes))
        width, height = img.size
        crop_box = (width - 600, 0, width, 150)
        cropped_img = img.crop(crop_box)

        img_array = np.array(cropped_img)
        if len(img_array.shape) == 3:
            gray = cv2.cvtColor(img_array, cv2.COLOR_RGB2GRAY)
        else:
            gray = img_array
            
        _, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY_INV)

        custom_config = r'--oem 3 --psm 6'
        text = pytesseract.image_to_string(thresh, lang='eng', config=custom_config)
        print(f"🤖 AI 看到的原始文字是: 【{text.strip()}】")

        date_match = re.search(r'(\d{1,2}[-/]\d{1,2})', text)
        time_match = re.search(r'(\d{2}:\d{2}:\d{2})', text)

        if date_match and time_match:
            raw_date = date_match.group(1).replace('/', '-')
            try:
                parsed_date = datetime.strptime(raw_date, "%m-%d")
                month_day = parsed_date.strftime("%m-%d")
            except:
                month_day = raw_date 

            hour_minute_second = time_match.group(1)
            bjt_now = datetime.utcnow() + timedelta(hours=8)
            current_year = bjt_now.year

            return f"{current_year}-{month_day} {hour_minute_second}"
    except Exception as e:
        print(f"OCR 识别出错: {e}")
    return None

# 4. 爬取目标
targets = [
    {"url": "http://view.iap.ac.cn:8080/imageview/northeast.jpg", "direction": "northeast"},
    {"url": "http://view.iap.ac.cn:8080/imageview/southwest.jpg", "direction": "southwest"}
]

new_data = []

for target in targets:
    url = target["url"]
    direction = target["direction"]
    print(f"\n正在访问网页: {url}")

    try:
        response = requests.get(url, timeout=10)
        response.raise_for_status()
        img_bytes = response.content

        print("正在识别图片时间...")
        extracted_time = get_real_time_from_image(img_bytes)
        bjt_now = datetime.utcnow() + timedelta(hours=8)

        # 智能判断：时间来源与有效性
        if extracted_time:
            final_time = extracted_time
            time_source = 'OCR'
            is_valid = 1
        else:
            final_time = bjt_now.strftime('%Y-%m-%d %H:%M:%S')
            time_source = 'System'
            is_valid = 0
            print(f"⚠️ 识别失败，使用当前北京时间兜底: {final_time}")

        if final_time in existing_times:
            print(f"🛑 发现重复数据！时间 {final_time} 已存在，跳过 {direction} 的保存。")
            continue

        # 智能计算：延迟分钟数 (time_diff_min) 和 白天黑夜 (is_daytime)
        try:
            final_dt = datetime.strptime(final_time, '%Y-%m-%d %H:%M:%S')
            # 计算延迟：当前北京时间 - 图片上的时间
            diff_seconds = (bjt_now - final_dt).total_seconds()
            time_diff_min = round(diff_seconds / 60)
            
            # 判断白天：6点(含)到18点(不含)视为白天
            is_daytime = 1 if 6 <= final_dt.hour < 18 else 0
        except Exception as e:
            time_diff_min = ''
            is_daytime = ''
            print(f"时间计算出错: {e}")

        safe_time_str = final_time.replace(':', '-')
        image_filename = f"{direction}_{safe_time_str}.jpg"
        image_path = os.path.join(image_dir, image_filename)

        with open(image_path, 'wb') as f:
            f.write(img_bytes)
        print(f"✅ 成功保存新图片: {image_filename}")

        # 填入 12 个字段 (能算的都算好，不能算的留空)
        new_data.append({
            '时间': final_time,
            '方向': direction,
            '图片路径': image_path,
            '图片名': image_filename,
            'time_source': time_source,
            'is_valid': is_valid,
            'is_daytime': is_daytime,
            'visibility_time': '',  # 需人工或大模型判断
            'visibility': '',       # 需人工或大模型判断
            'time_diff_min': time_diff_min,
            'label': '',            # 需人工标注
            'remark': ''            # 备注留空
        })

    except Exception as e:
        print(f"抓取 {direction} 失败: {e}")

# 5. 更新 Excel 表格 (12列定制排版)
if new_data:
    headers = [
        '时间', '方向', '图片路径', '图片名', 'time_source', 'is_valid', 
        'is_daytime', 'visibility_time', 'visibility', 'time_diff_min', 'label', 'remark'
    ]
    
    if not os.path.exists(excel_path):
        wb = Workbook()
        ws = wb.active
        ws.title = "爬虫数据"
        ws.append(headers)
        
        ws.column_dimensions['A'].width = 20
        ws.column_dimensions['B'].width = 12
        ws.column_dimensions['C'].width = 45
        ws.column_dimensions['D'].width = 30
        for col in ['E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']:
            ws.column_dimensions[col].width = 15
        
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center')
        wb.save(excel_path)

    try:
        wb = load_workbook(excel_path)
        ws = wb.active
        for row in new_data:
            row_values = [
                row['时间'], row['方向'], row['图片路径'], row['图片名'],
                row['time_source'], row['is_valid'], row['is_daytime'],
                row['visibility_time'], row['visibility'], row['time_diff_min'],
                row['label'], row['remark']
            ]
            ws.append(row_values)
            
            for cell in ws[ws.max_row]:
                cell.alignment = Alignment(horizontal='left')
        wb.save(excel_path)
        print(f"\n✅ 数据已按 12 列完美格式追加至: {excel_path}")
    except Exception as e:
        print(f"保存表格失败: {e}")
else:
    print("\n🤷‍♂️ 本次运行没有产生新数据，表格未更新。")
