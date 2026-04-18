import os
import re
import requests
import pandas as pd
from io import BytesIO
from PIL import Image
import pytesseract
import platform
from datetime import datetime, timedelta

# 1. 智能识别系统配置 Tesseract 路径
if platform.system() == "Windows":
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# 2. 目录配置 (相对路径，云端本地通用)
base_dir = "."
image_dir = os.path.join(base_dir, "data", "raw_images")
excel_dir = os.path.join(base_dir, "data", "metadata")
excel_path = os.path.join(excel_dir, "表.xlsx")

os.makedirs(image_dir, exist_ok=True)
os.makedirs(excel_dir, exist_ok=True)

# 3. 读取现有表格 (用于防重复检查)
if os.path.exists(excel_path):
    df = pd.read_excel(excel_path)
else:
    # 如果表格不存在，创建一个空的 (请确保列名和你的实际需求一致)
    df = pd.DataFrame(columns=['时间', '方向', '图片路径'])


def get_real_time_from_image(img_bytes):
    """裁剪图片右上角并识别时间（精准提取 + 北京时间校准版）"""
    try:
        img = Image.open(BytesIO(img_bytes))
        width, height = img.size
        crop_box = (width - 600, 0, width, 150)
        cropped_img = img.crop(crop_box).convert('L')

        text = pytesseract.image_to_string(cropped_img, lang='eng')
        print(f"🤖 AI 看到的原始文字是: 【{text.strip()}】")

        date_match = re.search(r'(\d{1,2}[-/]\d{1,2})', text)
        time_match = re.search(r'(\d{2}:\d{2}:\d{2})', text)

        if date_match and time_match:
            month_day = date_match.group(1).replace('/', '-')
            hour_minute_second = time_match.group(1)

            # 【核心修复1】获取云端/本地的当前北京时间 (UTC+8)
            bjt_now = datetime.utcnow() + timedelta(hours=8)
            current_year = bjt_now.year

            final_time = f"{current_year}-{month_day} {hour_minute_second}"
            return final_time
    except Exception as e:
        print(f"OCR 识别出错: {e}")
    return None


# 4. 爬取目标
targets = [
    {"url": "http://view.iap.ac.cn:8080/imageview/northeast.jpg", "direction": "northeast"},
    {"url": "http://view.iap.ac.cn:8080/imageview/southwest.jpg", "direction": "southwest"}
]

new_data = []  # 暂存本次抓取的新数据

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

        # 【核心修复2】兜底时间也必须使用北京时间
        bjt_now = datetime.utcnow() + timedelta(hours=8)
        if extracted_time:
            final_time = extracted_time
        else:
            final_time = bjt_now.strftime('%Y-%m-%d %H:%M:%S')
            print(f"⚠️ 识别失败，使用当前北京时间兜底: {final_time}")

        # 【核心修复3】防重复机制：检查这个时间是否已经在表格里了
        # 注意：这里假设你的 Excel 里有一列叫 '时间'
        if not df.empty and '时间' in df.columns and final_time in df['时间'].values:
            print(f"🛑 发现重复数据！时间 {final_time} 已存在，跳过 {direction} 的保存。")
            continue

        # 【核心修复4】安全命名图片 (Windows 系统不允许文件名中包含冒号 :)
        safe_time_str = final_time.replace(':', '-')
        image_filename = f"{direction}_{safe_time_str}.jpg"
        image_path = os.path.join(image_dir, image_filename)

        # 保存图片
        with open(image_path, 'wb') as f:
            f.write(img_bytes)
        print(f"✅ 成功保存新图片: {image_filename}")

        # 记录新数据准备写入 Excel
        new_data.append({
            '时间': final_time,
            '方向': direction,
            '图片路径': image_path
        })

    except Exception as e:
        print(f"抓取 {direction} 失败: {e}")

# 5. 更新 Excel 表格
if new_data:
    df_new = pd.DataFrame(new_data)
    # 将新数据追加到旧数据后面
    df_combined = pd.concat([df, df_new], ignore_index=True)
    df_combined.to_excel(excel_path, index=False)
    print(f"\n✅ 数据已成功更新至: {excel_path}")
else:
    print("\n🤷‍♂️ 本次运行没有产生新数据，表格未更新。")
