# -*- coding: utf-8 -*-
"""Annotate tutorial screenshots with numbered markers, circles, and arrows."""
from PIL import Image, ImageDraw, ImageFont
import os, math

FONT = None
# Try common Chinese fonts
for fp in ['C:/Windows/Fonts/simhei.ttf', 'C:/Windows/Fonts/msyh.ttc',
           'C:/Windows/Fonts/simsun.ttc', 'C:/Windows/Fonts/simkai.ttf']:
    if os.path.exists(fp):
        FONT = fp
        break
if not FONT:
    FONT = ImageFont.load_default()

def get_font(size):
    try:
        return ImageFont.truetype(FONT, size)
    except Exception:
        return ImageFont.load_default()

def draw_circle(draw, cx, cy, r, outline='#FF0000', width=4):
    """Draw a red circle."""
    for i in range(width):
        draw.ellipse([cx-r-i, cy-r-i, cx+r+i, cy+r+i], outline=outline)

def draw_arrow(draw, x1, y1, x2, y2, color='#FF0000', width=4):
    """Draw an arrow from (x1,y1) to (x2,y2)."""
    draw.line([(x1, y1), (x2, y2)], fill=color, width=width)
    # Arrowhead
    angle = math.atan2(y2-y1, x2-x1)
    alen = 15
    draw.line([(x2, y2),
               (x2 - alen*math.cos(angle-math.pi/6),
                y2 - alen*math.sin(angle-math.pi/6))], fill=color, width=4)
    draw.line([(x2, y2),
               (x2 - alen*math.cos(angle+math.pi/6),
                y2 - alen*math.sin(angle+math.pi/6))], fill=color, width=4)

def draw_rect(draw, x, y, w, h, color='#FF0000', width=4, radius=8):
    """Draw a rounded rectangle."""
    for i in range(width):
        draw.rounded_rectangle([x-i, y-i, x+w+i, y+h+i], radius=radius+i, outline=color)

def draw_number(draw, x, y, num, size=28):
    """Draw a circled number marker."""
    r = size // 2 + 8
    draw.ellipse([x-r, y-r, x+r, y+r], fill='#FF0000')
    font = get_font(size)
    text = str(num)
    bbox = draw.textbbox((0, 0), text, font=font)
    tw, th = bbox[2]-bbox[0], bbox[3]-bbox[1]
    draw.text((x-tw/2, y-th/2-2), text, fill='white', font=font)

def draw_label(draw, x, y, text, bg='#FF0000', size=16):
    """Draw a text label with red background."""
    font = get_font(size)
    bbox = draw.textbbox((0, 0), text, font=font)
    tw, th = bbox[2]-bbox[0], bbox[3]-bbox[1]
    pad = 6
    draw.rounded_rectangle([x-pad, y-pad, x+tw+pad, y+th+pad], radius=4, fill=bg)
    draw.text((x, y), text, fill='white', font=font)

# ================================================================
# ANNOTATE EACH IMAGE
# ================================================================
tutorial_dir = 'tutorial'
annotated_dir = 'tutorial_annotated'
os.makedirs(annotated_dir, exist_ok=True)

# --- Image 1: Home Page ---
img = Image.open(f'{tutorial_dir}/01-home.png'); d = ImageDraw.Draw(img)
w, h = img.size
# Left sidebar: nav area
draw_rect(d, 8, 40, 260, h-60, '#FF6B35', 4, 8)
draw_label(d, 20, 15, '① 导航栏: 8个功能页面', '#FF6B35', 18)
# Company selector area
draw_rect(d, 15, 110, 250, 65, '#2196F3', 3, 6)
draw_label(d, 20, 185, '② 选择公司 & 年月', '#2196F3', 16)
# Main content area
draw_rect(d, 285, 40, w-300, 200, '#4CAF50', 3, 6)
draw_label(d, 300, 50, '③ 主工作区', '#4CAF50', 16)
draw_number(d, 140, 145, 1, 28)
draw_number(d, 140, 250, 2, 28)
draw_number(d, w//2, 150, 3, 28)
img.save(f'{annotated_dir}/01-home.png'); print('Annotated 01-home.png')

# --- Image 2: Company Creation ---
img = Image.open(f'{tutorial_dir}/02-company-empty.png'); d = ImageDraw.Draw(img)
w, h = img.size
# Company name field
draw_rect(d, 310, 160, 350, 45, '#FF6B35', 3, 6)
draw_label(d, 310, 130, '① 输入公司简称', '#FF6B35', 16)
# Full name field
draw_rect(d, 310, 230, 350, 45, '#2196F3', 3, 6)
draw_label(d, 310, 200, '② 输入公司全称', '#2196F3', 16)
# Create button
draw_rect(d, 310, 295, 120, 42, '#4CAF50', 4, 8)
draw_label(d, 310, 350, '③ 点击创建', '#4CAF50', 16)
draw_number(d, 490, 185, 1, 28)
draw_number(d, 490, 255, 2, 28)
draw_number(d, 370, 320, 3, 28)
img.save(f'{annotated_dir}/02-company.png'); print('Annotated 02-company.png')

# --- Image 3: Opening Balance ---
img = Image.open(f'{tutorial_dir}/03-opening-balance.png'); d = ImageDraw.Draw(img)
w, h = img.size
# BS upload area
draw_rect(d, 290, 100, w-310, 80, '#FF6B35', 3, 6)
draw_label(d, 300, 75, '① 上传去年BS自动填写 (新功能!)', '#FF6B35', 16)
# Asset fields
draw_rect(d, 290, 200, 300, 250, '#2196F3', 3, 6)
draw_label(d, 300, 175, '② 资产类科目', '#2196F3', 16)
# Liability fields
draw_rect(d, 620, 200, 300, 250, '#4CAF50', 3, 6)
draw_label(d, 630, 175, '③ 负债及权益类', '#4CAF50', 16)
# Save button
draw_rect(d, 290, 480, 200, 42, '#E91E63', 4, 8)
draw_label(d, 300, 535, '④ 保存', '#E91E63', 16)
draw_number(d, w//2, 145, 1, 26)
draw_number(d, 440, 330, 2, 26)
draw_number(d, 770, 330, 3, 26)
draw_number(d, 390, 505, 4, 26)
img.save(f'{annotated_dir}/03-opening.png'); print('Annotated 03-opening.png')

# --- Image 4: Data Import (Batch) ---
img = Image.open(f'{tutorial_dir}/05-import.png'); d = ImageDraw.Draw(img)
w, h = img.size
# Batch upload area
draw_rect(d, 290, 90, w-310, 120, '#FF6B35', 4, 8)
draw_label(d, 300, 65, '① 智能批量导入: 一次性选所有文件', '#FF6B35', 16)
# Individual uploaders
draw_rect(d, 290, 235, 350, 200, '#2196F3', 3, 6)
draw_label(d, 300, 220, '② 也可单个上传', '#2196F3', 16)
# Save button
draw_rect(d, 290, 460, 200, 42, '#4CAF50', 4, 8)
draw_label(d, 300, 515, '③ 保存并核验', '#4CAF50', 16)
draw_number(d, w//2, 155, 1, 28)
draw_number(d, 460, 340, 2, 28)
draw_number(d, 390, 485, 3, 28)
img.save(f'{annotated_dir}/04-import.png'); print('Annotated 04-import.png')

# --- Image 5: Verify ---
img = Image.open(f'{tutorial_dir}/06-verify.png'); d = ImageDraw.Draw(img)
w, h = img.size
# Metrics area
draw_rect(d, 290, 100, w-310, 100, '#FF6B35', 3, 6)
draw_label(d, 300, 75, '① 检查数据摘要', '#FF6B35', 16)
# Confirm button
draw_rect(d, 290, 230, 200, 42, '#4CAF50', 4, 8)
draw_label(d, 300, 285, '② 确认无误，去生成报表', '#4CAF50', 16)
draw_number(d, w//2, 155, 1, 28)
draw_number(d, 390, 255, 2, 28)
img.save(f'{annotated_dir}/05-verify.png'); print('Annotated 05-verify.png')

# --- Image 6: Generate ---
img = Image.open(f'{tutorial_dir}/07-generate.png'); d = ImageDraw.Draw(img)
w, h = img.size
# Custom template expander
draw_rect(d, 290, 90, w-310, 70, '#9C27B0', 3, 6)
draw_label(d, 300, 65, '📋 自定义模板(可选)', '#9C27B0', 14)
# Batch gen
draw_rect(d, 290, 180, w-310, 100, '#FF6B35', 3, 6)
draw_label(d, 300, 155, '① 批量生成(多月份)', '#FF6B35', 16)
# Single gen button
draw_rect(d, 290, 310, 200, 45, '#4CAF50', 4, 8)
draw_label(d, 300, 365, '② 单月生成', '#4CAF50', 16)
draw_number(d, w//2, 235, 1, 28)
draw_number(d, 390, 335, 2, 28)
img.save(f'{annotated_dir}/06-generate.png'); print('Annotated 06-generate.png')

# --- Image 7: Export ---
img = Image.open(f'{tutorial_dir}/08-export.png'); d = ImageDraw.Draw(img)
w, h = img.size
# Download buttons
draw_rect(d, 290, 100, 300, 80, '#FF6B35', 3, 6)
draw_label(d, 300, 75, '① 下载三大报表', '#FF6B35', 16)
# Analysis report
draw_rect(d, 290, 210, w-310, 120, '#2196F3', 3, 6)
draw_label(d, 300, 188, '② 生成财务分析报告', '#2196F3', 16)
draw_number(d, w//2, 145, 1, 28)
draw_number(d, w//2, 275, 2, 28)
img.save(f'{annotated_dir}/07-export.png'); print('Annotated 07-export.png')

# --- Image 8: History ---
img = Image.open(f'{tutorial_dir}/09-history.png'); d = ImageDraw.Draw(img)
w, h = img.size
# Multi-select
draw_rect(d, 290, 100, w-310, 80, '#FF6B35', 3, 6)
draw_label(d, 300, 75, '① 多选对比月份', '#FF6B35', 16)
# Table area
draw_rect(d, 290, 200, w-310, 200, '#2196F3', 3, 6)
draw_label(d, 300, 175, '② 横排对比表 (自动标注异常)', '#2196F3', 16)
draw_number(d, w//2, 145, 1, 28)
draw_number(d, w//2, 305, 2, 28)
img.save(f'{annotated_dir}/08-history.png'); print('Annotated 08-history.png')

print('\nAll 8 images annotated!')
print(f'Output: {annotated_dir}/')
