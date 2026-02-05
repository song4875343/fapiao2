import fitz
import re

doc = fitz.open('【风韵出行-24.01元-1个行程】高德打车电子发票.pdf')
page = doc[0]
blocks = page.get_text('dict')['blocks']

print('=== 查找小写关键字位置 ===')
keyword_rect = None
for block in blocks:
    if 'lines' in block:
        for line in block['lines']:
            for span in line['spans']:
                text = span.get('text', '')
                if '小写' in text:
                    print(f'找到小写: [{text}]')
                    print(f'坐标: {span["bbox"]}')
                    keyword_rect = span['bbox']
                    break

if keyword_rect:
    kw_y0, kw_y1 = keyword_rect[1], keyword_rect[3]
    kw_x1 = keyword_rect[2]
    print(f'\n=== 查找同一行右侧的文本 ===')
    print(f'关键字Y范围: {kw_y0:.2f} - {kw_y1:.2f}')
    print(f'关键字右边界X: {kw_x1:.2f}')
    
    candidates = []
    for block in blocks:
        if 'lines' in block:
            for line in block['lines']:
                for span in line['spans']:
                    span_bbox = span['bbox']
                    span_y0, span_y1 = span_bbox[1], span_bbox[3]
                    span_x0 = span_bbox[0]
                    span_text = span['text'].strip()
                    
                    y_overlap = not (span_y1 < kw_y0 or span_y0 > kw_y1)
                    is_right = span_x0 >= kw_x1 - 10
                    
                    if y_overlap and is_right and span_text:
                        distance = span_x0 - kw_x1
                        print(f'文本: [{span_text}], X坐标: {span_x0:.2f}, 距离: {distance:.2f}, Y: {span_y0:.2f}-{span_y1:.2f}')
                        
                        # 检查是否包含金额符号
                        if '¥' in span_text or '￥' in span_text:
                            print(f'  -> 包含金额符号！')
                            amount_match = re.search(r'[¥￥]\s*([\d,]+\.?\d*)', span_text)
                            if amount_match:
                                print(f'  -> 提取到金额: {amount_match.group(1)}')
                                candidates.append((distance, amount_match.group(1)))

    print(f'\n=== 候选金额 ===')
    if candidates:
        candidates.sort(key=lambda x: x[0])
        print(f'最近的金额: {candidates[0][1]}')
    else:
        print('未找到金额！')

print('\n=== 使用正则表达式查找 ===')
text = page.get_text()
patterns = [
    r'小写[：:\s]*[¥￥]\s*([\d,]+\.?\d*)',
    r'\(小写\)[：:\s]*[¥￥]\s*([\d,]+\.?\d*)',
    r'价税合计.*?[¥￥]\s*([\d,]+\.?\d*)',
]

for pattern in patterns:
    matches = re.findall(pattern, text, re.S)
    if matches:
        print(f'模式 [{pattern}] 匹配到: {matches}')

doc.close()
