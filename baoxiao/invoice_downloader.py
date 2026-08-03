#!/usr/bin/env python3
"""
Invoice Email Downloader
通过POP3下载邮件附件，自动解压ZIP，分类发票PDF
"""

import os
import poplib
import email
from email.header import decode_header
from datetime import datetime
from pathlib import Path
import argparse
import zipfile
from dotenv import load_dotenv
import re

try:
    from .invoice_processor import categorize_pdfs, generate_excel
    from .taxi_reimbursement import generate_taxi_reimbursement
except ImportError:  # Allow running this module directly from baoxiao/.
    from invoice_processor import categorize_pdfs, generate_excel
    from taxi_reimbursement import generate_taxi_reimbursement


def load_config():
    """从.env加载配置"""
    load_dotenv()
    config = {
        'pop3_server': os.getenv('POP3_SERVER', 'pop.126.com'),
        'pop3_port': int(os.getenv('POP3_PORT', '995')),
        'email_user': os.getenv('EMAIL_USER'),
        'email_auth_code': os.getenv('EMAIL_AUTH_CODE')
    }
    
    missing = [k for k, v in config.items() if not v]
    if missing:
        raise ValueError(f"缺少配置: {', '.join(missing)}，请检查.env文件")
    
    return config


def decode_str(s):
    """解码邮件头信息"""
    if s is None:
        return ""
    decoded_parts = decode_header(s)
    result = []
    for part, charset in decoded_parts:
        if isinstance(part, bytes):
            result.append(part.decode(charset or 'utf-8', errors='ignore'))
        else:
            result.append(part)
    return ''.join(result)


def get_email_date(msg):
    """获取邮件日期"""
    date_str = msg.get('Date', '')
    try:
        # 尝试多种日期格式
        for fmt in [
            '%a, %d %b %Y %H:%M:%S %z',
            '%a, %d %b %Y %H:%M:%S %Z',
            '%d %b %Y %H:%M:%S %z',
            '%Y-%m-%d %H:%M:%S'
        ]:
            try:
                return datetime.strptime(date_str.strip(), fmt)
            except ValueError:
                continue
        # 如果都失败，尝试解析简化格式
        match = re.search(r'(\d{1,2})\s+(\w{3})\s+(\d{4})', date_str)
        if match:
            day, month, year = match.groups()
            month_map = {'Jan': 1, 'Feb': 2, 'Mar': 3, 'Apr': 4, 'May': 5, 'Jun': 6,
                        'Jul': 7, 'Aug': 8, 'Sep': 9, 'Oct': 10, 'Nov': 11, 'Dec': 12}
            return datetime(int(year), month_map.get(month, 1), int(day))
    except Exception:
        pass
    return None


def is_invoice_file(filename):
    """判断是否为发票相关文件"""
    invoice_keywords = ['发票', 'invoice', '收据', '报销', '行程单', '行程明细', '客票']
    return any(kw in filename.lower() for kw in invoice_keywords)


def save_attachments(msg, output_dir, start_date, end_date):
    """保存邮件附件"""
    saved_files = []
    
    for part in msg.walk():
        content_type = part.get_content_type()
        content_disposition = str(part.get('Content-Disposition', ''))
        
        # 获取文件名
        filename = decode_str(part.get_filename())
        if not filename:
            continue
        
        # 检查日期范围
        email_date = get_email_date(msg)
        if email_date and not (start_date <= email_date.replace(tzinfo=None) <= end_date):
            continue
        
        # 保存附件
        if 'attachment' in content_disposition or content_type in [
            'application/pdf', 'application/zip', 'application/x-zip-compressed',
            'application/octet-stream'
        ]:
            filepath = os.path.join(output_dir, filename)
            
            # 处理文件名冲突
            counter = 1
            while os.path.exists(filepath):
                name, ext = os.path.splitext(filename)
                filepath = os.path.join(output_dir, f"{name}_{counter}{ext}")
                counter += 1
            
            with open(filepath, 'wb') as f:
                f.write(part.get_payload(decode=True))
            
            saved_files.append(filepath)
            print(f"  保存: {filename}")
    
    return saved_files


def extract_zip_files(output_dir):
    """解压所有ZIP文件"""
    extracted_files = []
    
    for file in Path(output_dir).glob('*.zip'):
        try:
            extract_dir = output_dir
            with zipfile.ZipFile(file, 'r') as zip_ref:
                zip_ref.extractall(extract_dir)
                extracted_files.extend([os.path.join(extract_dir, f) for f in zip_ref.namelist()])
                print(f"  解压: {file.name}")
            os.remove(file)  # 删除已解压的ZIP
        except Exception as e:
            print(f"  解压失败 {file.name}: {e}")
    
    return extracted_files


def download_emails(start_date, end_date, raw_dir, classify_dir):
    """主下载函数"""
    config = load_config()
    
    print(f"连接到 {config['pop3_server']}...")
    
    # 连接POP3服务器
    server = poplib.POP3_SSL(config['pop3_server'], config['pop3_port'])
    server.user(config['email_user'])
    server.pass_(config['email_auth_code'])
    
    # 获取邮件列表
    messages = server.list()
    message_count = len(messages[1])
    print(f"共 {message_count} 封邮件")
    
    # 创建输出目录
    os.makedirs(raw_dir, exist_ok=True)
    os.makedirs(classify_dir, exist_ok=True)
    
    saved_files = []
    
    # 遍历邮件
    for i in range(message_count):
        try:
            msg_lines = server.retr(i + 1)[1]
            msg = email.message_from_bytes(b'\r\n'.join(msg_lines))
            
            subject = decode_str(msg.get('Subject', ''))
            sender = decode_str(msg.get('From', ''))
            
            # 检查是否与发票相关
            if not is_invoice_file(subject) and not is_invoice_file(sender):
                # 也检查附件名
                has_invoice_attachment = False
                for part in msg.walk():
                    fname = decode_str(part.get_filename())
                    if fname and is_invoice_file(fname):
                        has_invoice_attachment = True
                        break
                if not has_invoice_attachment:
                    continue
            
            print(f"\n处理邮件: {subject}")
            files = save_attachments(msg, raw_dir, start_date, end_date)
            saved_files.extend(files)
            
        except Exception as e:
            print(f"处理邮件 {i+1} 时出错: {e}")
    
    server.quit()
    
    # 解压ZIP文件
    print("\n解压ZIP文件...")
    zip_files = extract_zip_files(raw_dir)
    
    # 收集所有PDF文件
    pdf_files = list(Path(raw_dir).rglob('*.pdf'))
    pdf_files = [str(f) for f in pdf_files]
    
    # 分类PDF
    print("\n分类发票...")
    results, invoice_info = categorize_pdfs(pdf_files, classify_dir)
    
    # 输出统计
    print("\n=== 分类统计 ===")
    for category, files in results.items():
        print(f"{category}: {len(files)} 份")
    
    return results, invoice_info


def main():
    parser = argparse.ArgumentParser(description='下载并分类邮件中的发票附件')
    parser.add_argument('--start', required=True, help='开始日期 (YYYY-MM-DD)')
    parser.add_argument('--end', required=True, help='结束日期 (YYYY-MM-DD)')
    parser.add_argument('--raw-dir', default='./tem', help='原始文件目录 (默认: ./tem)')
    parser.add_argument('--classify-dir', default='.', help='分类文件目录 (默认: 当前目录)')
    parser.add_argument('--taxi-template', default=str(Path(__file__).with_name('出租车报销单模板.xlsx')), help='出租车报销单模板路径')
    
    args = parser.parse_args()
    
    try:
        start_date = datetime.strptime(args.start, '%Y-%m-%d')
        end_date = datetime.strptime(args.end, '%Y-%m-%d').replace(hour=23, minute=59, second=59)
    except ValueError:
        print("日期格式错误，请使用 YYYY-MM-DD 格式")
        return
    
    print(f"=== 发票邮件下载器 ===")
    print(f"时间范围: {args.start} 至 {args.end}")
    print(f"原始文件目录: {args.raw_dir}")
    print(f"分类文件目录: {args.classify_dir}\n")
    
    try:
        results, invoice_info = download_emails(start_date, end_date, args.raw_dir, args.classify_dir)
        
        # 生成四分类 Excel 工作表
        excel_file = os.path.join(args.classify_dir, '发票.xlsx')
        print(f"\n生成Excel文件: {excel_file}")
        generate_excel(invoice_info, excel_file)
        print(f"Excel文件已生成: {excel_file}")

        taxis = [row for row in invoice_info if row['category'] == '出租票']
        if taxis:
            taxi_form = os.path.join(args.classify_dir, '出租车报销单.xlsx')
            print(f"\n生成出租车报销单: {taxi_form}")
            generate_taxi_reimbursement(taxis, args.taxi_template, taxi_form)
            print(f"出租车报销单已生成: {taxi_form}")
        
        print("\n下载完成！")
    except Exception as e:
        print(f"\n错误: {e}")
        raise


if __name__ == '__main__':
    main()

