"""
FastAPI 服务器 - 为 pdfm_v3 提供 Web 访问能力
最小侵入式设计：自动适配所有 PDFMergerAPI 方法
"""
import os
import sys
import json
import base64
import tempfile
import uuid
from typing import Any, Dict, List, Optional
from pathlib import Path

from fastapi import FastAPI, File, UploadFile, HTTPException, Request
from fastapi.responses import FileResponse, JSONResponse, HTMLResponse
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
import uvicorn

# 导入业务逻辑
from pdfm_v3 import PDFMergerAPI
import ui

app = FastAPI(title="PDF Merger API Server")

# 配置 CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# 创建 API 实例
api = PDFMergerAPI()

# 临时文件存储
TEMP_DIR = Path(tempfile.gettempdir()) / "pdfm_server"
TEMP_DIR.mkdir(exist_ok=True)

# 文件映射：虚拟ID -> 真实路径
file_mapping: Dict[str, str] = {}

# 会话存储：每个客户端的上传文件
sessions: Dict[str, Dict[str, Any]] = {}


def get_session_id(request: Request) -> str:
    """获取或创建会话ID"""
    session_id = request.cookies.get("session_id")
    if not session_id:
        session_id = str(uuid.uuid4())
    if session_id not in sessions:
        sessions[session_id] = {"files": {}}
    return session_id


@app.get("/", response_class=HTMLResponse)
async def root():
    """返回前端页面"""
    # 注入垫片脚本
    html = ui.html_content
    
    # 在 </head> 前插入垫片
    shim_script = """
    <script src="/static/shim.js"></script>
    """
    html = html.replace("</head>", f"{shim_script}</head>")
    
    return html


@app.post("/api/upload_files")
async def upload_files(request: Request, files: List[UploadFile] = File(...)):
    """上传 PDF 文件"""
    session_id = get_session_id(request)
    result = []
    
    for file in files:
        if not file.filename.lower().endswith('.pdf'):
            continue
        
        # 保存到临时目录
        file_id = str(uuid.uuid4())
        file_path = TEMP_DIR / f"{file_id}.pdf"
        
        content = await file.read()
        with open(file_path, "wb") as f:
            f.write(content)
        
        # 记录文件映射
        file_mapping[file_id] = str(file_path)
        sessions[session_id]["files"][file_id] = {
            "path": str(file_path),
            "name": file.filename
        }
        
        # 获取页数
        try:
            import fitz
            doc = fitz.open(str(file_path))
            page_count = doc.page_count
            doc.close()
            
            result.append({
                "path": file_id,  # 返回虚拟ID
                "name": file.filename,
                "page_count": page_count
            })
        except Exception as e:
            print(f"解析文件失败: {e}")
    
    response = JSONResponse(content=result)
    response.set_cookie("session_id", session_id)
    return response


@app.post("/api/get_page_image")
async def get_page_image(request: Request):
    """获取页面图片"""
    data = await request.json()
    file_path = data.get("file_path")
    page_index = data.get("page_index")
    quality = data.get("quality", 1.0)
    
    # 转换虚拟路径为真实路径
    real_path = file_mapping.get(file_path, file_path)
    
    result = api.get_page_image(real_path, page_index, quality)
    return result


@app.post("/api/get_file_info")
async def get_file_info(request: Request):
    """获取文件信息"""
    data = await request.json()
    file_path = data.get("file_path")
    
    # 转换虚拟路径为真实路径
    real_path = file_mapping.get(file_path, file_path)
    
    result = api.get_file_info(real_path)
    return result


@app.post("/api/merge_pages")
async def merge_pages(request: Request):
    """合并页面"""
    data = await request.json()
    page_list = data.get("page_list", [])
    mode = data.get("mode", "normal")
    
    # 转换虚拟路径为真实路径
    for item in page_list:
        if item["path"] in file_mapping:
            item["path"] = file_mapping[item["path"]]
    
    # 生成输出文件路径
    output_id = str(uuid.uuid4())
    output_path = TEMP_DIR / f"merged_{output_id}.pdf"
    
    result = api.merge_pages(page_list, str(output_path), mode)
    
    if result.get("success"):
        # 读取文件内容
        with open(output_path, "rb") as f:
            pdf_content = f.read()
        
        # 转换为 Base64
        pdf_base64 = base64.b64encode(pdf_content).decode()
        
        # ========== 修改：不立即删除，而是保留在 file_mapping 中 ==========
        # 这样可以支持后续的 get_page_image 请求
        file_mapping[output_id] = str(output_path)
        
        # 返回 Base64 内容供前端缓存
        result["output_path"] = output_id
        result["pdf_content"] = pdf_base64
        result["file_size"] = len(pdf_content)
    
    return result


@app.get("/api/download/{file_id}")
async def download_file(file_id: str):
    """下载文件"""
    file_path = file_mapping.get(file_id)
    if not file_path or not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="文件不存在")
    
    return FileResponse(
        file_path,
        media_type="application/pdf",
        filename=f"merged_{file_id[:8]}.pdf"
    )


@app.post("/api/get_routes")
async def get_routes():
    """获取路线配置"""
    result = api.get_routes()
    return result


@app.post("/api/save_routes")
async def save_routes(request: Request):
    """保存路线配置"""
    data = await request.json()
    routes = data.get("routes", [])
    result = api.save_routes(routes)
    return result


@app.post("/api/generate_reimbursement_form")
async def generate_reimbursement_form(request: Request):
    """生成报销单"""
    data = await request.json()
    file_paths = data.get("file_paths", [])
    date_range = data.get("date_range", "")
    
    # 转换虚拟路径
    real_paths = []
    for item in file_paths:
        if isinstance(item, dict):
            path = item.get("path")
            if path in file_mapping:
                item["path"] = file_mapping[path]
            real_paths.append(item)
        else:
            real_paths.append(file_mapping.get(item, item))
    
    result = api.generate_reimbursement_form(real_paths, date_range)
    return result


@app.post("/api/calculate_invoice_amounts")
async def calculate_invoice_amounts(request: Request):
    """计算发票金额"""
    data = await request.json()
    pages_info = data.get("pages_info", [])
    
    # 转换虚拟路径
    for page in pages_info:
        if page["path"] in file_mapping:
            page["path"] = file_mapping[page["path"]]
    
    result = api.calculate_invoice_amounts(pages_info)
    return result


@app.post("/api/save_reimbursement_csv")
async def save_reimbursement_csv(request: Request):
    """保存报销单CSV（浏览器模式返回文件内容）"""
    data = await request.json()
    rows = data.get("rows", [])
    
    # 生成CSV内容
    import csv
    from io import StringIO
    
    output = StringIO()
    writer = csv.writer(output)
    writer.writerow(['序号', '发票号', '来源', '人数', '日期', '起点', '终点', '票额'])
    
    total = 0.0
    for r in rows:
        writer.writerow([r['id'], r.get('invoiceNo', '未识别'), r.get('source', ''), 
                        r['people'], r['date'], r['start'], r['end'], r['amount']])
        total += float(r['amount'])
    
    writer.writerow(['', '', '', '', '', '', '合计', f'{total:.2f}'])
    
    csv_content = output.getvalue()
    
    # 返回Base64编码的CSV
    csv_base64 = base64.b64encode(csv_content.encode('utf-8-sig')).decode()
    
    return {
        "success": True,
        "content": csv_base64,
        "filename": "报销单.csv"
    }


@app.post("/api/save_statistics_csv")
async def save_statistics_csv(request: Request):
    """保存统计CSV（浏览器模式返回文件内容）"""
    data = await request.json()
    amounts = data.get("amounts", [])
    
    import csv
    from io import StringIO
    
    output = StringIO()
    writer = csv.writer(output)
    writer.writerow(['序号', '发票号', '来源页面', '金额'])
    
    total = 0.0
    for idx, item in enumerate(amounts):
        pages_str = ', '.join(item.get('pages', []))
        writer.writerow([idx + 1, item.get('invoiceNo', '未识别'), pages_str, item['amount']])
        total += float(item['amount'])
    
    writer.writerow(['', '', '合计（已去重）', f'{total:.2f}'])
    
    csv_content = output.getvalue()
    csv_base64 = base64.b64encode(csv_content.encode('utf-8-sig')).decode()
    
    return {
        "success": True,
        "content": csv_base64,
        "filename": "发票统计.csv"
    }


@app.post("/api/print_pdf")
async def print_pdf(request: Request):
    """打印PDF"""
    data = await request.json()
    file_path = data.get("file_path")
    
    # 转换虚拟路径
    real_path = file_mapping.get(file_path, file_path)
    
    result = api.print_pdf(real_path)
    return result


# 静态文件服务（用于垫片脚本）
@app.get("/static/shim.js")
async def get_shim():
    """返回垫片脚本"""
    with open("shim.js", "r", encoding="utf-8") as f:
        content = f.read()
    return HTMLResponse(content=content, media_type="application/javascript")


if __name__ == "__main__":
    print("=" * 60)
    print("🚀 PDF Merger Server 启动中...")
    print("=" * 60)
    print(f"📁 临时文件目录: {TEMP_DIR}")
    print(f"🌐 访问地址: http://localhost:8000")
    print(f"📖 API 文档: http://localhost:8000/docs")
    print("=" * 60)
    
    uvicorn.run(app, host="0.0.0.0", port=8000, log_level="info")
