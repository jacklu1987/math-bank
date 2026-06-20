#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""题库 Word 导出 · 云端服务（方案 B）。
线上页面 POST 选中的题 → 这里用与本地完全相同的 导出Word.py 生成两份 .docx → 打包 zip 返回下载。
图片按需从已上线图床(PAGES_BASE/images/...)抓取并缓存，无需搬运整库图片。

环境变量：
  WORD_EXPORT_TOKEN  访问口令（页面请求需带相同 token，留空则不校验）
  PAGES_BASE         图床根地址，默认 https://math-bank.pages.dev
启动：uvicorn app:app --host 0.0.0.0 --port $PORT
"""
import os, io, sys, zipfile, tempfile, urllib.request
from urllib.parse import quote
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import 导出Word as W
from fastapi import FastAPI, Request, HTTPException
from fastapi.responses import Response, JSONResponse
from fastapi.middleware.cors import CORSMiddleware

PAGES_BASE = os.environ.get('PAGES_BASE', 'https://math-bank.pages.dev').rstrip('/')
TOKEN = os.environ.get('WORD_EXPORT_TOKEN', '')

# —— 图片解析：从图床按需抓取 + 进程内缓存 ——
# 带浏览器 UA，避免 Cloudflare 把默认 Python-urllib 当机器人拦掉（403）。
_cache = {}
_UA = ('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
       '(KHTML, like Gecko) Chrome/124.0 Safari/537.36')
def fetch_image(ref):
    name = (ref or '').split('/')[-1]
    if not name:
        return None
    if name in _cache:
        return _cache[name]
    url = f'{PAGES_BASE}/images/{name}'
    for _ in range(2):
        try:
            req = urllib.request.Request(url, headers={'User-Agent': _UA, 'Accept': 'image/*,*/*'})
            with urllib.request.urlopen(req, timeout=25) as r:
                data = r.read()
            if data and len(data) > 100:        # 像真图片（排除被拦返回的小页面）
                _cache[name] = data
                return data
        except Exception:
            pass
    return None
W.IMAGE_BYTES = fetch_image

app = FastAPI(title='题库Word导出')
app.add_middleware(CORSMiddleware, allow_origins=['*'], allow_methods=['*'], allow_headers=['*'])

@app.get('/')
def health():
    return {'ok': True, 'service': 'word-export', 'pages': PAGES_BASE}

@app.post('/api/export-word')
async def export_word(req: Request):
    try:
        body = await req.json()
    except Exception:
        raise HTTPException(400, 'bad json')
    if TOKEN and body.get('token') != TOKEN:
        raise HTTPException(401, '口令不正确')
    qs = body.get('questions') or []
    title = (body.get('title') or '练习试卷').strip()
    if not qs:
        raise HTTPException(400, '没有选中题目')
    tmp = tempfile.mkdtemp()
    try:
        blank, full = W.export_both(qs, title, out_dir=tmp)
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as z:
            z.write(blank, os.path.basename(blank))
            z.write(full, os.path.basename(full))
        fn = f'{title or "试卷"}.zip'
        return Response(content=buf.getvalue(), media_type='application/zip',
                        headers={'Content-Disposition': f"attachment; filename*=UTF-8''{quote(fn)}"})
    except Exception as e:
        import traceback; traceback.print_exc()
        return JSONResponse({'error': str(e)}, status_code=500)
