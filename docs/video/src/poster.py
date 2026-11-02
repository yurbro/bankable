"""Poster frame for the README: logo moment + play button."""
import os
from playwright.sync_api import sync_playwright
from PIL import Image
HERE = os.path.dirname(os.path.abspath(__file__))
BTN = """
const b=document.createElement('div');
b.style.cssText='position:absolute;left:50%;top:870px;transform:translateX(-50%);display:flex;align-items:center;gap:22px;padding:18px 34px 18px 20px;border-radius:999px;background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.18);backdrop-filter:blur(8px);font:600 28px/1 Inter,sans-serif;color:#eef2f8;box-shadow:0 20px 60px rgba(0,0,0,.5)';
b.innerHTML='<span style="width:64px;height:64px;border-radius:50%;background:#3b8cff;display:flex;align-items:center;justify-content:center;box-shadow:0 0 40px rgba(59,140,255,.7)"><svg width="26" height="26" viewBox="0 0 24 24"><path d="M7 4.5v15l13-7.5z" fill="#fff"/></svg></span>Watch the guide <span style="color:#8b95a8;font-weight:500">· 1:21 · with voice-over</span>';
document.getElementById('stage').appendChild(b);
document.getElementById('cap').style.display='none';document.getElementById('prog').style.display='none';
"""
with sync_playwright() as p:
    br = p.chromium.launch(); pg = br.new_page(viewport={"width": 1920, "height": 1080})
    pg.goto("file://" + os.path.join(HERE, "_built.html")); pg.wait_for_timeout(600)
    pg.evaluate("render(0, 8.9)")
    pg.evaluate(BTN)
    pg.screenshot(path=os.path.join(HERE, "poster_full.png"))
    br.close()
im = Image.open(os.path.join(HERE, "poster_full.png")).convert("RGB").resize((1600, 900), Image.LANCZOS)
im.save(os.path.join(HERE, "video_poster.jpg"), quality=88, optimize=True, progressive=True)
print(os.path.getsize(os.path.join(HERE, "video_poster.jpg")) // 1024, "KB")
