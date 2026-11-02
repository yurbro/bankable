"""Render docs/src/*.html to docs/images/*.png with the bundled Chromium (run from repo root: python docs/render.py)."""
import asyncio, os
from playwright.async_api import async_playwright

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "src"); OUT = os.path.join(HERE, "images")
PAGES = ["banner", "workflow", "audience", "outputs", "banner_en", "workflow_en", "audience_en", "outputs_en"]

async def main():
    os.makedirs(OUT, exist_ok=True)
    async with async_playwright() as p:
        b = await p.chromium.launch()
        for name in PAGES:
            page = await b.new_page(viewport={"width": 1600, "height": 600}, device_scale_factor=1.5)
            await page.goto("file://" + os.path.join(SRC, name + ".html"))
            await page.wait_for_timeout(300)
            out = os.path.join(OUT, name + ".png")
            await page.screenshot(path=out, full_page=True)
            from PIL import Image
            im = Image.open(out).convert("RGB")
            im = im.quantize(colors=256, method=Image.Quantize.MEDIANCUT, dither=Image.Dither.FLOYDSTEINBERG) if not name.startswith("outputs") else im
            im.save(out, optimize=True)
            print("[OK]", name, os.path.getsize(out) // 1024, "KB")
        await b.close()

asyncio.run(main())
