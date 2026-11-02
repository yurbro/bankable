"""Render video.html scenes.
  python render.py stills s1:2.0 s1:7.5 ...   -> stills/<id>_<t>.jpg
  python render.py scenes [ids...]            -> scenes/<id>.mp4 (silent, 1920x1080@30, H.264)
"""
import json, os, sys, subprocess, time
from playwright.sync_api import sync_playwright

HERE = os.path.dirname(os.path.abspath(__file__))
TLD = json.load(open(os.path.join(HERE, "timeline.json"), encoding="utf-8"))
src = open(os.path.join(HERE, "video.html"), encoding="utf-8").read().replace("__TIMELINE__", json.dumps(TLD, ensure_ascii=False))
BUILT = os.path.join(HERE, "_built.html")
open(BUILT, "w", encoding="utf-8").write(src)
FPS = 30
IDS = [s["id"] for s in TLD["scenes"]]

def page_ctx(p):
    b = p.chromium.launch(args=["--disable-gpu", "--font-render-hinting=none"])
    pg = b.new_page(viewport={"width": 1920, "height": 1080}, device_scale_factor=1)
    errs = []
    pg.on("pageerror", lambda e: errs.append(str(e)))
    pg.goto("file://" + BUILT); pg.wait_for_timeout(600)
    pg.evaluate("document.fonts.ready")
    return b, pg, errs

def stills(specs):
    os.makedirs(os.path.join(HERE, "stills"), exist_ok=True)
    with sync_playwright() as p:
        b, pg, errs = page_ctx(p)
        for s in specs:
            sid, t = s.split(":"); i = IDS.index(sid)
            pg.evaluate(f"render({i},{float(t)})")
            out = os.path.join(HERE, "stills", f"{sid}_{t}.jpg")
            pg.screenshot(path=out, type="jpeg", quality=92)
            print(out)
        if errs: print("PAGE ERRORS:", errs)
        b.close()

def scenes(ids):
    os.makedirs(os.path.join(HERE, "scenes"), exist_ok=True)
    with sync_playwright() as p:
        b, pg, errs = page_ctx(p)
        for sid in ids:
            i = IDS.index(sid); sc = TLD["scenes"][i]
            n = round(sc["dur"] * FPS)
            tmp = os.path.join(HERE, "scenes", f"_{sid}.tmp.mp4"); out = os.path.join(HERE, "scenes", f"{sid}.mp4")
            ff = subprocess.Popen(["ffmpeg", "-y", "-loglevel", "error", "-f", "image2pipe", "-framerate", str(FPS), "-c:v", "mjpeg", "-i", "-",
                                   "-c:v", "libx264", "-preset", "veryfast", "-crf", "19", "-pix_fmt", "yuv420p", "-r", str(FPS),
                                   "-g", "60", "-movflags", "+faststart", tmp], stdin=subprocess.PIPE)
            t0 = time.time()
            for f in range(n):
                pg.evaluate(f"render({i},{f / FPS})")
                ff.stdin.write(pg.screenshot(type="jpeg", quality=92))
            ff.stdin.close(); rc = ff.wait()
            if rc == 0: os.replace(tmp, out)
            print(sid, n, "frames", f"{time.time() - t0:.0f}s", "rc", rc, flush=True)
        if errs: print("PAGE ERRORS:", errs)
        b.close()

if __name__ == "__main__":
    mode = sys.argv[1]
    if mode == "stills": stills(sys.argv[2:])
    else: scenes(sys.argv[2:] or IDS)
