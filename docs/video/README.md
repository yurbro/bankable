# Guide video

`bankable-guide.mp4` — 1:21, 1920×1080, English voice-over with burnt-in English + Chinese captions. Both READMEs embed it inline from a GitHub upload (user-attachments link), since GitHub only plays videos uploaded through its web editor; `docs/images/video_poster.jpg` is the poster frame. After rebuilding, drag the new mp4 into the README editor on GitHub and swap the link.

Everything is code, so the video can be changed and rebuilt like the rest of the repo.

| File | What it does |
|---|---|
| `src/script_en.json` | Narration, one entry per scene |
| `src/tts.py` | Narration → per-scene WAVs with Kokoro v1.1 (ONNX, voice `bf_vale`) via `onnxruntime` |
| `src/timeline.json` | Scene lengths and line timings (built from the narration), plus the Chinese caption lines |
| `src/video.html` | All nine scenes as one deterministic page: `render(scene, t)` draws any frame |
| `src/render.py` | `stills s3:4.0 …` for review frames; `scenes` renders each scene to its own MP4 (Playwright + ffmpeg) |
| `src/audio.py` | Generated music bed + impacts, narration placed per scene, ducked mix → `soundtrack.wav` |
| `src/poster.py` | README poster frame |

## Rebuild

```bash
cd docs/video/src
# 1. narration (needs the Kokoro v1.1 model: kokoro-multi-lang-v1_1 from k2-fsa/sherpa-onnx tts-models releases)
KOKORO_DIR=/path/to/kokoro-multi-lang-v1_1 KLEX=gb python tts.py script_en.json vo --sid 2
# 2. if the script changed, rebuild timeline.json from vo/durations.json (keep the zh caption lines in step)
# 3. frames → scenes → soundtrack → final file
python render.py scenes
python audio.py                      # reads vo/*.wav
for s in s1 s2 s3 s4 s5 s6 s7 s8 s9; do echo "file 'scenes/$s.mp4'"; done > list.txt
ffmpeg -f concat -safe 0 -i list.txt -c copy silent.mp4
ffmpeg -i silent.mp4 -i soundtrack.wav -map 0:v -map 1:a -c:v libx264 -preset medium -crf 25 -tune animation \
       -pix_fmt yuv420p -c:a aac -b:a 160k -shortest -movflags +faststart ../bankable-guide.mp4
python poster.py && cp video_poster.jpg ../../images/
```

Changing one scene only needs `python render.py scenes s4` and the last two ffmpeg steps.
