"""Kokoro v1.1 (ONNX) English narration via onnxruntime + misaki-style US lexicon.
Usage: python tts.py script.json outdir --sid 2 --speed 1.0
script.json: {"scenes":[{"id":"s1","lines":["sentence", ...]}, ...]}
Writes outdir/<id>.wav (24 kHz) and outdir/durations.json
"""
import json, re, sys, os, argparse
import numpy as np, onnxruntime as ort
from scipy.io import wavfile

D = os.environ.get("KOKORO_DIR", os.path.join(os.path.dirname(os.path.abspath(__file__)), "kokoro-multi-lang-v1_1"))
tok = {}
for line in open(os.path.join(D, "tokens.txt"), encoding="utf-8"):
    line = line.rstrip("\n")
    if not line: continue
    s, i = line.rsplit(" ", 1)
    tok[s if s else " "] = int(i)
tok[" "] = 16

lex = {}
LEXF = os.environ.get("KLEX", "us")
for line in open(os.path.join(D, f"lexicon-{LEXF}-en.txt"), encoding="utf-8"):
    p = line.split()
    if p and p[0] not in lex: lex[p[0]] = "".join(p[1:])
VOICELESS = set("ptkfθsʃʧ")
def morph(lw):
    if lw.endswith("es") and lw[:-2] in lex and lex[lw[:-2]][-1] in "sʃʧzʒʤ": return lex[lw[:-2]] + "ᵻz"
    if lw.endswith("s") and lw[:-1] in lex:
        b = lex[lw[:-1]]; return b + ("s" if b[-1] in VOICELESS else "z")
    for suf in ("ed", "d"):
        if lw.endswith(suf) and lw[:-len(suf)] in lex:
            b = lex[lw[:-len(suf)]]
            return b + ("ᵻd" if b[-1] in "td" else ("t" if b[-1] in VOICELESS else "d"))
    if lw.endswith("ing") and lw[:-3] in lex: return lex[lw[:-3]] + "ɪŋ"
    return None

# hand overrides: weak forms & spelled items
OVR = {
    "a": "ə", "an": "ən", "and": "ænd", "of": "ʌv", "to": "tə", "for": "fɔɹ", "or": "ɔɹ",
    "cfo": "sˌiɛfˈO", "github": "ɡˈɪthˌʌb", "ai": "ˌAˈI", "your": "jʊɹ", "is": "ɪz",
    "it": "ɪt", "in": "ɪn", "as": "æz", "with": "wɪð", "by": "bI", "at": "æt", "be": "bi",
}
if LEXF == "gb":
    OVR = {"a": "ə", "an": "ən", "cfo": "sˌiːɛfˈQ", "github": "ɡˈɪthˌʌb", "ai": "ˌAˈI", "and": "ənd", "to": "tə", "for": "fə", "your": "jɔː", "of": "əv", "or": "ɔː"}
VOWEL_START = re.compile(r"^[ˈˌ]?[aeiouæɑɐɒɔəɛɜɪʊʌAIOWYᵊ]")

def word_ph(w, nxt):
    lw = w.lower()
    if lw == "the":
        return ("ðiː" if LEXF == "gb" else "ði") if (nxt and VOWEL_START.match(nxt)) else "ðə"
    if lw in OVR: return OVR[lw]
    if lw in lex: return lex[lw]
    if lw.endswith("'s") and lw[:-2] in lex: return lex[lw[:-2]] + "z"
    m = morph(lw)
    if m: return m
    raise KeyError(w)

def g2p(text):
    # split hyphenated words, keep punctuation
    text = text.replace("—", ", ").replace("-", " ")
    parts = re.findall(r"[A-Za-z']+|[.,!?;:]", text)
    words = [p for p in parts]
    out = []
    # compute phonemes right-to-left to know next word's sound for "the"
    ph = [None] * len(words)
    nxt = None
    for i in range(len(words) - 1, -1, -1):
        w = words[i]
        if re.match(r"[.,!?;:]", w):
            ph[i] = w; continue
        ph[i] = word_ph(w, nxt); nxt = ph[i]
    s = ""
    for p in ph:
        if p in ".,!?;:": s = s.rstrip(" ") + p + " "
        else: s += p + " "
    return s.strip()

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("script"); ap.add_argument("outdir")
    ap.add_argument("--sid", type=int, default=2); ap.add_argument("--speed", type=float, default=1.0)
    ap.add_argument("--gap", type=float, default=0.32, help="silence between sentences (s)")
    a = ap.parse_args()
    os.makedirs(a.outdir, exist_ok=True)
    voices = np.fromfile(os.path.join(D, "voices.bin"), dtype=np.float32).reshape(-1, 510, 256)
    sess = ort.InferenceSession(os.path.join(D, "model.onnx"), providers=["CPUExecutionProvider"])
    sc = json.load(open(a.script, encoding="utf-8"))
    durs = {}
    for scene in sc["scenes"]:
        chunks, marks, t = [], [], 0.0
        for line in scene["lines"]:
            ps = g2p(line)
            ids = [tok[c] for c in ps if c in tok]
            missing = [c for c in ps if c not in tok]
            if missing: print("WARN missing tokens", set(missing))
            style = voices[a.sid][min(len(ids), 509)][None, :]
            audio = sess.run(None, {"tokens": np.array([[0] + ids + [0]], dtype=np.int64),
                                    "style": style.astype(np.float32),
                                    "speed": np.array([a.speed], dtype=np.float32)})[0]
            audio = np.asarray(audio, dtype=np.float32).reshape(-1)
            # trim leading/trailing near-silence
            nz = np.where(np.abs(audio) > 0.01)[0]
            if len(nz): audio = audio[max(0, nz[0] - 600): nz[-1] + 1200]
            marks.append({"text": line, "start": round(t, 3), "end": round(t + len(audio) / 24000, 3)})
            chunks += [audio, np.zeros(int(a.gap * 24000), np.float32)]
            t += len(audio) / 24000 + a.gap
            print(scene["id"], f"{len(audio)/24000:5.2f}s", ps)
        wav = np.concatenate(chunks[:-1])
        wavfile.write(os.path.join(a.outdir, scene["id"] + ".wav"), 24000, np.clip(wav, -1, 1).astype(np.float32))
        durs[scene["id"]] = {"dur": round(len(wav) / 24000, 3), "lines": marks}
    json.dump(durs, open(os.path.join(a.outdir, "durations.json"), "w"), indent=1, ensure_ascii=False)

if __name__ == "__main__":
    main()
