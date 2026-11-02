"""Build the soundtrack: ambient pad + soft arpeggio + two impacts, voice placed per scene, ducked mix."""
import json, os, numpy as np
from scipy.io import wavfile
from scipy.signal import lfilter, butter, sosfilt

SR = 48000
VO = os.environ.get("VO_DIR", "vo")  # per-scene narration wavs from tts.py
TL = json.load(open("timeline.json"))
T = TL["total"]; N = int(T * SR) + SR
rng = np.random.default_rng(11)
t = np.arange(N) / SR

def hz(m): return 440 * 2 ** ((m - 69) / 12)
# chords (MIDI): Dmaj9, Bm9, Gmaj9, Asus
CH = [[50, 57, 61, 64, 66, 69], [47, 54, 57, 61, 62, 66], [43, 50, 54, 57, 59, 62], [45, 52, 57, 59, 62, 64]]
CL = 10.0  # seconds per chord

def env_chord(i):
    a = i * CL; e = np.clip((t - a + 1.5) / 3.0, 0, 1) * np.clip((a + CL + 1.5 - t) / 3.0, 0, 1)
    return np.sin(e * np.pi / 2) ** 2

pad = np.zeros(N)
nch = int(T // CL) + 2
for i in range(nch):
    ch = CH[i % 4]; e = env_chord(i)
    m = e > 0
    if not m.any(): continue
    tt = t[m]; s = np.zeros(m.sum())
    for k, note in enumerate(ch):
        f = hz(note)
        for det in (-0.004, 0.0, 0.004):
            ph = rng.uniform(0, 2 * np.pi)
            for h in range(1, 5):
                s += np.sin(2 * np.pi * f * h * (1 + det) * tt + ph * h) / (h ** 1.8) / (1 + 0.15 * k)
    # sub root
    s += 0.9 * np.sin(2 * np.pi * hz(ch[0] - 12) * tt)
    pad[m] += s * e[m]
pad *= (0.85 + 0.15 * np.sin(2 * np.pi * 0.11 * t))
sos = butter(2, 1800, "low", fs=SR, output="sos"); pad = sosfilt(sos, pad)
pad /= np.abs(pad).max()

# arpeggio: soft sine plucks on 8ths at 96 bpm, enters after the logo reveal
arp = np.zeros(N); step = 60 / 96 / 2
s1_reveal = TL["scenes"][0]["start"] + 5.45
k = 0; ts = s1_reveal
pat = [0, 2, 3, 4, 3, 2, 5, 3]
while ts < T - 4:
    ch = CH[int(ts // CL) % 4]; note = ch[pat[k % 8] % len(ch)] + 12
    i0 = int(ts * SR); L = int(1.2 * SR); tt = np.arange(L) / SR
    v = (np.sin(2 * np.pi * hz(note) * tt) + 0.25 * np.sin(4 * np.pi * hz(note) * tt)) * np.exp(-tt * 5.5) * (1 - np.exp(-tt * 400))
    acc = 1.0 if k % 4 == 0 else 0.6
    arp[i0:i0 + L] += v[:max(0, min(L, N - i0))] * acc
    k += 1; ts += step
arp *= np.clip((t - s1_reveal) / 4, 0, 1) * np.clip((T - 3.5 - t) / 3, 0, 1)
arp /= max(1e-9, np.abs(arp).max())

def reverb(x, wet=0.35):
    out = np.zeros_like(x)
    for d, g in ((1557, .84), (1617, .83), (1491, .82), (1422, .81), (1277, .80), (1356, .79)):
        d = int(d * SR / 25000); a = np.zeros(d + 1); a[0] = 1; a[d] = -g
        out += lfilter([1], a, x)
    for d in (225, 556):
        d = int(d * SR / 25000); g = .5; b = np.zeros(d + 1); b[0] = -g; b[d] = 1; a = np.zeros(d + 1); a[0] = 1; a[d] = -g
        out = lfilter(b, a, out)
    out /= np.abs(out).max() + 1e-9
    return (1 - wet) * x + wet * out

music = reverb(0.75 * pad + 0.22 * arp, 0.4)

# impacts: riser into the logo reveal, booms at reveal and outro
def boom(at, gain=1.0):
    i0 = int(at * SR); L = int(2.2 * SR); tt = np.arange(L) / SR
    f = 32 + 40 * np.exp(-tt * 6); ph = 2 * np.pi * np.cumsum(f) / SR
    v = np.sin(ph) * np.exp(-tt * 2.2) * (1 - np.exp(-tt * 300))
    n = rng.normal(0, 1, L) * np.exp(-tt * 9) * 0.25
    n = sosfilt(butter(2, 1200, "low", fs=SR, output="sos"), n)
    x = np.zeros(N); x[i0:i0 + L] = (v + n)[:N - i0] * gain; return x
def riser(a, b):
    i0, i1 = int(a * SR), int(b * SR); L = i1 - i0; tt = np.linspace(0, 1, L)
    n = rng.normal(0, 1, L)
    n = sosfilt(butter(2, [800, 6000], "band", fs=SR, output="sos"), n) * tt ** 2.5
    x = np.zeros(N); x[i0:i1] = n / np.abs(n).max(); return x
fx = 0.9 * boom(s1_reveal) + 0.12 * riser(s1_reveal - 1.6, s1_reveal) + 0.55 * boom(TL["scenes"][8]["start"] + 0.1)
fx = reverb(fx, 0.25)

# voice
voice = np.zeros(N)
for sc in TL["scenes"]:
    sr, v = wavfile.read(os.path.join(VO, sc["id"] + ".wav"))
    v = v.astype(np.float64)
    idx = np.round(np.arange(int(len(v) * SR / sr)) * sr / SR).astype(int).clip(0, len(v) - 1)
    v = np.interp(np.arange(int(len(v) * SR / sr)) * sr / SR, np.arange(len(v)), v)
    i0 = int((sc["start"] + sc["pre"]) * SR); voice[i0:i0 + len(v)] += v[:N - i0]
voice = sosfilt(butter(2, 70, "high", fs=SR, output="sos"), voice)
voice /= np.abs(voice).max()

# duck music under voice
envv = np.abs(voice); win = int(0.25 * SR)
envv = np.convolve(envv, np.ones(win) / win, mode="same"); envv /= envv.max()
duck = 1 - 0.55 * np.clip(envv * 4, 0, 1)
fade = np.clip(t / 1.0, 0, 1) * np.clip((T - t) / 3.0, 0, 1)
mix = 0.92 * voice + 0.16 * music * duck * fade + 0.5 * fx * fade
mix = mix[:int(T * SR)]
mix /= np.abs(mix).max() / 0.89
st = np.stack([mix, mix], 1).astype(np.float32)
wavfile.write("soundtrack.wav", SR, st)
print("ok", len(mix) / SR)
