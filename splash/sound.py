"""Звуковой дизайн заставки по правилам Ultimate Editors:
whoosh (−12 dB) на движение с пиком на самом быстром кадре, текстуры (−6…−7 dB) на события,
нарастание + удар (−5 dB) на сборку знака, музыки-подложки под анимацией нет.
Каждый звук синтезируется отдельно (библиотека UE в проект не подключена)."""
import numpy as np, wave
SR = 48000; DUR = 2.7
N = int(SR * DUR)
rng = np.random.default_rng(7)
db = lambda d: 10 ** (d / 20)

def env_peak(n, peak_pos, attack_curve=2.0, release_curve=2.5):
    t = np.linspace(0, 1, n); e = np.where(t < peak_pos, (t / peak_pos) ** attack_curve, ((1 - t) / (1 - peak_pos)) ** release_curve)
    return e

def bandnoise(n, f_lo, f_hi, f_lo2=None, f_hi2=None):
    """Шум с полосой, которая может скользить (FFT-фильтр по кускам)."""
    x = rng.standard_normal(n); out = np.zeros(n); blk = 2048; hop = 1024; win = np.hanning(blk)
    f_lo2 = f_lo if f_lo2 is None else f_lo2; f_hi2 = f_hi if f_hi2 is None else f_hi2
    for s in range(0, n - blk, hop):
        k = s / max(n - blk, 1)
        lo, hi = f_lo + (f_lo2 - f_lo) * k, f_hi + (f_hi2 - f_hi) * k
        X = np.fft.rfft(x[s:s+blk] * win); f = np.fft.rfftfreq(blk, 1/SR)
        m = np.exp(-0.5 * ((np.log(np.maximum(f, 1)) - np.log((lo*hi)**.5)) / (np.log(hi/lo)/2.2 + 1e-9))**2)
        out[s:s+blk] += np.fft.irfft(X * m)
    return out / (np.max(np.abs(out)) + 1e-9)

def whoosh(dur, peak, lo, hi, lo2=None, hi2=None):
    n = int(SR * dur); w = bandnoise(n, lo, hi, lo2, hi2) * env_peak(n, peak)
    return w / np.max(np.abs(w))

def click(dur=0.05, freq=2400, tone=0.0, noise=1.0, decay=60):
    n = int(SR * dur); t = np.arange(n) / SR
    s = noise * bandnoise(n + 2048, freq * .6, freq * 1.6)[:n] + tone * np.sin(2*np.pi*freq*t)
    s *= np.exp(-t * decay); return s / np.max(np.abs(s))

def thump(dur=0.12, f0=180, f1=70, decay=28):
    n = int(SR * dur); t = np.arange(n) / SR
    f = f1 + (f0 - f1) * np.exp(-t * 40); ph = 2*np.pi*np.cumsum(f)/SR
    s = np.sin(ph) * np.exp(-t * decay); return s / np.max(np.abs(s))

def pluck(freq, dur=0.6, bright=0.5):
    """Karplus–Strong: короткий деревянный щипок — тональная текстура на касание пачки."""
    n = int(SR * dur); L = int(SR / freq); buf = rng.uniform(-1, 1, L); out = np.zeros(n)
    for i in range(n):
        out[i] = buf[i % L]; buf[i % L] = 0.996 * (bright * buf[i % L] + (1 - bright) * buf[(i + 1) % L])
    out *= np.exp(-np.arange(n) / SR * 5); return out / np.max(np.abs(out))

def chord(freqs, dur, decay=1.6):
    n = int(SR * dur); t = np.arange(n) / SR; s = np.zeros(n)
    for k, f in enumerate(freqs):
        s += (np.sin(2*np.pi*f*t) + .35*np.sin(4*np.pi*f*t) + .12*np.sin(6*np.pi*f*t)) * np.exp(-t * (decay + k*.2))
    s *= np.minimum(1, t / 0.008); return s / np.max(np.abs(s))

def mx(*parts):
    n = max(len(p) for p, g in parts); o = np.zeros(n)
    for p, g in parts: o[:len(p)] += p * g
    return o

tracks = {k: np.zeros(N) for k in ['WHOOSHES', 'TEXTURES', 'RISERS', 'HITS']}
def place(track, sig, at, level_db, peak_at=None):
    """at — время события; если задан peak_at (доля), пик звука ставится точно на at."""
    off = int((at - (peak_at or 0) * len(sig) / SR) * SR)
    s0 = max(off, 0); sig = sig[s0 - off:]; e = min(N, s0 + len(sig))
    tracks[track][s0:e] += sig[:e - s0] * db(level_db)

# --- PASS 1: WHOOSHES (пик на самом быстром кадре; у ease-in-out — середина движения)
place('WHOOSHES', whoosh(0.45, .5, 1800, 7000), 0.19, -12, peak_at=.5)          # линия прилавка рисуется (0–0.38)
place('WHOOSHES', whoosh(0.55, .5, 500, 2600), 0.50, -12, peak_at=.5)           # ставня вырастает (0.28–0.72)
place('WHOOSHES', whoosh(0.45, .5, 900, 4000), 0.95, -11, peak_at=.5)           # ставня сворачивается вверх (0.78–1.12)
# --- PASS 2: TEXTURES (на кадр события)
for k in range(10):                                                            # трещотка рулона — звук самой ставни
    place('TEXTURES', click(0.03, 3200, decay=120), 0.80 + k * 0.032, -16)
pent = [587.3, 659.3, 740.0, 880.0, 987.8, 1174.7, 1318.5, 1480.0]             # D-мажорная пентатоника вверх — «счёт» пачек
for i in range(8):
    land = 0.98 + i * 0.07 + 0.26
    place('TEXTURES', mx((thump(0.09, 220, 90, 40), .6), (click(0.04, 1600, decay=90), .4)), land, -7)
    place('TEXTURES', pluck(pent[i], 0.5, .55), land, -13)
for k in range(6):                                                             # фестоны навеса раскрываются (лёгкий хлопок ткани)
    place('TEXTURES', click(0.06, 900, decay=55), 1.12 + k * 0.07 + 0.18, -15)
# --- PASS 3: финальный акцент — магазин готов (последняя пачка встала на полку)
last = 0.98 + 7 * 0.07 + 0.26
place('TEXTURES', chord([587.3, 740.0, 880.0, 1174.7], 1.0, 2.6), last + 0.02, -12)   # «success reward»
place('HITS', mx((thump(0.25, 150, 55, 12), .7), (click(0.04, 2600, decay=80), .3)), last + 0.02, -9)
riser = whoosh(0.5, .96, 500, 1600, 1500, 5200)
place('RISERS', riser, last, -10, peak_at=.96)                                       # мягкий подъём к финалу (тише: финал небольшой)
mix = sum(tracks.values())
fade = np.ones(N); fl = int(SR * 0.25); fade[-fl:] = np.linspace(1, 0, fl) ** 2
mix *= fade
peak = np.max(np.abs(mix));
if peak > db(-1): mix *= db(-1) / peak
print('peak dBFS', round(20*np.log10(np.max(np.abs(mix))), 2), {k: round(20*np.log10(np.max(np.abs(v))+1e-9),1) for k,v in tracks.items()})
def write(path, x):
    x16 = (np.clip(x, -1, 1) * 32767).astype(np.int16)
    with wave.open(path, 'w') as w: w.setnchannels(2); w.setsampwidth(2); w.setframerate(SR); w.writeframes(np.repeat(x16[:, None], 2, axis=1).tobytes())
write('mix.wav', mix)
for k, v in tracks.items(): write(f'stem_{k}.wav', v)
