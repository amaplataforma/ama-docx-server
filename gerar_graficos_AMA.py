"""
Gráficos clínicos v4 — Plataforma AMA / Oxy Recovery
Baseado em v3. Mudança única: altura dos gráficos reduzida de 11.11 para 7.78 polegadas
(~17.6 cm), permitindo que texto + gráfico convivam na mesma página A4 sem espaço em
branco excessivo. G5 usa figsize=(7.94, 7.78) por ter largura diferente (1430px vs 1452px).
"""
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.lines import Line2D
from scipy.interpolate import UnivariateSpline
import warnings
warnings.filterwarnings('ignore')

# ─── PALETA ────────────────────────────────────────────────────────────────
C_CHO    = '#5B9BD5'
C_GORD   = '#ED7D31'
C_FC     = '#8B0000'
C_VO2    = '#D62728'
C_VCO2   = '#1F9ED4'
C_RER    = '#70AD47'
C_VE_VO2 = '#ED7D31'
C_VE_VCO2= '#5B9BD5'
C_GRID   = '#E8E8E8'
C_LABEL  = '#999999'
Z_COLORS = {'Z1':'#E3F0FF','Z2':'#D4EDDA','Z3':'#FFF9C4','Z4':'#FFE0B2','Z5':'#FFCCBC'}
Z_ALPHA  = 0.50

# ─── PARÂMETROS CLÍNICOS ───────────────────────────────────────────────────
p25 = dict(L1_vel=14.3, L1_fc=136, RER1_vel=15.0, L2_vel=20.0, L2_fc=159,
           VO2pico_vel=22.5, FCmax=175, FATmax_vel=13.5, FATmax_val=0.27)
p26 = dict(L1_vel=15.0, L1_fc=129, RER1_vel=18.1, L2_vel=21.9, L2_fc=160,
           VO2pico_vel=24.0, FCmax=171, FATmax_vel=13.1, FATmax_val=0.34)

# ─── DADOS 2025 ────────────────────────────────────────────────────────────
data25_raw = [
    ('PRE', 54, 0.62,0.48,17.04,10.33,8,   27.48,35.5,0.7742,0.3269,0.2007,0),
    ('PRE', 56, 0.38,0.30,11.00, 6.36,5.02,28.30,35.9,0.7895,0.1866,0.1496,0),
    ('PRE', 57, 0.42,0.34,12.60, 6.98,5.60,30.00,37.5,0.8095,0.1863,0.2040,0),
    ('PRE', 63, 0.37,0.30,11.10, 6.14,4.97,30.00,37.0,0.8108,0.1630,0.1819,0),
    ('EXE', 88, 0.73,0.58,18.1,12.2, 9.69,25.3,32.0,0.7945,0.3497,0.3043, 4.6),
    ('EXE', 91, 0.92,0.73,21.7,15.3,12.1, 23.8,30.1,0.7935,0.4431,0.3791, 5.3),
    ('EXE', 86, 1.03,0.86,24.7,17.2,14.3, 24.5,30.2,0.8350,0.3947,0.6203, 6.0),
    ('EXE',103, 0.88,0.69,21.0,14.7,11.5, 23.9,30.8,0.7841,0.4434,0.3248, 6.8),
    ('EXE',111, 1.26,1.04,28.3,21.0,17.3, 22.5,27.5,0.8254,0.5114,0.7036, 7.6),
    ('EXE',109, 1.23,1.04,28.5,20.6,17.3, 23.2,27.8,0.8455,0.4405,0.8004, 8.3),
    ('EXE',111, 1.73,1.54,39.1,28.8,25.7, 22.7,25.4,0.8902,0.4363,1.4799, 9.0),
    ('EXE',114, 1.68,1.53,38.4,28.1,25.5, 22.8,25.1,0.9107,0.3418,1.5954, 9.8),
    ('EXE',114, 1.77,1.60,40.0,29.4,26.7, 22.7,25.0,0.9040,0.3885,1.6260,10.5),
    ('EXE',120, 1.90,1.74,43.3,31.7,29.1, 22.8,25.0,0.9158,0.3637,1.8485,11.3),
    ('EXE',124, 2.04,1.90,46.5,33.9,31.7, 22.8,24.5,0.9314,0.3151,2.1305,12.0),
    ('EXE',131, 1.91,1.74,43.4,31.8,29.0, 22.7,25.0,0.9110,0.3873,1.8162,12.8),
    ('EXE',134, 2.27,2.20,53.2,37.8,36.7, 23.4,24.2,0.9692,0.1471,2.7640,13.5),
    ('EXE',136, 2.11,1.94,47.9,35.1,32.4, 22.7,24.7,0.9194,0.3856,2.0880,14.3),
    ('EXE',140, 2.32,2.34,55.3,38.7,39.0, 23.8,23.7,1.0086,0.0,   3.2446,15.0),
    ('EXE',146, 2.44,2.45,59.0,40.8,40.9, 24.1,24.0,1.0041,0.0,   3.3618,15.8),
    ('EXE',148, 2.23,2.25,53.1,37.2,36.4, 23.6,24.4,1.0090,0.0,   3.1223,16.5),
    ('EXE',151, 2.68,2.76,67.4,44.7,46.0, 25.1,24.4,1.0299,0.0,   4.0089,17.3),
    ('EXE',154, 2.79,2.94,74.1,46.6,49.0, 26.4,25.1,1.0538,0.0,   4.4794,18.0),
    ('EXE',157, 2.86,3.03,75.7,47.7,50.6, 26.4,25.0,1.0594,0.0,   4.6662,18.8),
    ('EXE',160, 2.74,2.82,69.9,45.7,47.0, 25.5,24.7,1.0292,0.0,   4.0905,19.5),
    ('EXE',159, 2.75,2.88,71.0,45.9,47.9, 25.8,24.7,1.0473,0.0,   4.3333,20.0),
    ('EXE',160, 2.97,3.20,83.3,49.5,53.3, 27.9,26.0,1.0774,0.0,   5.0908,20.3),
    ('EXE',162, 2.85,2.97,74.4,47.5,49.6, 26.1,25.0,1.0421,0.0,   4.4234,21.0),
    ('EXE',167, 3.04,3.32,86.3,50.6,55.3, 28.4,26.0,1.0921,0.0,   5.4152,21.8),
    ('EXE',175, 3.32,3.69,108.3,56.3,62.5,32.6,29.3,1.1114,0.0,   6.2083,22.5),
    ('REC',175, 3.40,3.73,106.4,57.6,63.2,31.3,28.5,1.0971,0.0,   6.1337, 2.0),
    ('REC',164, 2.87,3.18, 79.5,48.6,53.9,27.7,25.0,1.1080,0.0,   5.3217, 2.0),
    ('REC',126, 2.34,2.27, 54.5,39.7,38.5,23.3,24.0,0.9701,0.1465,2.8591, 3.0),
    ('REC',110, 2.19,2.03, 49.9,37.1,34.4,22.8,24.6,0.9269,0.3613,2.2426, 3.0),
    ('REC', 94, 1.35,1.15, 30.4,22.9,19.5,22.5,26.4,0.8519,0.4632,0.9177, 0.0),
    ('REC', 92, 1.07,0.87, 24.4,18.1,14.8,22.8,28.1,0.8131,0.4655,0.5371, 0.0),
]

data26_raw = [
    ('PRE', 58, 0.10,0.08, 3.12, 1.69, 1.36,31.2,39.0,0.8000,0.0466,0.0442, 1.8),
    ('EXE', 62, 0.21,0.18, 6.24, 3.62, 3.05,29.3,34.7,0.8571,0.0694,0.1478, 4.3),
    ('EXE', 77, 0.44,0.35,13.30, 7.46, 5.99,29.5,36.2,0.7955,0.2098,0.1853, 5.1),
    ('EXE', 81, 0.68,0.55,17.00,11.50, 9.26,26.5,32.8,0.8088,0.3027,0.3281, 6.2),
    ('EXE', 88, 0.73,0.57,18.80,12.40, 9.66,26.2,33.6,0.7808,0.3735,0.2585, 7.3),
    ('EXE', 97, 0.97,0.76,22.70,16.40,12.90,23.7,30.2,0.7835,0.4901,0.3554, 8.1),
    ('EXE',104, 1.05,0.84,24.30,17.90,14.20,23.1,28.8,0.8000,0.4894,0.4641, 9.0),
    ('EXE',106, 1.17,0.94,26.60,19.80,16.00,22.8,28.2,0.8034,0.5359,0.5355, 9.8),
    ('EXE',110, 1.42,1.19,31.60,24.00,20.20,22.5,26.5,0.8380,0.5338,0.8752,10.6),
    ('EXE',116, 1.71,1.53,38.80,29.00,25.90,22.6,25.4,0.8947,0.4127,1.4986,11.4),
    ('EXE',117, 1.83,1.66,41.10,31.00,28.10,22.6,24.8,0.9071,0.3880,1.7075,12.2),
    ('EXE',119, 1.96,1.81,44.50,33.20,30.70,22.7,24.6,0.9235,0.3395,1.9759,13.1),
    ('EXE',121, 1.85,1.70,42.10,31.40,28.90,22.8,24.8,0.9189,0.3404,1.8264,13.9),
    ('EXE',126, 1.86,1.72,42.40,31.60,29.10,22.7,24.8,0.9247,0.3166,1.8858,14.7),
    ('EXE',129, 1.86,1.71,42.20,31.50,29.00,22.7,24.8,0.9194,0.3403,1.8400,15.0),
    ('EXE',133, 2.13,1.99,48.70,36.00,33.80,22.9,24.4,0.9343,0.3143,2.2528,15.6),
    ('EXE',136, 2.24,2.11,51.50,38.00,35.70,23.0,24.4,0.9420,0.2897,2.4481,16.4),
    ('EXE',139, 2.31,2.29,54.40,39.20,38.80,23.6,23.8,0.9913,0.0281,3.0476,17.2),
    ('EXE',142, 2.42,2.43,58.20,41.10,41.20,24.0,24.2,1.0041,0.0,   3.3346,18.1),
    ('EXE',147, 2.66,2.72,66.40,45.10,46.20,25.0,24.4,1.0226,0.0,   3.8900,18.9),
    ('EXE',152, 2.72,2.80,68.30,46.00,47.50,25.1,24.4,1.0294,0.0,   4.0633,19.7),
    ('EXE',155, 2.78,2.89,70.80,47.10,48.90,25.5,24.6,1.0396,0.0,   4.2824,20.5),
    ('EXE',157, 2.75,2.89,71.20,46.60,49.00,25.9,24.6,1.0509,0.0,   4.3792,21.4),
    ('EXE',160, 2.76,2.90,71.60,46.80,49.20,25.9,24.6,1.0507,0.0,   4.3927,21.9),
    ('EXE',161, 2.88,3.06,76.70,48.80,51.90,26.6,25.0,1.0625,0.0,   4.7392,22.2),
    ('EXE',164, 2.98,3.24,82.90,54.60,54.90,27.6,25.5,1.0872,0.0,   5.2419,23.0),
    ('EXE',168, 3.02,3.26,84.80,64.10,55.20,27.9,26.0,1.0795,0.0,   5.2046,23.8),
    ('EXE',171, 3.15,3.36,89.80,53.40,57.00,28.6,26.6,1.0667,0.0,   5.2437,24.0),
    ('REC',172, 3.23,3.47,93.70,54.70,58.80,29.2,27.0,1.0743,0.0,   5.4900, 2.0),
    ('REC',174, 3.22,3.51,95.30,54.50,59.50,29.4,27.1,1.0901,0.0,   5.7056, 2.0),
    ('REC',160, 2.78,2.91,72.50,47.10,49.30,26.0,24.9,1.0468,0.0,   4.3741, 3.0),
    ('REC',133, 2.35,2.31,55.50,39.80,39.10,23.6,24.4,0.9830,0.0752,3.0103, 3.0),
    ('REC',101, 1.90,1.77,43.60,32.30,29.90,22.8,24.7,0.9316,0.2925,1.9861, 0.0),
    ('REC', 90, 1.58,1.36,35.60,26.80,23.10,22.4,26.2,0.8608,0.5087,1.1385, 0.0),
    ('REC', 86, 1.28,1.04,28.90,21.70,17.70,22.6,27.7,0.8125,0.5587,0.6391, 0.0),
]

def parse_data(raw):
    return dict(
        stages  = [r[0] for r in raw],
        FC      = np.array([r[1]  for r in raw], float),
        VO2_L   = np.array([r[2]  for r in raw], float),
        VCO2_L  = np.array([r[3]  for r in raw], float),
        VE_VO2  = np.array([r[7]  for r in raw], float),
        VE_VCO2 = np.array([r[8]  for r in raw], float),
        RER     = np.array([r[9]  for r in raw], float),
        OxiG    = np.array([r[10] for r in raw], float),
        OxiC    = np.array([r[11] for r in raw], float),
        Vel     = np.array([r[12] for r in raw], float),
    )

d25 = parse_data(data25_raw)
d26 = parse_data(data26_raw)

def get_idx(data, stage):
    return np.array([i for i,s in enumerate(data['stages']) if s==stage])

def build_x(data, pico_vel, step=0.8):
    stages = data['stages']
    vel    = data['Vel']
    n      = len(stages)
    x      = np.zeros(n)
    pre = [i for i,s in enumerate(stages) if s=='PRE']
    exe = [i for i,s in enumerate(stages) if s=='EXE']
    rec = [i for i,s in enumerate(stages) if s=='REC']
    x_exe_start = vel[exe[0]]
    for k, i in enumerate(pre):
        x[i] = x_exe_start - (len(pre) - k) * step
    for i in exe:
        x[i] = vel[i]
    for k, i in enumerate(rec):
        x[i] = pico_vel + (k+1) * step
    return x

x25 = build_x(d25, p25['VO2pico_vel'])
x26 = build_x(d26, p26['VO2pico_vel'])

# ─── EIXO X UNIFICADO ──────────────────────────────────────────────────────
XLIM = (min(x25.min(), x26.min()) - 0.6,
        max(x25.max(), x26.max()) + 0.6)

EXE_TICK_VELS = np.arange(4, 25, 2)

def get_unified_ticks(x25, x26, d25, d26):
    ticks  = []
    labels = []
    pre25 = get_idx(d25,'PRE'); pre26 = get_idx(d26,'PRE')
    rec25 = get_idx(d25,'REC'); rec26 = get_idx(d26,'REC')
    pre_x = np.mean([np.mean(x25[pre25]), np.mean(x26[pre26])])
    ticks.append(pre_x); labels.append('PRE')
    for v in EXE_TICK_VELS:
        ticks.append(v)
        labels.append(str(int(v)))
    rec_x = np.mean([np.mean(x25[rec25]), np.mean(x26[rec26])])
    ticks.append(rec_x); labels.append('REC')
    return ticks, labels

UNIFIED_TICKS, UNIFIED_LABELS = get_unified_ticks(x25, x26, d25, d26)

def set_unified_xticks(ax):
    ax.set_xticks(UNIFIED_TICKS)
    ax.set_xticklabels(UNIFIED_LABELS, fontsize=7.5)
    ax.set_xlim(XLIM)

def smooth_spline(xp, yp, xf, k=5, s_factor=5.0):
    if len(xp) < k+1:
        return np.interp(xf, xp, yp)
    spl = UnivariateSpline(xp, yp, k=k, s=len(xp)*s_factor)
    return spl(xf)

def add_zones(ax, p, exe_start_vel):
    L1, RER1, L2, VO2p = p['L1_vel'], p['RER1_vel'], p['L2_vel'], p['VO2pico_vel']
    z1_start = exe_start_vel
    zones = [
        ('Z1', z1_start,    0.70*L1),
        ('Z2', 0.70*L1,     L1),
        ('Z3', L1,          RER1),
        ('Z4', RER1,        L2),
        ('Z5', L2,          VO2p),
    ]
    for zname, za, zb in zones:
        if zb > za:
            ax.axvspan(za, zb, color=Z_COLORS[zname], alpha=Z_ALPHA, zorder=0)
            mid = (za+zb)/2
            ax.text(mid, 0.93, zname, transform=ax.get_xaxis_transform(),
                    ha='center', va='top', fontsize=7, color='#555', fontweight='bold')

def add_threshold_lines(ax, p, show_rer1=True):
    thresholds = [('L1', p['L1_vel']), ('L2', p['L2_vel'])]
    if show_rer1:
        thresholds.append(('QR=1', p['RER1_vel']))
    placed = []
    for label, vel in sorted(thresholds, key=lambda t: t[1]):
        ax.axvline(vel, color=C_GORD, lw=1.6, zorder=3)
        offset = 0
        for pv in placed:
            if abs(vel - pv) < 1.5:
                offset += 0.06
        ax.text(vel+0.15, 0.97-offset, label,
                transform=ax.get_xaxis_transform(),
                ha='left', va='top', fontsize=7.5,
                color=C_GORD, fontweight='bold')
        placed.append(vel)

def style_ax(ax):
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.yaxis.grid(True, color=C_GRID, lw=0.8, zorder=0)
    ax.set_axisbelow(True)
    ax.tick_params(labelsize=8)

def pre_rec_labels(ax, x_all, data):
    pre_idx = get_idx(data,'PRE'); rec_idx = get_idx(data,'REC')
    if len(pre_idx):
        ax.text(np.mean(x_all[pre_idx]), 0.04, 'Repouso',
                transform=ax.get_xaxis_transform(),
                ha='center', style='italic', color=C_LABEL, fontsize=8)
    if len(rec_idx):
        ax.text(np.mean(x_all[rec_idx]), 0.04, 'Recuperação',
                transform=ax.get_xaxis_transform(),
                ha='center', style='italic', color=C_LABEL, fontsize=8)

def build_fatmax_bell(x_all, exe_idx, p, data):
    x_exe = x_all[exe_idx]
    g_exe = data['OxiG'][exe_idx]
    spl   = UnivariateSpline(x_exe, g_exe, k=5, s=len(x_exe)*5.0)
    xf    = np.linspace(x_exe[0], x_exe[-1], 500)
    gf    = np.maximum(spl(xf), 0)
    peak_x = xf[np.argmax(gf)]
    shift  = p['FATmax_vel'] - peak_x
    gf_shifted = np.interp(xf, xf + shift, gf)
    gf_shifted = np.maximum(gf_shifted, 0)
    if gf_shifted.max() > 0:
        gf_shifted = gf_shifted * p['FATmax_val'] / gf_shifted.max()
    return xf, gf_shifted

# ══════════════════════════════════════════════════════════════════════════
# G1 — MAPA METABÓLICO
# figsize v4: (8.07, 7.78) — altura reduzida de 11.11 → 7.78 polegadas (~17.6 cm)
# ══════════════════════════════════════════════════════════════════════════
def plot_g1(data, x_all, p, ax, title):
    exe_idx = get_idx(data,'EXE')
    exe_start_vel = data['Vel'][exe_idx[0]]
    set_unified_xticks(ax)
    add_zones(ax, p, exe_start_vel)
    ax2 = ax.twinx()
    xf = np.linspace(x_all.min(), x_all.max(), 600)
    cho_sm = np.maximum(smooth_spline(x_all, data['OxiC'], xf), 0)
    ax.scatter(x_all, data['OxiC'], color=C_CHO, s=45, edgecolors='white', lw=1.1, zorder=4, alpha=0.85)
    ax.plot(xf, cho_sm, color=C_CHO, lw=2, zorder=3)
    xb, gb = build_fatmax_bell(x_all, exe_idx, p, data)
    ax2.scatter(x_all, data['OxiG'], color=C_GORD, s=45, edgecolors='white', lw=1.1, zorder=4, alpha=0.85)
    ax2.plot(xb, gb, color=C_GORD, lw=2, zorder=3)
    add_threshold_lines(ax, p, show_rer1=True)
    style_ax(ax); ax2.spines['top'].set_visible(False); ax2.tick_params(labelsize=8)
    pre_rec_labels(ax, x_all, data)
    ax.set_ylabel('Oxidação de CHO (g/min)', color=C_CHO, fontsize=9)
    ax2.set_ylabel('Oxidação de Gordura (g/min)', color=C_GORD, fontsize=9)
    ax.tick_params(axis='y', colors=C_CHO)
    ax2.tick_params(axis='y', colors=C_GORD)
    ax.set_title(title, fontsize=10, fontweight='bold', loc='left', pad=8)
    return ax2

fig1, (ax1a, ax1b) = plt.subplots(2,1, figsize=(8.07, 7.78), constrained_layout=True)
fig1.patch.set_facecolor('white')
r1a = plot_g1(d25, x25, p25, ax1a, 'Mar / 2025')
r1b = plot_g1(d26, x26, p26, ax1b, 'Jan / 2026')
cho_max = max(ax1a.get_ylim()[1], ax1b.get_ylim()[1])
grd_max = max(r1a.get_ylim()[1], r1b.get_ylim()[1])
for ax_, r_ in [(ax1a,r1a),(ax1b,r1b)]:
    ax_.set_ylim(0, cho_max*1.05); r_.set_ylim(0, grd_max*1.15)
leg1 = [mpatches.Patch(color=C_CHO,  label='Oxidação de CHO (g/min)'),
        mpatches.Patch(color=C_GORD, label='Oxidação de Gordura (g/min)')]
for ax_ in [ax1a, ax1b]:
    ax_.legend(handles=leg1, loc='upper center', bbox_to_anchor=(0.5,1.14), ncol=2, fontsize=8, frameon=False)
fig1.suptitle('Mapa Metabólico Modificado — Oxidação de Substratos', fontsize=12, fontweight='bold', y=1.01)
ax1b.set_xlabel('Velocidade (km/h)', fontsize=9)
fig1.savefig('/mnt/user-data/outputs/G1_mapa_metabolico_v4.png', dpi=180, bbox_inches='tight', facecolor='white')
plt.close(fig1); print('G1 ✓')

# ══════════════════════════════════════════════════════════════════════════
# G2 — RESPOSTA CARDÍACA
# figsize v4: (8.07, 7.78)
# ══════════════════════════════════════════════════════════════════════════
FC_YMAX = 220

def plot_g2(data, x_all, p, ax, title):
    exe_idx = get_idx(data,'EXE')
    exe_start_vel = data['Vel'][exe_idx[0]]
    set_unified_xticks(ax)
    add_zones(ax, p, exe_start_vel)
    ax2 = ax.twinx()
    xf = np.linspace(x_all.min(), x_all.max(), 600)
    fc_sm  = smooth_spline(x_all, data['FC'],    xf)
    vo2_sm = np.maximum(smooth_spline(x_all, data['VO2_L'],  xf), 0)
    vco2_sm= np.maximum(smooth_spline(x_all, data['VCO2_L'], xf), 0)
    ax.plot(xf, fc_sm, color=C_FC, lw=2, zorder=3)
    ax.scatter(x_all, data['FC'], color=C_FC, s=45, edgecolors='white', lw=1.1, zorder=4)
    ax2.plot(xf, vo2_sm,  color=C_VO2,  lw=1.8, ls='--', zorder=3)
    ax2.plot(xf, vco2_sm, color=C_VCO2, lw=1.8, ls='--', zorder=3)
    ax.plot(p['L1_vel'], p['L1_fc'], marker='^', color=C_GORD, ms=10, zorder=5, lw=0)
    ax.plot(p['L2_vel'], p['L2_fc'], marker='v', color=C_GORD, ms=10, zorder=5, lw=0)
    ax.plot(p['VO2pico_vel'], p['FCmax'], marker='D', color='#555', ms=8, zorder=5, lw=0)
    add_threshold_lines(ax, p, show_rer1=False)
    style_ax(ax); ax2.spines['top'].set_visible(False); ax2.tick_params(labelsize=8)
    ax.set_ylim(0, FC_YMAX)
    gas_max = max(data['VO2_L'].max(), data['VCO2_L'].max())
    ax2.set_ylim(0, gas_max * 2.2)
    pre_rec_labels(ax, x_all, data)
    ax.set_ylabel('Frequência Cardíaca (bpm)', color=C_FC, fontsize=9)
    ax2.set_ylabel('VO₂ / VCO₂ (L/min)', fontsize=9)
    ax.tick_params(axis='y', colors=C_FC)
    ax.set_title(title, fontsize=10, fontweight='bold', loc='left', pad=8)
    return ax2

fig2, (ax2a, ax2b) = plt.subplots(2,1, figsize=(8.07, 7.78), constrained_layout=True)
fig2.patch.set_facecolor('white')
r2a = plot_g2(d25, x25, p25, ax2a, 'Mar / 2025')
r2b = plot_g2(d26, x26, p26, ax2b, 'Jan / 2026')
gas_max2 = max(r2a.get_ylim()[1], r2b.get_ylim()[1])
r2a.set_ylim(0, gas_max2); r2b.set_ylim(0, gas_max2)
def leg2(p_):
    return [Line2D([0],[0], color=C_FC,   lw=2,   marker='o', ms=5,  label='FC (bpm)'),
            Line2D([0],[0], color=C_VCO2, lw=1.8, ls='--',           label='VCO₂ (L/min)'),
            Line2D([0],[0], color=C_VO2,  lw=1.8, ls='--',           label='VO₂ (L/min)'),
            Line2D([0],[0], color=C_GORD, marker='^', ms=8, lw=0,    label=f'L1 — {p_["L1_vel"]} km/h · {p_["L1_fc"]} bpm'),
            Line2D([0],[0], color=C_GORD, marker='v', ms=8, lw=0,    label=f'L2 — {p_["L2_vel"]} km/h · {p_["L2_fc"]} bpm'),
            Line2D([0],[0], color='#555', marker='D', ms=7, lw=0,    label=f'FCmáx — {p_["FCmax"]} bpm')]
ax2a.legend(handles=leg2(p25), loc='upper center', bbox_to_anchor=(0.5,1.20), ncol=3, fontsize=7.5, frameon=False)
ax2b.legend(handles=leg2(p26), loc='upper center', bbox_to_anchor=(0.5,1.20), ncol=3, fontsize=7.5, frameon=False)
fig2.suptitle('Resposta Cardíaca e Cinética de Gases ao Exercício', fontsize=12, fontweight='bold', y=1.01)
ax2b.set_xlabel('Velocidade (km/h)', fontsize=9)
fig2.savefig('/mnt/user-data/outputs/G2_resposta_cardiaca_v4.png', dpi=180, bbox_inches='tight', facecolor='white')
plt.close(fig2); print('G2 ✓')

# ══════════════════════════════════════════════════════════════════════════
# G3 — EQUIVALENTES VENTILATÓRIOS
# figsize v4: (8.07, 7.78)
# ══════════════════════════════════════════════════════════════════════════
def plot_g3(data, x_all, p, ax, title):
    set_unified_xticks(ax)
    ax.set_ylim(18, 42); ax.set_yticks(range(18,43,4))
    xf = np.linspace(x_all.min(), x_all.max(), 600)
    ve_vo2_sm  = smooth_spline(x_all, data['VE_VO2'],  xf)
    ve_vco2_sm = smooth_spline(x_all, data['VE_VCO2'], xf)
    ax.scatter(x_all, data['VE_VO2'],  color=C_VE_VO2,  s=45, edgecolors='white', lw=1.1, zorder=4)
    ax.plot(xf, ve_vo2_sm,  color=C_VE_VO2,  lw=2, zorder=3)
    ax.scatter(x_all, data['VE_VCO2'], color=C_VE_VCO2, s=45, edgecolors='white', lw=1.1, zorder=4)
    ax.plot(xf, ve_vco2_sm, color=C_VE_VCO2, lw=2, ls='--', zorder=3)
    for label, vel in [('L1',p['L1_vel']),('L2',p['L2_vel'])]:
        ax.axvline(vel, color=C_GORD, lw=1.6, zorder=3)
        ax.text(vel+0.15, 0.97, label, transform=ax.get_xaxis_transform(),
                ha='left', va='top', fontsize=8, color=C_GORD, fontweight='bold')
    style_ax(ax)
    pre_rec_labels(ax, x_all, data)
    ax.set_ylabel('Equivalentes Ventilatórios', fontsize=9)
    ax.set_title(title, fontsize=10, fontweight='bold', loc='left', pad=8)

fig3, (ax3a, ax3b) = plt.subplots(2,1, figsize=(8.07, 7.78), constrained_layout=True)
fig3.patch.set_facecolor('white')
plot_g3(d25, x25, p25, ax3a, 'Mar / 2025')
plot_g3(d26, x26, p26, ax3b, 'Jan / 2026')
leg3 = [Line2D([0],[0], color=C_VE_VO2,  lw=2, marker='o', ms=5,        label='VE/VO₂'),
        Line2D([0],[0], color=C_VE_VCO2, lw=2, ls='--', marker='o', ms=5, label='VE/VCO₂')]
for ax_ in [ax3a, ax3b]:
    ax_.legend(handles=leg3, loc='upper center', bbox_to_anchor=(0.5,1.14), ncol=2, fontsize=8, frameon=False)
fig3.suptitle('Equivalentes Ventilatórios  (VE/VO₂  e  VE/VCO₂)', fontsize=12, fontweight='bold', y=1.01)
ax3b.set_xlabel('Velocidade (km/h)', fontsize=9)
fig3.savefig('/mnt/user-data/outputs/G3_equivalentes_ventilatorios_v4.png', dpi=180, bbox_inches='tight', facecolor='white')
plt.close(fig3); print('G3 ✓')

# ══════════════════════════════════════════════════════════════════════════
# G4 — RER
# figsize v4: (8.07, 7.78)
# ══════════════════════════════════════════════════════════════════════════
def plot_g4(data, x_all, p, ax, title):
    set_unified_xticks(ax)
    ax.set_ylim(0.70, 1.25)
    ax.axvspan(p['L1_vel'], p['RER1_vel'], color='#D4EDDA', alpha=0.60, zorder=1)
    ax.axhline(1.0, color='#AAAAAA', lw=1.2, ls='--', zorder=2)
    ax.text(XLIM[1]-0.5, 1.01, 'RER = 1,0', ha='right', va='bottom', fontsize=8, color='#888')
    xf = np.linspace(x_all.min(), x_all.max(), 600)
    rer_sm = smooth_spline(x_all, data['RER'], xf)
    ax.scatter(x_all, data['RER'], color=C_RER, s=45, edgecolors='white', lw=1.1, zorder=4)
    ax.plot(xf, rer_sm, color=C_RER, lw=2, zorder=3)
    placed = []
    for label, vel in sorted([('L1',p['L1_vel']),('QR=1',p['RER1_vel']),('L2',p['L2_vel'])],
                               key=lambda t: t[1]):
        ax.axvline(vel, color=C_GORD, lw=1.6, zorder=3)
        off = sum(0.06 for pv in placed if abs(vel-pv)<1.5)
        ax.text(vel+0.15, 0.97-off, label, transform=ax.get_xaxis_transform(),
                ha='left', va='top', fontsize=7.5, color=C_GORD, fontweight='bold')
        placed.append(vel)
    style_ax(ax)
    pre_rec_labels(ax, x_all, data)
    ax.set_ylabel('RER (VCO₂/VO₂)', fontsize=9)
    ax.set_title(title, fontsize=10, fontweight='bold', loc='left', pad=8)
    janela = p['RER1_vel'] - p['L1_vel']
    return janela

fig4, (ax4a, ax4b) = plt.subplots(2,1, figsize=(8.07, 7.78), constrained_layout=True)
fig4.patch.set_facecolor('white')
j25 = plot_g4(d25, x25, p25, ax4a, 'Mar / 2025')
j26 = plot_g4(d26, x26, p26, ax4b, 'Jan / 2026')
def leg4(j):
    return [Line2D([0],[0], color=C_RER, lw=2, marker='o', ms=5, label='RER (VCO₂/VO₂)'),
            mpatches.Patch(color='#D4EDDA', alpha=0.7, label=f'Janela aeróbia  {j:.1f} km/h')]
ax4a.legend(handles=leg4(j25), loc='upper center', bbox_to_anchor=(0.5,1.14), ncol=2, fontsize=8, frameon=False)
ax4b.legend(handles=leg4(j26), loc='upper center', bbox_to_anchor=(0.5,1.14), ncol=2, fontsize=8, frameon=False)
fig4.suptitle('Relação de Troca Respiratória (RER)', fontsize=12, fontweight='bold', y=1.01)
ax4b.set_xlabel('Velocidade (km/h)', fontsize=9)
fig4.savefig('/mnt/user-data/outputs/G4_rer_v4.png', dpi=180, bbox_inches='tight', facecolor='white')
plt.close(fig4); print('G4 ✓')

# ══════════════════════════════════════════════════════════════════════════
# G5 — FATMAX ISOLADO
# figsize v4: (7.94, 7.78) — largura reduzida 8.07→7.94 (proporcional 1430px vs 1452px)
# ══════════════════════════════════════════════════════════════════════════
def plot_g5(data, x_all, p, ax, title):
    exe_idx = get_idx(data,'EXE')
    set_unified_xticks(ax)
    ax.axvspan(p['FATmax_vel']*0.92, p['FATmax_vel']*1.08, color='#FFE0CC', alpha=0.70, zorder=1)
    ax.axvline(p['RER1_vel'], color='#AAAAAA', lw=1.2, ls='--', zorder=2)
    ax.text(p['RER1_vel']+0.15, 0.97, f'QR=1  {p["RER1_vel"]} km/h',
            transform=ax.get_xaxis_transform(), ha='left', va='top', fontsize=7.5, color='#888')
    xb, gb = build_fatmax_bell(x_all, exe_idx, p, data)
    ax.scatter(x_all, data['OxiG'], color=C_GORD, s=45, edgecolors='white', lw=1.1, zorder=4, alpha=0.85)
    ax.plot(xb, gb, color=C_GORD, lw=2.5, zorder=3)
    ax.plot(p['FATmax_vel'], p['FATmax_val'], marker='D', color=C_GORD,
            ms=10, zorder=5, markeredgecolor='white', markeredgewidth=1.2, lw=0)
    ax.text(0.54, 0.80, f'MFO = {p["FATmax_val"]:.2f} g/min @ {p["FATmax_vel"]} km/h',
            transform=ax.transAxes, fontsize=9, color=C_GORD, fontweight='bold')
    style_ax(ax)
    ax.set_ylim(bottom=0)
    pre_rec_labels(ax, x_all, data)
    ax.set_ylabel('Oxidação de Gordura (g/min)', color=C_GORD, fontsize=9)
    ax.tick_params(axis='y', colors=C_GORD)
    ax.set_title(title, fontsize=10, fontweight='bold', loc='left', pad=8)

fig5, (ax5a, ax5b) = plt.subplots(2,1, figsize=(7.94, 7.78), constrained_layout=True)
fig5.patch.set_facecolor('white')
plot_g5(d25, x25, p25, ax5a, 'Mar / 2025')
plot_g5(d26, x26, p26, ax5b, 'Jan / 2026')
ymax5 = max(ax5a.get_ylim()[1], ax5b.get_ylim()[1])
ax5a.set_ylim(0, ymax5); ax5b.set_ylim(0, ymax5)
def leg5(p_):
    return [Line2D([0],[0], color=C_GORD, lw=2, marker='o', ms=5, label='Oxidação de Gordura (g/min)'),
            Line2D([0],[0], color=C_GORD, marker='D', ms=8, lw=0,
                   label=f'MFO — {p_["FATmax_val"]:.2f} g/min @ {p_["FATmax_vel"]} km/h'),
            mpatches.Patch(color='#FFE0CC', alpha=0.8,
                   label=f'Zona FATmax  ({p_["FATmax_vel"]*0.92:.1f}–{p_["FATmax_vel"]*1.08:.1f} km/h)')]
ax5a.legend(handles=leg5(p25), loc='upper center', bbox_to_anchor=(0.5,1.16), ncol=2, fontsize=8, frameon=False)
ax5b.legend(handles=leg5(p26), loc='upper center', bbox_to_anchor=(0.5,1.16), ncol=2, fontsize=8, frameon=False)
fig5.suptitle('FATmax — Pico de Oxidação de Gordura', fontsize=12, fontweight='bold', y=1.01)
ax5b.set_xlabel('Velocidade (km/h)', fontsize=9)
fig5.text(0.5, -0.01,
          f'Evolução MFO:  {p25["FATmax_val"]:.2f} → {p26["FATmax_val"]:.2f} g/min   (Δ = +{(p26["FATmax_val"]/p25["FATmax_val"]-1)*100:.0f}%)',
          ha='center', fontsize=9, fontweight='bold')
fig5.savefig('/mnt/user-data/outputs/G5_fatmax_v4.png', dpi=180, bbox_inches='tight', facecolor='white')
plt.close(fig5); print('G5 ✓')

print('\n✅ 5 gráficos gerados — v4 (altura padronizada ~17.6 cm).')
