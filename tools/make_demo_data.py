"""Generate a synthetic SAGD workbook for the public demo of the Prod / Temp viewers.

Same layout as the field workbook (one sheet per well, 41 columns, daily rows) but every
value is simulated: fictional pads A/B, made-up well behaviour and events. Nothing is
derived from field data, so the output is safe to publish.

    python3 tools/make_demo_data.py            # -> demo/SAGD_Demo_Data.xlsx
"""
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import Workbook

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "demo" / "SAGD_Demo_Data.xlsx"
SEED = 20261008

START, END = "2019-01-01", "2026-06-30"
DAYS = pd.date_range(START, END, freq="D")
N = len(DAYS)
NPT = 9  # thermocouple points along the lateral, heel -> toe

COLUMNS = [
    "Date", "Long_Tubing_Injection_Rate_t/hr", "Short_Tubing_Injection_Rate_t/hr", "Blanket_Gas_Kg/hr",
    "Intake_Pressure_kPa", "Injection_Pressure_kPa", "Frequency_Hz", "Emulsion_Rate_Am3/h", "Choke_Size_%",
    "WHT_Temp", "Casing_Pressure_kPa", "Casing_Choke_%", "Current_Amp", "Torque_N/m", "Subcool_Min",
    *[f"Subcool_Point_{i}" for i in range(1, NPT + 1)], *[f"Temp_Point_{i}" for i in range(1, NPT + 1)],
    "Temp_Max", "Prod_Oil_rate_bbld", "Prod_Water_rate_bbld", "Inj_Steam_rate_bbld",
    "Oil_Operation", "ESP_Operation", "Redrill_Operation", "Production_Hours",
]

# Field-wide steam outages (boiler maintenance): no steam; producers' ESPs held at ~35 Hz, no production.
FIELD_OUTAGES = [("2021-02-08", "2021-02-12"), ("2023-07-10", "2023-08-24"), ("2025-05-19", "2025-05-23")]

# Fictional wells. start = first steam (circulation); peak = plateau oil (bbl/d); sor = steam/oil at plateau;
# tau = decline time constant (days). Events are optional and drive the demo stories.
WELLS = [
    dict(name="A_01", start="2019-03-04", peak=1350, sor=2.7, tau=4200),
    dict(name="A_02", start="2019-03-06", peak=900, sor=3.2, tau=3000,
         redrill=dict(trouble="2022-06-01", shut="2023-02-20", back="2023-04-15", gain=1.25, hot=6)),
    dict(name="A_03", start="2019-03-11", peak=1100, sor=2.9, tau=3600),
    dict(name="A_04", start="2019-03-15", peak=800, sor=3.4, tau=2600,
         redrill=dict(trouble="2023-10-01", shut="2024-06-03", back="2024-07-29", gain=1.4, hot=3)),
    dict(name="A_05", start="2019-11-18", peak=600, sor=3.8, tau=2800, dead_tc=(7, "2021-09-01", 0.45)),
    dict(name="A_06", start="2020-06-01", peak=500, sor=4.0, tau=2500, esp_decline="2024-02-01"),
    dict(name="B_01", start="2019-04-22", peak=1700, sor=2.5, tau=4500),
    dict(name="B_02", start="2019-04-24", peak=1250, sor=2.8, tau=3800),
    dict(name="B_03", start="2019-04-29", peak=1450, sor=2.6, tau=3300,
         redrill=dict(trouble="2024-05-01", shut="2025-01-27", back="2025-03-24", gain=1.15, hot=8)),
    dict(name="B_04", start="2019-05-02", peak=1000, sor=3.0, tau=3400, dead_tc=(9, "2022-05-10", 1.0)),
    dict(name="B_05", start="2019-05-06", peak=1150, sor=2.9, tau=3500,
         esp_fail=("2022-11-14", "2022-12-28")),
    dict(name="B_06", start="2019-05-13", peak=950, sor=3.1, tau=3000, shut_in="2025-11-01"),
    dict(name="B_07", start="2024-04-15", peak=850, sor=2.8, tau=3000),  # infill
    dict(name="B_08", start="2024-10-21", peak=750, sor=2.9, tau=3000),  # infill
]

CIRC_DAYS = 85         # steam circulation before conversion to SAGD
BBL_PER_M3 = 6.2898
T_PER_BBL_CWE = 0.158987


def tsat(p_kpa):
    """Saturation temperature (°C) of water at p (kPa), table interpolation over the SAGD range."""
    p = np.array([500, 1000, 1500, 2000, 2500, 3000, 3500, 4000, 4500, 5000, 6000, 7000])
    t = np.array([151.8, 179.9, 198.3, 212.4, 223.9, 233.9, 242.6, 250.4, 257.4, 263.9, 275.6, 285.8])
    return np.interp(p_kpa, p, t)


def ar1(rng, n, sigma, phi=0.9):
    """Smooth multiplicative-style noise: AR(1) with stationary std `sigma`."""
    e = rng.normal(0, sigma * np.sqrt(1 - phi * phi), n)
    out = np.empty(n)
    acc = rng.normal(0, sigma)
    for i in range(n):
        acc = phi * acc + e[i]
        out[i] = acc
    return out


def steps(rng, n, every=(45, 140), size=0.08):
    """Piecewise-constant operator set-point changes around 1.0."""
    out = np.ones(n)
    i, level = 0, 1.0
    while i < n:
        k = int(rng.integers(*every))
        level = float(np.clip(level * np.exp(rng.normal(0, size)), 0.75, 1.25))
        out[i:i + k] = level
        i += k
    return out


def idx(date):
    return int((pd.Timestamp(date) - DAYS[0]).days)


def ramp(n0, n1, n, a=0.0, b=1.0):
    """0 before n0, linear a->b between n0 and n1, b after."""
    x = np.arange(n, dtype=float)
    return np.clip((x - n0) / max(1, n1 - n0), 0, 1) * (b - a) + a


def simulate(w, rng):
    d = np.arange(N)
    s0 = idx(w["start"])
    conv = s0 + CIRC_DAYS
    active = d >= s0
    sagd = d >= conv
    t = np.maximum(0, d - conv).astype(float)

    # ---- availability: uptime hours, steam on/off ----
    hours = np.where(sagd, 24.0, 0.0)
    steam_on = active.copy()
    trip_rate = 7 / 365
    for i in np.flatnonzero(sagd & (rng.random(N) < trip_rate)):
        dur = int(rng.integers(1, 4))
        hours[i] = rng.uniform(4, 20)
        hours[i + 1:i + dur] = 0
    outage = np.zeros(N, bool)
    for a, b in FIELD_OUTAGES:
        outage[idx(a):idx(b) + 1] = True
    outage &= active
    steam_on &= ~outage
    hours[outage & sagd] = 0

    # ---- reservoir response ----
    rise = 1 - np.exp(-t / 210)
    decline = np.exp(-t / w["tau"])
    perf = np.ones(N)          # well performance multiplier (events)
    esp_down = np.zeros(N, bool)
    redrill = np.zeros(N)
    hot_point, hot_level = None, np.zeros(N)

    if "redrill" in w:
        r = w["redrill"]
        i0, i1, i2 = idx(r["trouble"]), idx(r["shut"]), idx(r["back"])
        perf *= ramp(i0, i1, N, 1.0, 0.45) * (d < i2) + (d >= i2) * ramp(i2, i2 + 75, N, 0.6, r["gain"])
        # deteriorating uptime before the redrill
        bad = (d >= i0) & (d < i1) & (rng.random(N) < ramp(i0, i1, N, 0.01, 0.12))
        hours[bad] = rng.uniform(0, 12, bad.sum())
        hours[i1:i2] = 0
        steam_on[i1:i2] = False
        esp_down[i1:i2] = True
        redrill[i2:] = 1
        hot_point = r["hot"] - 1
        hot_level = ramp(i0, i1, N, 0, 1) * (d < i1)
    if "esp_fail" in w:
        a, b = idx(w["esp_fail"][0]), idx(w["esp_fail"][1])
        hours[a:b] = 0
        esp_down[a:b] = True
    if "shut_in" in w:
        a = idx(w["shut_in"])
        hours[a:] = 0
        esp_down[a:] = True
        steam_on[a:] = False

    up = hours / 24
    set_steam = steps(rng, N)
    plateau_steam = w["peak"] * w["sor"]
    steam = plateau_steam * np.where(sagd, (0.35 + 0.65 * rise) * decline ** 0.8, 0.55) * set_steam * np.exp(ar1(rng, N, 0.05))
    steam = np.where(steam_on, steam, 0)
    # producing days follow the injector set-point with a lag
    steam_eff = pd.Series(set_steam).rolling(30, min_periods=1).mean().to_numpy()
    oil = w["peak"] * rise * decline * perf * steam_eff ** 0.7 * np.exp(ar1(rng, N, 0.07))
    oil = np.where(sagd, oil * up, 0)
    oil[outage] = 0
    wsr = 0.92 + 0.12 * np.tanh(t / 900) + ar1(rng, N, 0.04)
    water = np.where(sagd, plateau_steam * (0.35 + 0.65 * rise) * decline ** 0.8 * steam_eff * wsr * up, 0)
    water[outage] = 0
    circ = active & ~sagd
    water[circ] = steam[circ] * rng.uniform(0.55, 0.75, circ.sum())

    # ---- ESP ----
    emul = (oil + water) / BBL_PER_M3 / 24 / np.maximum(up, 1e-9) * (1.08 + ar1(rng, N, 0.02))
    emul = np.where(hours > 0, emul, 0)
    q_ref = w["peak"] * (1 + w["sor"]) / BBL_PER_M3 / 24
    eff = np.ones(N)
    if "esp_decline" in w:
        eff = ramp(idx(w["esp_decline"]), N, N, 1.0, 0.55)
    freq = np.clip(28 + 24 * emul / (q_ref * eff), 30, 60) + ar1(rng, N, 0.4)
    hold = outage & sagd & ~esp_down     # ESPs held at ~35 Hz through field outages
    freq = np.where(hold, 35 + rng.normal(0, 0.3, N), freq)
    freq = np.where((hours > 0) | hold, freq, 0)
    freq[esp_down | ~sagd] = 0
    amps = np.where(freq > 0, 14 + 0.023 * freq ** 2 * (0.75 + 0.25 * emul / max(q_ref, 1)) / eff ** 0.5, 0)
    amps += np.where(freq > 0, ar1(rng, N, 1.2), 0)
    torque = np.where(freq > 0, amps / 98 + ar1(rng, N, 0.015), 0)

    # ---- pressures ----
    p_res = 3300 + rng.uniform(-350, 350) + 250 * np.sin(d / 365 * 2 * np.pi * 0.35 + rng.uniform(0, 6)) + ar1(rng, N, 40)
    p_res = np.where(steam_on, p_res, np.nan)
    p_res = pd.Series(p_res).ffill().fillna(1500).to_numpy()
    cool = pd.Series(np.where(steam_on, 0.0, 1.0)).rolling(25, min_periods=1).mean().to_numpy()
    intake = np.where(active, p_res * (1 - 0.35 * cool), rng.uniform(0, 30, N))
    inj = np.where(steam_on, intake + 280 + ar1(rng, N, 25), 0)
    casing = np.where(active, intake - 950 + ar1(rng, N, 60), 0)
    choke = np.where(hours > 0, np.clip(45 * steps(rng, N, size=0.15) + ar1(rng, N, 2), 5, 100), 0)
    casing_choke = np.where(active & (rng.random(N) < 0.04), rng.uniform(5, 40, N), 0)
    blanket = np.where(steam_on & (w["name"] in ("A_01", "B_01", "B_02")), 5 + ar1(rng, N, 0.6), 0)

    # ---- thermocouples ----
    t_sat = tsat(intake)
    sc_base = rng.uniform(18, 50, NPT) + np.linspace(0, 12, NPT) * rng.choice([-1, 1])
    sc_drift = np.stack([ar1(rng, N, 5, phi=0.995) for _ in range(NPT)], 1)
    sc_target = np.clip(sc_base + sc_drift, 2, 90)
    if hot_point is not None:  # steam coning toward one point before the redrill
        sc_target[:, hot_point] = sc_target[:, hot_point] * (1 - hot_level) + 1.5 * hot_level
    T_amb = 12.0
    T = np.full((N, NPT), T_amb)
    cur = np.full(NPT, T_amb)
    for i in range(N):
        if steam_on[i]:
            target, a = t_sat[i] - sc_target[i], (0.06 if not sagd[i] else 0.25)
        elif active[i]:
            target, a = 90.0, 0.02
        else:
            target, a = T_amb, 0.5
        cur = cur + a * (target - cur)
        T[i] = cur
    T += rng.normal(0, 0.6, T.shape)
    sub = np.where(active[:, None], t_sat[:, None] - T, np.nan)  # no subcool before first steam
    sub = np.clip(sub, -1, None)

    if "dead_tc" in w:  # dead thermocouple logs 0 °C; Subcool then shows Tsat (a known data quirk)
        pt, since, frac = w["dead_tc"]
        dead = (d >= idx(since)) & (rng.random(N) < frac)
        if frac < 1:  # intermittent: dead in runs, not single days
            dead = pd.Series(dead).rolling(9, center=True, min_periods=1).max().astype(bool).to_numpy() & (d >= idx(since))
        T[dead, pt - 1] = 0
        sub[dead, pt - 1] = t_sat[dead]

    valid = T > 1
    t_max = np.where(valid.any(1), np.where(valid, T, -np.inf).max(1), 0)
    sub_min = np.where(active, np.where(valid, sub, np.inf).min(1), np.nan)
    sub_min[np.isinf(sub_min)] = np.nan
    wht = np.where(hours > 0, np.clip(t_max - 25 + ar1(rng, N, 3), 20, None),
                   12 + 10 * np.sin((d - 100) / 365 * 2 * np.pi) + rng.normal(0, 1.2, N))

    lt = steam * T_PER_BBL_CWE / 24 * (0.5 + 0.12 * np.tanh(ar1(rng, N, 0.6, 0.99)))
    st = steam * T_PER_BBL_CWE / 24 - lt

    cols = {
        "Long_Tubing_Injection_Rate_t/hr": lt, "Short_Tubing_Injection_Rate_t/hr": st,
        "Blanket_Gas_Kg/hr": np.clip(blanket, 0, None), "Intake_Pressure_kPa": np.clip(intake, 0, None),
        "Injection_Pressure_kPa": inj, "Frequency_Hz": np.clip(freq, 0, None), "Emulsion_Rate_Am3/h": emul,
        "Choke_Size_%": choke, "WHT_Temp": wht, "Casing_Pressure_kPa": np.clip(casing, 0, None),
        "Casing_Choke_%": casing_choke, "Current_Amp": np.clip(amps, 0, None),
        "Torque_N/m": np.clip(torque, 0, None), "Subcool_Min": sub_min,
        **{f"Subcool_Point_{i + 1}": sub[:, i] for i in range(NPT)},
        **{f"Temp_Point_{i + 1}": T[:, i] for i in range(NPT)},
        "Temp_Max": t_max, "Prod_Oil_rate_bbld": oil, "Prod_Water_rate_bbld": water, "Inj_Steam_rate_bbld": steam,
        "Oil_Operation": (oil > 0).astype(int), "ESP_Operation": (freq > 0).astype(int),
        "Redrill_Operation": redrill.astype(int), "Production_Hours": hours,
    }
    df = pd.DataFrame({"Date": DAYS, **cols})[COLUMNS]
    # infill wells: sheet starts a few weeks before first steam, like the field workbook
    first = max(0, s0 - 21) if s0 > idx("2023-01-01") else 0
    return df.iloc[first:].reset_index(drop=True)


def main():
    rng = np.random.default_rng(SEED)
    OUT.parent.mkdir(exist_ok=True)
    wb = Workbook(write_only=True)
    for w in WELLS:
        df = simulate(w, rng)
        ws = wb.create_sheet(w["name"])
        ws.append(COLUMNS)
        num = df.columns[1:]
        vals = df[num].to_numpy(dtype=float).round(2)
        for date, row in zip(df["Date"].dt.to_pydatetime(), vals):
            ws.append([date, *[None if v != v else int(v) if v == int(v) else float(v) for v in row]])
        print(f"{w['name']}: {len(df)} days, cum oil {df.Prod_Oil_rate_bbld.sum() / 1e6:.2f} MMbbl, "
              f"CSOR {df.Inj_Steam_rate_bbld.sum() / max(1, df.Prod_Oil_rate_bbld.sum()):.2f}")
    wb.save(OUT)
    print(f"wrote {OUT} ({OUT.stat().st_size / 1048576:.1f} MB)")


if __name__ == "__main__":
    main()
