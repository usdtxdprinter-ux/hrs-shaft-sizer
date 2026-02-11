"""
HRS Exhaust Shaft Sizing Calculator — Streamlit Chatbot
=========================================================
A chatbot-style application for sizing fire-rated exhaust shafts
in high-rise buildings using the LF Systems HRS constant pressure system.

Products: DEF (Dryer Exhaust Fan), DBF (Dryer Booster Fan), L150/L152 controllers
Website:  https://www.lfsystems.net
System:   HRS (High Rise Shaft)

Engineering Basis:
  - Darcy-Weisbach friction loss: Δpf = f*(L/Dh)*ρ*(V/1096.2)²
  - Colebrook friction factor for turbulent flow
  - Huebscher equivalent diameter for rectangular ducts
  - ASHRAE 2009 Duct Design Chapter 21 fitting loss coefficients
  - Subduct area deductions: 4"→15 sq.in., 6"→31.5 sq.in., 8"→54 sq.in.

Deploy:   pip install streamlit pandas plotly
          streamlit run hrs_shaft_sizer.py
"""

import streamlit as st
import math
import pandas as pd
import json

# ─────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────
AIR_DENSITY = 0.075          # lb/ft³ at standard conditions
ROUGHNESS   = 0.0003         # ft — galvanized steel, medium-smooth
KIN_VISC    = 1.63e-4        # ft²/s — air at ~70 °F

SUBDUCT_AREA = {4: 15.0, 6: 31.5, 8: 54.0}  # sq.in. removed per penetration

ROUND_SIZES = [8,10,12,14,16,18,20,22,24,26,28,30,32,34,36,38,40,42,44,46,48]
RECT_SIDES  = [6,8,10,12,14,16,18,20,22,24,26,28,30,32,34,36,38,40,42,44,46,48]

# Fitting loss coefficients (ASHRAE basis)
K_ELBOW_90      = 0.60   # 90° elbow in shaft offset
K_ELBOW_45      = 0.30   # 45° elbow
K_TEE_BRANCH    = 0.50   # tee branch entry from subduct
K_EXIT          = 1.00   # abrupt exit / fan entry
K_ENTRY_BELL    = 0.03   # bellmouth entry at bottom of shaft
K_ENTRY_ABRUPT  = 0.50   # abrupt entry

# ─────────────────────────────────────────────
# ENGINEERING FUNCTIONS
# ─────────────────────────────────────────────
def colebrook_friction_factor(dh_ft: float, velocity_fpm: float) -> float:
    """Iterative Colebrook equation for Darcy friction factor."""
    v_fps = velocity_fpm / 60.0
    Re = v_fps * dh_ft / KIN_VISC
    if Re < 1.0:
        return 0.0
    if Re < 2300:
        return 64.0 / Re
    f = 0.02  # initial guess
    for _ in range(80):
        rhs = -2.0 * math.log10(ROUGHNESS / (3.7 * dh_ft) + 2.51 / (Re * math.sqrt(f)))
        f_new = 1.0 / (rhs ** 2)
        if abs(f_new - f) < 1e-10:
            break
        f = f_new
    return f

def velocity_pressure(velocity_fpm: float) -> float:
    """Velocity pressure in inches of water column."""
    return AIR_DENSITY * (velocity_fpm / 1096.2) ** 2

def darcy_pressure_drop(length_ft: float, dh_in: float, sum_k: float, velocity_fpm: float) -> float:
    """
    Combined friction + fitting loss:
        Δp = [f*(12L/Dh_in) + ΣK] * ρ * (V/1096.2)²
    Returns in. WC.
    """
    if dh_in <= 0 or velocity_fpm <= 0:
        return 0.0
    dh_ft = dh_in / 12.0
    f = colebrook_friction_factor(dh_ft, velocity_fpm)
    friction_term = f * (length_ft / dh_ft)
    vp = velocity_pressure(velocity_fpm)
    return (friction_term + sum_k) * vp

def huebscher_equiv_diameter(a_in: float, b_in: float) -> float:
    """Circular equivalent diameter for a rectangular duct (Huebscher)."""
    return 1.30 * (a_in * b_in) ** 0.625 / (a_in + b_in) ** 0.25

def hydraulic_diameter_rect(a_in: float, b_in: float) -> float:
    """Hydraulic diameter Dh = 4A/P for rectangular duct."""
    area = a_in * b_in
    perim = 2.0 * (a_in + b_in)
    return 4.0 * area / perim

def circular_area(d_in: float) -> float:
    """Cross-sectional area of round duct in sq. inches."""
    return math.pi * (d_in / 2.0) ** 2


# ─────────────────────────────────────────────
# SHAFT SIZING ENGINE
# ─────────────────────────────────────────────
def size_shaft(params: dict) -> dict:
    """
    Main sizing calculation.
    Returns dict with best_result, alternatives list, and per-floor data.
    """
    floors           = params["floors"]
    floor_data       = params["floor_data"]       # list of dicts per floor
    floor_height     = params["floor_height"]
    duct_after_last  = params["duct_after_last"]
    diversity_pct    = params["diversity_pct"]
    max_delta_p      = params["max_delta_p"]
    shape_choice     = params["shape_choice"]
    user_diameter    = params.get("user_diameter", 0)
    user_rect_a      = params.get("user_rect_a", 0)
    user_rect_b      = params.get("user_rect_b", 0)
    offset_elbows    = params.get("offset_elbows", 0)
    offset_length    = params.get("offset_length", 0)
    offset_angle     = params.get("offset_angle", 90)

    # ── aggregate floor data ──
    total_cfm = 0
    total_pens = 0
    max_subduct_area_on_any_floor = 0
    floor_cumulative_cfm = []

    for fd in floor_data:
        pens = fd["penetrations"]
        cfm_each = fd["cfm_per_pen"]
        sub_size = fd["subduct_size"]
        total_pens += pens
        total_cfm += pens * cfm_each
        sub_area = pens * SUBDUCT_AREA[sub_size]
        if sub_area > max_subduct_area_on_any_floor:
            max_subduct_area_on_any_floor = sub_area

    design_cfm = total_cfm * diversity_pct / 100.0
    total_height = floors * floor_height

    # ── offset fitting losses ──
    k_offset = 0.0
    if offset_elbows > 0:
        k_per = K_ELBOW_90 if offset_angle >= 60 else K_ELBOW_45
        k_offset = offset_elbows * k_per

    # ── try a specific shaft size ──
    def evaluate(shaft_area_sqin, dh_in, label, is_round, dim_a, dim_b):
        eff_area = shaft_area_sqin - max_subduct_area_on_any_floor
        if eff_area <= 0:
            return None
        eff_area_sqft = eff_area / 144.0
        vel = design_cfm / eff_area_sqft  # fpm

        if vel < 50:
            return None

        vp = velocity_pressure(vel)

        # shaft friction (vertical run through building)
        dp_shaft = darcy_pressure_drop(total_height, dh_in, 0, vel)

        # duct after last unit
        dp_after = darcy_pressure_drop(duct_after_last, dh_in, 0, vel) if duct_after_last > 0 else 0.0

        # offset section
        dp_offset = darcy_pressure_drop(offset_length, dh_in, k_offset, vel) if (offset_elbows > 0) else 0.0

        # entry loss (bellmouth assumed at shaft base)
        dp_entry = K_ENTRY_BELL * vp

        # exit / fan entry loss
        dp_exit = K_EXIT * vp

        # total system
        dp_total = dp_shaft + dp_after + dp_offset + dp_entry + dp_exit

        # ── floor-by-floor balance ──
        # Bottom floor (floor 1): air must travel full shaft height
        dp_bottom = dp_shaft + dp_after + dp_offset + dp_exit
        # Top floor: air only travels ~1 floor height + after-last + offset
        dp_top_shaft = darcy_pressure_drop(floor_height, dh_in, 0, vel)
        dp_top = dp_top_shaft + dp_after + dp_offset + dp_exit

        delta_p = abs(dp_bottom - dp_top)

        # per-floor breakdown for detailed report
        floor_dp_list = []
        for fi in range(floors):
            floors_above = floors - fi  # floor 1 (bottom) → all floors above
            h = floors_above * floor_height
            dp_fl = darcy_pressure_drop(h, dh_in, 0, vel) + dp_after + dp_offset + dp_exit
            floor_dp_list.append(round(dp_fl, 5))

        return {
            "label":        label,
            "is_round":     is_round,
            "dim_a":        dim_a,
            "dim_b":        dim_b,
            "shaft_area":   round(shaft_area_sqin, 2),
            "eff_area":     round(eff_area, 2),
            "dh_in":        round(dh_in, 2),
            "velocity":     round(vel, 0),
            "vp":           round(vp, 5),
            "dp_shaft":     round(dp_shaft, 5),
            "dp_after":     round(dp_after, 5),
            "dp_offset":    round(dp_offset, 5),
            "dp_entry":     round(dp_entry, 5),
            "dp_exit":      round(dp_exit, 5),
            "dp_total":     round(dp_total, 5),
            "dp_bottom":    round(dp_bottom, 5),
            "dp_top":       round(dp_top, 5),
            "delta_p":      round(delta_p, 5),
            "passes":       delta_p <= max_delta_p,
            "total_cfm":    total_cfm,
            "design_cfm":   round(design_cfm, 0),
            "total_pens":   total_pens,
            "total_height": total_height,
            "floor_dp":     floor_dp_list,
        }

    # ── run sizing ──
    results = []

    if shape_choice in ("round_auto", "rect_auto"):
        if shape_choice == "round_auto":
            for d in ROUND_SIZES:
                area = circular_area(d)
                r = evaluate(area, d, f'{d}" Round', True, d, d)
                if r and 100 < r["velocity"] < 3000:
                    results.append(r)
        else:
            seen = set()
            for a in RECT_SIDES:
                for b in RECT_SIDES:
                    if b > a:
                        continue
                    if a / b > 4:
                        continue
                    key = (a, b)
                    if key in seen:
                        continue
                    seen.add(key)
                    area = a * b
                    dh = hydraulic_diameter_rect(a, b)
                    r = evaluate(area, dh, f'{a}" × {b}" Rect', False, a, b)
                    if r and 100 < r["velocity"] < 3000:
                        results.append(r)
        results.sort(key=lambda x: x["shaft_area"])
        best = next((r for r in results if r["passes"]), results[-1] if results else None)
        alts = [r for r in results if r["passes"]][:8]
    elif shape_choice == "round_user":
        area = circular_area(user_diameter)
        best = evaluate(area, user_diameter, f'{user_diameter}" Round', True, user_diameter, user_diameter)
        alts = [best] if best else []
    elif shape_choice == "rect_user":
        a, b = max(user_rect_a, user_rect_b), min(user_rect_a, user_rect_b)
        area = a * b
        dh = hydraulic_diameter_rect(a, b)
        best = evaluate(area, dh, f'{a}" × {b}" Rect', False, a, b)
        alts = [best] if best else []
    else:
        best = None
        alts = []

    return {"best": best, "alternatives": alts}


# ─────────────────────────────────────────────
# STREAMLIT APP
# ─────────────────────────────────────────────
def init_state():
    """Initialize session state for the chatbot."""
    defaults = {
        "step":             0,
        "messages":         [],
        "exhaust_type":     "",
        "floors":           0,
        "floor_data":       [],
        "same_all":         True,
        "floor_height":     0.0,
        "duct_after_last":  0.0,
        "diversity_pct":    100.0,
        "has_offset":       False,
        "offset_elbows":    0,
        "offset_length":    0.0,
        "offset_angle":     90,
        "shape_choice":     "",
        "user_diameter":    0,
        "user_rect_a":      0,
        "user_rect_b":      0,
        "max_delta_p":      0.25,
        "current_floor":    0,
        "awaiting":         "",
        "result":           None,
        "calc_done":        False,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


def add_msg(role: str, text: str):
    st.session_state.messages.append({"role": role, "text": text})


def bot(text: str):
    add_msg("assistant", text)


def user(text: str):
    add_msg("user", text)


def reset():
    for k in list(st.session_state.keys()):
        del st.session_state[k]
    init_state()
    st.session_state.step = 0


# ─── Step functions ─────────────────────────
def step_welcome():
    bot(
        "👋 **Welcome to the HRS Exhaust Shaft Sizing Calculator!**\n\n"
        "This tool sizes fire-rated exhaust shafts in high-rise buildings "
        "using the **LF Systems HRS** constant pressure system.\n\n"
        "**Products:** DEF · DBF · L150/L152 controllers  \n"
        "**Website:** [lfsystems.net](https://www.lfsystems.net)\n\n"
        "---\n"
        "Let's get started! **What type of exhaust does this shaft serve?**"
    )
    st.session_state.step = 1


def process_input(user_input: str):
    """Route user input to the correct handler based on current step."""
    val = user_input.strip()
    lc = val.lower()
    step = st.session_state.step

    # ─── Step 1: Exhaust type ───
    if step == 1:
        user(val)
        if "dryer" in lc:
            st.session_state.exhaust_type = "Clothes Dryers"
        elif "bath" in lc:
            st.session_state.exhaust_type = "Bathroom Exhaust"
        elif "kitchen" in lc or "hood" in lc:
            st.session_state.exhaust_type = "Kitchen Hood Exhaust"
        else:
            bot("⚠️ Please select **Dryers**, **Bathrooms**, or **Kitchen Hoods**.")
            return
        bot(f"✅ **{st.session_state.exhaust_type}** selected.\n\n"
            f"**How many floors have penetrations into this shaft?**")
        st.session_state.step = 2

    # ─── Step 2: Number of floors ───
    elif step == 2:
        user(val)
        try:
            n = int(val)
            assert 1 <= n <= 120
        except:
            bot("⚠️ Enter a number between 1 and 120.")
            return
        st.session_state.floors = n
        st.session_state.floor_data = [
            {"penetrations": 1, "subduct_size": 4, "cfm_per_pen": 0} for _ in range(n)
        ]
        bot(f"✅ **{n} floors**.\n\n"
            "Are **all floors configured the same?** (same # of penetrations, subduct size, CFM)")
        st.session_state.step = 3

    # ─── Step 3: Same for all? ───
    elif step == 3:
        user(val)
        if lc in ("yes", "y", "true", "1"):
            st.session_state.same_all = True
            bot("Great — all floors the same.\n\n"
                "**How many penetrations per floor?** (1 or 2)")
            st.session_state.step = 4
            st.session_state.awaiting = "pens"
        else:
            st.session_state.same_all = False
            st.session_state.current_floor = 0
            bot(f"OK — per-floor config.\n\n"
                f"**Floor 1 of {st.session_state.floors}: How many penetrations?** (1 or 2)")
            st.session_state.step = 4
            st.session_state.awaiting = "pens"

    # ─── Step 4: Floor data (pens → subduct → cfm) ───
    elif step == 4:
        user(val)
        aw = st.session_state.awaiting

        if aw == "pens":
            try:
                n = int(val)
                assert n in (1, 2)
            except:
                bot("⚠️ Enter **1** or **2** penetrations per floor.")
                return
            if st.session_state.same_all:
                for fd in st.session_state.floor_data:
                    fd["penetrations"] = n
            else:
                st.session_state.floor_data[st.session_state.current_floor]["penetrations"] = n
            bot(f"✅ {n} penetration(s).\n\n"
                "**What subduct size?** (4, 6, or 8 inches)\n\n"
                "| Size | Area Removed |\n|---|---|\n"
                '| 4" | 15.0 sq.in. |\n| 6" | 31.5 sq.in. |\n| 8" | 54.0 sq.in. |')
            st.session_state.awaiting = "sub"

        elif aw == "sub":
            try:
                n = int(val)
                assert n in (4, 6, 8)
            except:
                bot("⚠️ Subduct must be **4**, **6**, or **8** inches.")
                return
            if st.session_state.same_all:
                for fd in st.session_state.floor_data:
                    fd["subduct_size"] = n
            else:
                st.session_state.floor_data[st.session_state.current_floor]["subduct_size"] = n
            bot(f'✅ {n}" subduct (removes {SUBDUCT_AREA[n]} sq.in.).\n\n'
                "**How many CFM per penetration?**")
            st.session_state.awaiting = "cfm"

        elif aw == "cfm":
            try:
                c = float(val)
                assert c > 0
            except:
                bot("⚠️ Enter a CFM value greater than 0.")
                return
            if st.session_state.same_all:
                for fd in st.session_state.floor_data:
                    fd["cfm_per_pen"] = c
                bot(f"✅ {c} CFM/penetration applied to all {st.session_state.floors} floors.\n\n"
                    "**What is the floor-to-floor height (ft)?**")
                st.session_state.step = 5
            else:
                st.session_state.floor_data[st.session_state.current_floor]["cfm_per_pen"] = c
                cf = st.session_state.current_floor
                if cf < st.session_state.floors - 1:
                    st.session_state.current_floor = cf + 1
                    nf = cf + 2
                    bot(f"✅ Floor {cf+1} done.\n\n"
                        f"**Floor {nf} of {st.session_state.floors}: How many penetrations?** (1 or 2)")
                    st.session_state.awaiting = "pens"
                else:
                    bot(f"✅ All {st.session_state.floors} floors configured!\n\n"
                        "**What is the floor-to-floor height (ft)?**")
                    st.session_state.step = 5

    # ─── Step 5: Floor height ───
    elif step == 5:
        user(val)
        try:
            h = float(val)
            assert h > 0
        except:
            bot("⚠️ Enter a valid height in feet.")
            return
        st.session_state.floor_height = h
        bot(f"✅ {h} ft floor-to-floor.\n\n"
            "**Length of duct from the last (top) floor penetration to the exhaust fan (ft)?**\n"
            "(Include all horizontal/vertical run after the highest connection.)")
        st.session_state.step = 6

    # ─── Step 6: Duct after last ───
    elif step == 6:
        user(val)
        try:
            d = float(val)
            assert d >= 0
        except:
            bot("⚠️ Enter 0 or a positive number of feet.")
            return
        st.session_state.duct_after_last = d
        bot(f"✅ {d} ft after last unit.\n\n"
            "**What is the diversity factor?**  \n"
            "Enter a percentage from 20 to 100.  \n"
            "(e.g., 50 = only 50% of connections active simultaneously)")
        st.session_state.step = 7

    # ─── Step 7: Diversity ───
    elif step == 7:
        user(val)
        try:
            dv = float(val.replace("%", ""))
            assert 20 <= dv <= 100
        except:
            bot("⚠️ Diversity must be between 20 and 100%.")
            return
        st.session_state.diversity_pct = dv
        bot(f"✅ {dv}% diversity.\n\n"
            "**Does the shaft offset after the last floor?**  \n"
            "(The shaft must be straight between floors, but can offset above the top floor.)")
        st.session_state.step = 8

    # ─── Step 8: Offset? ───
    elif step == 8:
        user(val)
        if lc in ("yes", "y", "true", "1"):
            st.session_state.has_offset = True
            bot("**How many elbows in the offset?** (typically 2)")
            st.session_state.step = 9
            st.session_state.awaiting = "elbows"
        else:
            st.session_state.has_offset = False
            st.session_state.offset_elbows = 0
            st.session_state.offset_length = 0
            bot("✅ No offset.\n\n"
                "**How would you like to size the shaft?**\n\n"
                "Choose one:\n"
                "- **round_auto** — find optimal round diameter\n"
                "- **rect_auto** — find optimal rectangular size\n"
                "- **round_user** — I'll specify a diameter\n"
                "- **rect_user** — I'll specify rectangular dims")
            st.session_state.step = 10

    # ─── Step 9: Offset details ───
    elif step == 9:
        user(val)
        aw = st.session_state.awaiting
        if aw == "elbows":
            try:
                n = int(val)
                assert n >= 1
            except:
                bot("⚠️ Enter number of elbows (≥ 1).")
                return
            st.session_state.offset_elbows = n
            bot(f"✅ {n} elbow(s).\n\n**Total length of the offset section (ft)?**")
            st.session_state.awaiting = "olen"
        elif aw == "olen":
            try:
                ol = float(val)
                assert ol >= 0
            except:
                bot("⚠️ Enter a length ≥ 0.")
                return
            st.session_state.offset_length = ol
            bot(f"✅ {ol} ft offset.\n\n"
                "**Elbow angle?** (Enter 45 or 90; default 90)")
            st.session_state.awaiting = "oang"
        elif aw == "oang":
            try:
                ang = int(val)
                assert ang in (45, 90)
            except:
                ang = 90
            st.session_state.offset_angle = ang
            bot(f"✅ {ang}° elbows.\n\n"
                "**How would you like to size the shaft?**\n\n"
                "- **round_auto** — find optimal round diameter\n"
                "- **rect_auto** — find optimal rectangular size\n"
                "- **round_user** — I'll specify a diameter\n"
                "- **rect_user** — I'll specify rectangular dims")
            st.session_state.step = 10

    # ─── Step 10: Shape choice ───
    elif step == 10:
        user(val)
        if lc in ("round_auto", "rect_auto", "round_user", "rect_user"):
            st.session_state.shape_choice = lc
            if lc == "round_user":
                bot("**Enter round duct diameter (inches):**")
                st.session_state.step = 11
                st.session_state.awaiting = "diam"
            elif lc == "rect_user":
                bot("**Enter rectangular dimensions as `width x height` (inches):**\n"
                    "(e.g., 24 x 18)")
                st.session_state.step = 11
                st.session_state.awaiting = "rect"
            else:
                bot("**Maximum allowable ΔP between bottom & top floors?**  \n"
                    "Max = 0.25 in. WC.  Enter your target (e.g., 0.20):")
                st.session_state.step = 12
        else:
            bot("⚠️ Choose: **round_auto**, **rect_auto**, **round_user**, or **rect_user**.")

    # ─── Step 11: User size ───
    elif step == 11:
        user(val)
        aw = st.session_state.awaiting
        if aw == "diam":
            try:
                d = float(val)
                assert 6 <= d <= 60
            except:
                bot("⚠️ Diameter must be 6–60 inches.")
                return
            st.session_state.user_diameter = d
        elif aw == "rect":
            import re
            parts = re.split(r'[x×,\s]+', val)
            try:
                a, b = float(parts[0]), float(parts[1])
                assert a >= 6 and b >= 6
            except:
                bot("⚠️ Enter two dimensions ≥ 6\", e.g. `24 x 18`.")
                return
            st.session_state.user_rect_a = max(a, b)
            st.session_state.user_rect_b = min(a, b)
        bot("**Maximum allowable ΔP between bottom & top floors?**  \n"
            "Max = 0.25 in. WC.  Enter your target:")
        st.session_state.step = 12

    # ─── Step 12: Max ΔP → run calculation ───
    elif step == 12:
        user(val)
        try:
            dp = float(val)
            assert 0.01 <= dp <= 0.25
        except:
            bot("⚠️ Enter a value between 0.01 and 0.25 in. WC.")
            return
        st.session_state.max_delta_p = dp

        # ── Build params and run ──
        params = {
            "floors":          st.session_state.floors,
            "floor_data":      st.session_state.floor_data,
            "floor_height":    st.session_state.floor_height,
            "duct_after_last": st.session_state.duct_after_last,
            "diversity_pct":   st.session_state.diversity_pct,
            "max_delta_p":     dp,
            "shape_choice":    st.session_state.shape_choice,
            "user_diameter":   st.session_state.user_diameter,
            "user_rect_a":     st.session_state.user_rect_a,
            "user_rect_b":     st.session_state.user_rect_b,
            "offset_elbows":   st.session_state.offset_elbows,
            "offset_length":   st.session_state.offset_length,
            "offset_angle":    st.session_state.offset_angle,
        }
        result = size_shaft(params)
        st.session_state.result = result
        st.session_state.calc_done = True
        st.session_state.step = 13

        if result["best"] is None:
            bot("❌ **No valid shaft size found.**\n\n"
                "The CFM may be too high or the area deductions too large for available sizes. "
                "Try adjusting your inputs.\n\nType **restart** to begin again.")
        else:
            bot("✅ **Calculation complete!** See the results below. ⬇️")

    # ─── Step 13: Post-result ───
    elif step == 13:
        user(val)
        if "restart" in lc or "new" in lc or "reset" in lc:
            reset()
            step_welcome()
        else:
            bot("Type **restart** to size another shaft.")


# ─────────────────────────────────────────────
# RENDER RESULTS
# ─────────────────────────────────────────────
def render_results():
    """Display the sizing results in a professional layout."""
    result = st.session_state.result
    if not result or not result.get("best"):
        return

    best = result["best"]
    alts = result["alternatives"]
    ss = st.session_state

    st.markdown("---")
    st.markdown(
        '<h2 style="color:#c72c41; margin-bottom:0;">📐 HRS Exhaust Shaft Sizing Results</h2>',
        unsafe_allow_html=True,
    )
    st.caption(f"LF Systems HRS — {ss.exhaust_type}")

    # ── System Summary ──
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("#### 🏗️ System Summary")
        summary = {
            "Exhaust Type":         ss.exhaust_type,
            "Number of Floors":     ss.floors,
            "Total Penetrations":   best["total_pens"],
            "Total CFM (100%)":     f'{best["total_cfm"]:,.0f}',
            "Diversity Factor":     f'{ss.diversity_pct}%',
            "Design CFM":           f'{best["design_cfm"]:,.0f}',
            "Floor-to-Floor Height":f'{ss.floor_height} ft',
            "Total Shaft Height":   f'{best["total_height"]} ft',
            "Duct After Last Unit": f'{ss.duct_after_last} ft',
        }
        st.table(pd.DataFrame(summary.items(), columns=["Parameter", "Value"]))

    with col2:
        st.markdown("#### 📏 Recommended Shaft Size")
        shaft_info = {
            "Shaft Size":           best["label"],
            "Gross Area":           f'{best["shaft_area"]} sq.in.',
            "Net Effective Area":   f'{best["eff_area"]} sq.in.',
            "Hydraulic Diameter":   f'{best["dh_in"]}" ',
            "Shaft Velocity":       f'{best["velocity"]:,.0f} FPM',
            "Velocity Pressure":    f'{best["vp"]:.4f} in. WC',
        }
        st.table(pd.DataFrame(shaft_info.items(), columns=["Parameter", "Value"]))

        if best["passes"]:
            st.success(f'✅ ΔP = {best["delta_p"]:.4f} in. WC  —  **PASSES**  (≤ {ss.max_delta_p})')
        else:
            st.error(f'❌ ΔP = {best["delta_p"]:.4f} in. WC  —  **FAILS**  (> {ss.max_delta_p})')

    # ── Pressure Drop Breakdown ──
    st.markdown("#### 📊 Pressure Drop Breakdown")
    dp_data = {
        "Component": ["Shaft Friction", "After-Unit Duct", "Offset Losses",
                       "Entry Loss", "Exit/Fan Loss", "**TOTAL**"],
        "ΔP (in. WC)": [
            f'{best["dp_shaft"]:.4f}',
            f'{best["dp_after"]:.4f}',
            f'{best["dp_offset"]:.4f}',
            f'{best["dp_entry"]:.4f}',
            f'{best["dp_exit"]:.4f}',
            f'**{best["dp_total"]:.4f}**',
        ],
    }
    st.table(pd.DataFrame(dp_data))

    # ── Floor Balance ──
    st.markdown("#### 🏢 Floor Balance Analysis")
    bal_data = {
        "Parameter": [
            "Bottom Floor ΔP",
            "Top Floor ΔP",
            "ΔP Difference (bottom − top)",
            f"Requirement (≤ {ss.max_delta_p} in. WC)",
        ],
        "Value": [
            f'{best["dp_bottom"]:.4f} in. WC',
            f'{best["dp_top"]:.4f} in. WC',
            f'{best["delta_p"]:.4f} in. WC',
            "✅ PASS" if best["passes"] else "❌ FAIL — consider larger shaft",
        ],
    }
    st.table(pd.DataFrame(bal_data))

    # ── Per-Floor Detail (expandable) ──
    if best.get("floor_dp"):
        with st.expander("📋 Per-Floor Pressure Drop Detail", expanded=False):
            rows = []
            for i, dp in enumerate(best["floor_dp"]):
                rows.append({
                    "Floor": i + 1,
                    "ΔP to Fan (in. WC)": f"{dp:.4f}",
                    "Position": "Bottom" if i == 0 else ("Top" if i == len(best["floor_dp"])-1 else ""),
                })
            st.table(pd.DataFrame(rows))

    # ── Alternatives ──
    if alts and len(alts) > 1:
        st.markdown("#### 🔄 Alternative Sizes (meet ΔP requirement)")
        alt_rows = []
        for a in alts:
            alt_rows.append({
                "Size":          a["label"],
                "Eff. Area (sq.in.)": a["eff_area"],
                "Velocity (FPM)":     int(a["velocity"]),
                "ΔP Diff (in. WC)":   f'{a["delta_p"]:.4f}',
                "Status":        "✅" if a["passes"] else "❌",
            })
        st.table(pd.DataFrame(alt_rows))

    # ── Equipment Recommendation ──
    st.markdown("#### 🔧 Recommended Equipment — LF Systems")
    if ss.exhaust_type == "Clothes Dryers":
        equip = {
            "Exhaust Fan":     "DEF (Dryer Exhaust Fan)",
            "Booster":         "DBF (Dryer Booster Fan) — if needed",
            "Controller":      "L150 or L152",
            "System":          "HRS (High Rise Shaft)",
            "Subducts":        f'{ss.floor_data[0]["subduct_size"]}" penetrations',
        }
    elif ss.exhaust_type == "Bathroom Exhaust":
        equip = {
            "Exhaust Fan":     "DEF or inline exhaust fan",
            "Controller":      "L150 or L152",
            "System":          "HRS (High Rise Shaft)",
            "Subducts":        f'{ss.floor_data[0]["subduct_size"]}" penetrations',
        }
    else:
        equip = {
            "Exhaust Fan":     "DEF or rated kitchen hood exhaust fan",
            "Controller":      "L150 or L152",
            "System":          "HRS (High Rise Shaft)",
            "Subducts":        f'{ss.floor_data[0]["subduct_size"]}" penetrations',
        }
    st.table(pd.DataFrame(equip.items(), columns=["Component", "Model / Specification"]))

    st.info(
        "🌐 Visit **[lfsystems.net](https://www.lfsystems.net)** for product specifications, "
        "CAD drawings, and ordering information."
    )

    # ── Engineering Notes ──
    with st.expander("📝 Engineering Notes & Methodology"):
        st.markdown("""
**Calculation Methodology:**
- **Friction Factor:** Colebrook equation (iterative) with ε = 0.0003 ft (galvanized steel)
- **Pressure Drop:** Darcy-Weisbach: `Δp = [f·(L/Dh) + ΣK] · ρ · (V/1096.2)²`
- **Rectangular Equivalence:** Huebscher equation: `De = 1.30·(a·b)^0.625 / (a+b)^0.25`
- **Hydraulic Diameter:** `Dh = 4·A / P`
- **Air Density:** 0.075 lb/ft³ (standard conditions)
- **Fitting Losses:** ASHRAE 2009 Duct Design Chapter 21 coefficients
  - 90° elbow: K = 0.60
  - 45° elbow: K = 0.30
  - Tee branch entry: K = 0.50
  - Exit/fan entry: K = 1.00
  - Bellmouth entry: K = 0.03

**Subduct Area Deductions:**
| Subduct Size | Area Removed |
|---|---|
| 4" | 15.0 sq.in. |
| 6" | 31.5 sq.in. |
| 8" | 54.0 sq.in. |

**Notes:**
- Shaft must be straight between floors (no offsets between occupied floors)
- Offsets are only permitted above the highest floor penetration
- The HRS system maintains constant negative pressure via EC-Flow Technology™
- Diversity factor accounts for simultaneous use (typically 20-100%)
- Maximum allowable ΔP between bottom and top floors: 0.25 in. WC
        """)


# ─────────────────────────────────────────────
# MAIN APP
# ─────────────────────────────────────────────
def main():
    st.set_page_config(
        page_title="HRS Shaft Sizer — LF Systems",
        page_icon="🏗️",
        layout="wide",
    )

    # ── Custom CSS ──
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600;700&display=swap');

    .stApp {
        font-family: 'IBM Plex Sans', sans-serif;
    }
    /* Header banner */
    .hrs-header {
        background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%);
        padding: 18px 28px;
        border-radius: 10px;
        margin-bottom: 20px;
        border-bottom: 4px solid #c72c41;
    }
    .hrs-header h1 {
        color: white;
        margin: 0;
        font-size: 24px;
        font-weight: 700;
    }
    .hrs-header p {
        color: rgba(255,255,255,0.7);
        margin: 4px 0 0 0;
        font-size: 13px;
    }
    .hrs-badge {
        display: inline-block;
        background: #c72c41;
        color: white;
        padding: 4px 12px;
        border-radius: 6px;
        font-weight: 800;
        font-size: 14px;
        margin-right: 10px;
        letter-spacing: -0.5px;
    }
    /* Chat messages */
    .chat-bot {
        background: #f7f4f0;
        border-left: 3px solid #c72c41;
        padding: 12px 16px;
        border-radius: 4px 10px 10px 4px;
        margin: 6px 0;
        font-size: 14px;
        line-height: 1.55;
    }
    .chat-user {
        background: linear-gradient(135deg, #c72c41, #a3213a);
        color: white;
        padding: 10px 16px;
        border-radius: 10px 10px 4px 10px;
        margin: 6px 0 6px auto;
        max-width: 70%;
        text-align: right;
        font-size: 14px;
    }
    /* Tables */
    table {
        font-size: 13px !important;
    }
    </style>
    """, unsafe_allow_html=True)

    # ── Header ──
    st.markdown("""
    <div class="hrs-header">
        <div>
            <span class="hrs-badge">HRS</span>
            <span style="color:white; font-size:20px; font-weight:700;">
                Exhaust Shaft Sizing Calculator
            </span>
        </div>
        <p>LF Systems — High Rise Shaft Constant Pressure System &nbsp;|&nbsp;
        DEF · DBF · L150/L152 &nbsp;|&nbsp;
        <a href="https://www.lfsystems.net" style="color:#ff8a9e;" target="_blank">lfsystems.net</a></p>
    </div>
    """, unsafe_allow_html=True)

    init_state()

    # ── Start the conversation ──
    if st.session_state.step == 0:
        step_welcome()

    # ── Render chat history ──
    for msg in st.session_state.messages:
        if msg["role"] == "assistant":
            with st.chat_message("assistant", avatar="🏗️"):
                st.markdown(msg["text"])
        else:
            with st.chat_message("user", avatar="👤"):
                st.markdown(msg["text"])

    # ── Render results if calculation is done ──
    if st.session_state.calc_done and st.session_state.result:
        render_results()

    # ── Quick-select buttons ──
    step = st.session_state.step
    buttons = []
    if step == 1:
        buttons = ["Dryers", "Bathrooms", "Kitchen Hoods"]
    elif step == 3:
        buttons = ["Yes", "No"]
    elif step == 4 and st.session_state.awaiting == "pens":
        buttons = ["1", "2"]
    elif step == 4 and st.session_state.awaiting == "sub":
        buttons = ["4", "6", "8"]
    elif step == 8:
        buttons = ["Yes", "No"]
    elif step == 10:
        buttons = ["round_auto", "rect_auto", "round_user", "rect_user"]
    elif step == 13:
        buttons = ["restart"]

    if buttons:
        cols = st.columns(len(buttons) + 2)
        for i, b in enumerate(buttons):
            if cols[i + 1].button(b, key=f"qb_{step}_{b}", use_container_width=True):
                process_input(b)
                st.rerun()

    # ── Chat input ──
    if prompt := st.chat_input("Type your answer here..."):
        process_input(prompt)
        st.rerun()

    # ── Sidebar info ──
    with st.sidebar:
        st.markdown("### 🏗️ HRS System Info")
        st.markdown(
            "The **HRS (High Rise Shaft)** system uses a constant pressure "
            "controller to maintain a slight negative pressure in fire-rated "
            "exhaust shafts in high-rise buildings.\n\n"
            "**Applications:**\n"
            "- Clothes dryer exhaust\n"
            "- Bathroom exhaust\n"
            "- Kitchen hood exhaust\n\n"
            "**Key Components:**\n"
            "- DEF — Dryer Exhaust Fan\n"
            "- DBF — Dryer Booster Fan\n"
            "- L150/L152 — Controllers\n\n"
            "**Rules:**\n"
            "- Shaft must be straight between floors\n"
            "- Offsets only after last floor\n"
            "- 1-2 penetrations per floor\n"
            "- Subducts: 4\", 6\", or 8\"\n"
            "- Max ΔP: 0.25 in. WC\n"
            "- Diversity: 20-100%\n"
        )
        st.markdown("---")
        st.markdown(
            "🌐 **[lfsystems.net](https://www.lfsystems.net)**  \n"
            "📞 Contact your LF Systems rep for product selection."
        )
        st.markdown("---")
        if st.button("🔄 Start Over", use_container_width=True):
            reset()
            st.rerun()

        st.markdown("---")
        st.caption("v1.0 — Engineering calculations per ASHRAE 2009 Chapter 21")


if __name__ == "__main__":
    main()
