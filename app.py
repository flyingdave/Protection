import hashlib
import io
import json
import math
from pathlib import Path
from typing import Optional, Tuple

import altair as alt
import pandas as pd
import streamlit as st


CURVE_CONSTANTS = {
    "IEC Standard Inverse": (0.14, 0.02),
    "IEC Very Inverse": (13.5, 1.0),
    "IEC Extremely Inverse": (80.0, 2.0),
    "IEC Long-Time Inverse": (120.0, 1.0),
}

OC_EF_RELAY_MODELS = ["P122", "CDG", "PBO"]
COMMISSIONING_CURVE_OPTIONS = list(CURVE_CONSTANTS.keys()) + ["Definite Time"]
COMMISSIONING_TEST_MULTIPLIERS = [1.5, 3.0, 6.0]
OC_EF_MODEL_NOTES = {
    "P122": "Digital OC/EF worksheet using the selected IEC inverse or definite-time stage.",
    "CDG": "Electromechanical OC/EF worksheet using the selected IEC-style commissioning curve approximation.",
    "PBO": "Static OC/EF worksheet using the selected IEC inverse or definite-time stage.",
}

SEL787_DEFAULT_NOTES = (
    "SEL-787 differential screening uses a standard percentage-restraint commissioning characteristic at relay secondary current level. "
    "Verify final test quantities against the installed compensation, vector group, and winding configuration before site injection."
)

TRANSLAY_RELAY_OPTIONS = ["H04", "H05", "H06", "H07", "Numerical (Newer)"]
TRANSLAY_MODEL_NOTES = {
    "H04": "H04 electromechanical translay worksheet using restrained-spill characteristic and fixed operate-time target.",
    "H05": "H05 translay worksheet using restrained-spill characteristic and fixed operate-time target.",
    "H06": "H06 translay worksheet using restrained-spill characteristic and fixed operate-time target.",
    "H07": "H07 translay worksheet using restrained-spill characteristic and fixed operate-time target.",
    "Numerical (Newer)": "Newer numerical translay worksheet with optional high-set spill element.",
}
TRANSLAY_MODEL_DEFAULTS = {
    "H04": {"pickup_a": 0.20, "slope_pct": 20.0, "operate_time_s": 0.10, "high_set_spill_a": 0.0},
    "H05": {"pickup_a": 0.18, "slope_pct": 22.0, "operate_time_s": 0.09, "high_set_spill_a": 0.0},
    "H06": {"pickup_a": 0.16, "slope_pct": 25.0, "operate_time_s": 0.08, "high_set_spill_a": 0.0},
    "H07": {"pickup_a": 0.14, "slope_pct": 30.0, "operate_time_s": 0.07, "high_set_spill_a": 0.0},
    "Numerical (Newer)": {"pickup_a": 0.10, "slope_pct": 35.0, "operate_time_s": 0.05, "high_set_spill_a": 0.60},
}
COMMISSIONING_WIDGET_DEFAULTS = {
    "oc_ef_device_name": "Feeder 1",
    "oc_ef_ct_primary_a": 200.0,
    "oc_ef_ct_secondary_a": 1.0,
    "oc_ef_pickup_secondary_a": 1.0,
    "oc_ef_tms": 0.10,
    "oc_ef_definite_time_s": 0.30,
    "oc_ef_inst_pickup_a": 0.00,
    "oc_ef_notes": "",
    "sel787_asset_name": "Transformer 1",
    "sel787_w1_ct_primary_a": 800.0,
    "sel787_w1_ct_secondary_a": 1.0,
    "sel787_w2_ct_primary_a": 800.0,
    "sel787_w2_ct_secondary_a": 1.0,
    "sel787_pickup_a": 0.30,
    "sel787_slope1_pct": 25.0,
    "sel787_slope2_pct": 50.0,
    "sel787_breakpoint_a": 4.0,
    "sel787_unrestrained_a": 8.0,
    "sel787_h2_block_pct": 15.0,
    "sel787_h5_block_pct": 35.0,
    "translay_circuit_name": "Feeder Differential 1",
    "translay_local_ct_primary_a": 200.0,
    "translay_local_ct_secondary_a": 1.0,
    "translay_remote_ct_primary_a": 200.0,
    "translay_remote_ct_secondary_a": 1.0,
    "translay_pickup_a": TRANSLAY_MODEL_DEFAULTS["H04"]["pickup_a"],
    "translay_slope_pct": TRANSLAY_MODEL_DEFAULTS["H04"]["slope_pct"],
    "translay_operate_time_s": TRANSLAY_MODEL_DEFAULTS["H04"]["operate_time_s"],
    "translay_high_set_spill_a": TRANSLAY_MODEL_DEFAULTS["H04"]["high_set_spill_a"],
    "translay_notes": "",
}

RELAY_EXPORT_COLUMNS = ["Order", "Device", "Curve", "Pickup_A", "TMS", "Inst_A"]
APP_STATE_FILE = Path(".streamlit/last_entered_state.json")
APP_REVISION = "Rev 0.2"

NETWORK_INPUT_MODE_OPTIONS = ["Template Inputs", "Line-by-line Builder"]
NETWORK_ELEMENT_TYPE_OPTIONS = ["Source", "Reactor", "Cable", "Transformer", "Busbar", "Custom Z"]
NETWORK_ELEMENT_DATA_MODE_OPTIONS = ["Common Data", "Direct Seq Z"]
NETWORK_ELEMENT_COLUMNS = [
    "Order",
    "Name",
    "Element_Type",
    "Data_Mode",
    "Voltage_kV",
    "Parallel_Count",
    "Source_SC_MVA",
    "XR",
    "Reactance_Ohm",
    "Length_km",
    "Cable_Type",
    "HV_kV",
    "LV_kV",
    "Rating_MVA",
    "Z_pct",
    "R1_ohm",
    "X1_ohm",
    "R0_ohm",
    "X0_ohm",
]

CABLE_SEQUENCE_LIBRARY = {
    "33kV 0.25sqin cu HSL": {
        "z1": complex(0.120, 0.105),
        "z2": complex(0.120, 0.105),
        "z0": complex(0.430, 0.310),
        "note": "Legacy HSL estimate from typical utility-era data.",
    },
    "33kV 300mm2 Al": {
        "z1": complex(0.100, 0.100),
        "z2": complex(0.100, 0.100),
        "z0": complex(0.360, 0.280),
        "note": "Typical 33kV 300mm² aluminium underground cable estimate.",
    },
    "11kV PILC 3c 300mm2 Cu": {
        "z1": complex(0.070, 0.085),
        "z2": complex(0.070, 0.085),
        "z0": complex(0.280, 0.240),
        "note": "Legacy PILC estimate.",
    },
    "11kV XLPE 3c 300mm2 Cu": {
        "z1": complex(0.075, 0.090),
        "z2": complex(0.075, 0.090),
        "z0": complex(0.260, 0.220),
        "note": "Typical XLPE estimate.",
    },
    "11kV 3c 800mm2 Cu": {
        "z1": complex(0.028, 0.075),
        "z2": complex(0.028, 0.075),
        "z0": complex(0.120, 0.200),
        "note": "Large copper cable estimate.",
    },
    "11kV PLY 3c 240mm2 Cu": {
        "z1": complex(0.095, 0.090),
        "z2": complex(0.095, 0.090),
        "z0": complex(0.330, 0.270),
        "note": "Legacy PLY cable estimate.",
    },
    "Default Cable": {
        "z1": complex(0.100, 0.100),
        "z2": complex(0.100, 0.100),
        "z0": complex(0.350, 0.300),
        "note": "Fallback default where site data is unavailable.",
    },
}

HV_CABLE_OPTIONS = ["33kV 0.25sqin cu HSL", "33kV 300mm2 Al", "Default Cable"]
LV_CABLE_OPTIONS = [
    "11kV PILC 3c 300mm2 Cu",
    "11kV XLPE 3c 300mm2 Cu",
    "11kV 3c 800mm2 Cu",
    "11kV PLY 3c 240mm2 Cu",
    "Default Cable",
]

ARC_FLASH_SWITCHGEAR_FACTOR_MAP = {
    ("Horizontal", "Enclosed"): 1.40,
    ("Vertical", "Enclosed"): 1.25,
    ("Horizontal", "Unenclosed"): 1.10,
    ("Vertical", "Unenclosed"): 1.00,
}

SOURCE_Z0_FACTOR = 3.0
TRANSFORMER_Z0_FACTOR = 1.0

BASELINE_RELAY_SETTINGS = [
    {
        "Order": 1,
        "Device": "Feeder Relay",
        "Curve": "IEC Standard Inverse",
        "Pickup_A": 300.0,
        "TMS": 0.10,
        "Inst_A": 2500.0,
    },
    {
        "Order": 2,
        "Device": "Incomer Relay",
        "Curve": "IEC Very Inverse",
        "Pickup_A": 500.0,
        "TMS": 0.20,
        "Inst_A": 5000.0,
    },
    {
        "Order": 3,
        "Device": "Source Relay",
        "Curve": "IEC Very Inverse",
        "Pickup_A": 700.0,
        "TMS": 0.35,
        "Inst_A": 0.0,
    },
]

DEFAULT_NETWORK_PRESET = {
    "project_name": "Protection - 33/11kV Network",
    "network_input_mode": "Template Inputs",
    "grading_margin_s": 0.30,
    "v_ll_kv": 11.0,
    "frequency_hz": 50,
    "source_v_kv": 33.0,
    "source_sc_mva": 1850.0,
    "source_xr": 10.0,
    "reactor_ohm": 3.23,
    "hv_cable_length_km": 3.0,
    "hv_cable_type": "33kV 0.25sqin cu HSL",
    "transformer_hv_kv": 33.0,
    "transformer_lv_kv": 11.0,
    "transformer_mva": 13.5,
    "transformer_z_pct": 7.2,
    "transformer_xr": 8.0,
    "lv_cable_length_km": 0.01,
    "lv_cable_type": "11kV XLPE 3c 300mm2 Cu",
    "duplicate_source_to_first_busbar_circuit": False,
    "parallel_lv_transformer_cable": False,
    "feeder_length_km": 0.5,
    "feeder_cable_type": "Default Cable",
    "parallel_feeder_cable": False,
    "study_bus": "11kV Transformer Busbar",
    "fault_type": "3-Phase",
    "bolted_fault_current_ka": 10.0,
    "arc_current_factor": 0.85,
    "clearing_time_s": 0.20,
    "working_distance_mm": 910,
    "arc_orientation": "Vertical",
    "arc_enclosure": "Enclosed",
}

DEFAULT_NETWORK_ELEMENTS = [
    {
        "Order": 1,
        "Name": "33kV Source",
        "Element_Type": "Source",
        "Data_Mode": "Common Data",
        "Voltage_kV": 33.0,
        "Parallel_Count": 1,
        "Source_SC_MVA": 1850.0,
        "XR": 10.0,
        "Reactance_Ohm": 0.0,
        "Length_km": 0.0,
        "Cable_Type": "Default Cable",
        "HV_kV": 33.0,
        "LV_kV": 33.0,
        "Rating_MVA": 0.0,
        "Z_pct": 0.0,
        "R1_ohm": 0.0,
        "X1_ohm": 0.0,
        "R0_ohm": 0.0,
        "X0_ohm": 0.0,
    },
    {
        "Order": 2,
        "Name": "33kV Reactor",
        "Element_Type": "Reactor",
        "Data_Mode": "Common Data",
        "Voltage_kV": 33.0,
        "Parallel_Count": 1,
        "Source_SC_MVA": 0.0,
        "XR": 0.0,
        "Reactance_Ohm": 3.23,
        "Length_km": 0.0,
        "Cable_Type": "Default Cable",
        "HV_kV": 33.0,
        "LV_kV": 33.0,
        "Rating_MVA": 0.0,
        "Z_pct": 0.0,
        "R1_ohm": 0.0,
        "X1_ohm": 0.0,
        "R0_ohm": 0.0,
        "X0_ohm": 0.0,
    },
    {
        "Order": 3,
        "Name": "33kV Cable",
        "Element_Type": "Cable",
        "Data_Mode": "Common Data",
        "Voltage_kV": 33.0,
        "Parallel_Count": 1,
        "Source_SC_MVA": 0.0,
        "XR": 0.0,
        "Reactance_Ohm": 0.0,
        "Length_km": 3.0,
        "Cable_Type": "33kV 0.25sqin cu HSL",
        "HV_kV": 33.0,
        "LV_kV": 33.0,
        "Rating_MVA": 0.0,
        "Z_pct": 0.0,
        "R1_ohm": 0.0,
        "X1_ohm": 0.0,
        "R0_ohm": 0.0,
        "X0_ohm": 0.0,
    },
    {
        "Order": 4,
        "Name": "33/11kV Transformer",
        "Element_Type": "Transformer",
        "Data_Mode": "Common Data",
        "Voltage_kV": 11.0,
        "Parallel_Count": 1,
        "Source_SC_MVA": 0.0,
        "XR": 8.0,
        "Reactance_Ohm": 0.0,
        "Length_km": 0.0,
        "Cable_Type": "Default Cable",
        "HV_kV": 33.0,
        "LV_kV": 11.0,
        "Rating_MVA": 13.5,
        "Z_pct": 7.2,
        "R1_ohm": 0.0,
        "X1_ohm": 0.0,
        "R0_ohm": 0.0,
        "X0_ohm": 0.0,
    },
    {
        "Order": 5,
        "Name": "11kV Transformer Cable",
        "Element_Type": "Cable",
        "Data_Mode": "Common Data",
        "Voltage_kV": 11.0,
        "Parallel_Count": 1,
        "Source_SC_MVA": 0.0,
        "XR": 0.0,
        "Reactance_Ohm": 0.0,
        "Length_km": 0.01,
        "Cable_Type": "11kV XLPE 3c 300mm2 Cu",
        "HV_kV": 11.0,
        "LV_kV": 11.0,
        "Rating_MVA": 0.0,
        "Z_pct": 0.0,
        "R1_ohm": 0.0,
        "X1_ohm": 0.0,
        "R0_ohm": 0.0,
        "X0_ohm": 0.0,
    },
    {
        "Order": 6,
        "Name": "11kV Transformer Busbar",
        "Element_Type": "Busbar",
        "Data_Mode": "Common Data",
        "Voltage_kV": 11.0,
        "Parallel_Count": 1,
        "Source_SC_MVA": 0.0,
        "XR": 0.0,
        "Reactance_Ohm": 0.0,
        "Length_km": 0.0,
        "Cable_Type": "Default Cable",
        "HV_kV": 11.0,
        "LV_kV": 11.0,
        "Rating_MVA": 0.0,
        "Z_pct": 0.0,
        "R1_ohm": 0.0,
        "X1_ohm": 0.0,
        "R0_ohm": 0.0,
        "X0_ohm": 0.0,
    },
    {
        "Order": 7,
        "Name": "11kV Feeder Cable",
        "Element_Type": "Cable",
        "Data_Mode": "Common Data",
        "Voltage_kV": 11.0,
        "Parallel_Count": 1,
        "Source_SC_MVA": 0.0,
        "XR": 0.0,
        "Reactance_Ohm": 0.0,
        "Length_km": 0.5,
        "Cable_Type": "Default Cable",
        "HV_kV": 11.0,
        "LV_kV": 11.0,
        "Rating_MVA": 0.0,
        "Z_pct": 0.0,
        "R1_ohm": 0.0,
        "X1_ohm": 0.0,
        "R0_ohm": 0.0,
        "X0_ohm": 0.0,
    },
    {
        "Order": 8,
        "Name": "11kV Remote Busbar",
        "Element_Type": "Busbar",
        "Data_Mode": "Common Data",
        "Voltage_kV": 11.0,
        "Parallel_Count": 1,
        "Source_SC_MVA": 0.0,
        "XR": 0.0,
        "Reactance_Ohm": 0.0,
        "Length_km": 0.0,
        "Cable_Type": "Default Cable",
        "HV_kV": 11.0,
        "LV_kV": 11.0,
        "Rating_MVA": 0.0,
        "Z_pct": 0.0,
        "R1_ohm": 0.0,
        "X1_ohm": 0.0,
        "R0_ohm": 0.0,
        "X0_ohm": 0.0,
    },
]

NETWORK_ARRANGEMENT = [
    "33kV source (fault level 1850 MVA)",
    "33kV CB",
    "33kV reactor (3.23 Ω)",
    "3 km underground cable",
    "33/11kV transformer (13.5 MVA, 7.2%)",
    "11kV single-core cable (10 m)",
    "Transformer CB",
    "11kV busbar",
    "11kV feeder CB",
    "11kV feeder",
    "11kV CB",
    "11kV busbar",
    "11kV CB",
    "11kV feeder",
    "11kV CB",
]

DEFAULT_STUDY_CASE = DEFAULT_NETWORK_PRESET.copy()

STUDY_CASE_OPTION_VALUES = {
    "network_input_mode": set(NETWORK_INPUT_MODE_OPTIONS),
    "frequency_hz": {50, 60},
    "fault_type": {"3-Phase", "Line-Line", "Line-Ground"},
    "arc_orientation": {"Horizontal", "Vertical"},
    "arc_enclosure": {"Enclosed", "Unenclosed"},
    "hv_cable_type": set(HV_CABLE_OPTIONS),
    "lv_cable_type": set(LV_CABLE_OPTIONS),
    "feeder_cable_type": set(LV_CABLE_OPTIONS),
}

STUDY_CASE_BOUNDS = {
    "grading_margin_s": (0.10, 1.00),
    "v_ll_kv": (0.4, 66.0),
    "source_v_kv": (1.0, 400.0),
    "source_sc_mva": (10.0, 50000.0),
    "source_xr": (0.0, 50.0),
    "reactor_ohm": (0.0, 50.0),
    "hv_cable_length_km": (0.0, 100.0),
    "transformer_hv_kv": (1.0, 400.0),
    "transformer_lv_kv": (0.4, 66.0),
    "transformer_mva": (0.1, 1000.0),
    "transformer_z_pct": (1.0, 25.0),
    "transformer_xr": (0.0, 50.0),
    "lv_cable_length_km": (0.0, 20.0),
    "feeder_length_km": (0.0, 100.0),
    "bolted_fault_current_ka": (0.1, 200.0),
    "arc_current_factor": (0.50, 1.00),
    "clearing_time_s": (0.01, 5.00),
    "working_distance_mm": (200, 2000),
}

STUDY_CASE_BOOL_FIELDS = {
    "duplicate_source_to_first_busbar_circuit",
    "parallel_lv_transformer_cable",
    "parallel_feeder_cable",
}

STUDY_CASE_INT_FIELDS = {"frequency_hz", "working_distance_mm"}

RELAY_COLUMN_ALIASES = {
    "order": "Order",
    "device": "Device",
    "curve": "Curve",
    "pickup": "Pickup_A",
    "pickupa": "Pickup_A",
    "pickupamps": "Pickup_A",
    "tms": "TMS",
    "inst": "Inst_A",
    "insta": "Inst_A",
    "instantaneous": "Inst_A",
    "instantaneousa": "Inst_A",
    "instantaneouspickup": "Inst_A",
    "instantaneouspickupa": "Inst_A",
}


def split_rx_from_z_magnitude(z_ohm: float, x_over_r: float) -> complex:
    if z_ohm <= 0:
        return complex(0.0, 0.0)
    if x_over_r <= 0:
        return complex(z_ohm, 0.0)

    resistance = z_ohm / math.sqrt(1 + x_over_r**2)
    reactance = resistance * x_over_r
    return complex(resistance, reactance)


def cable_sequence_impedance_per_km(cable_type: str) -> dict:
    return CABLE_SEQUENCE_LIBRARY.get(cable_type, CABLE_SEQUENCE_LIBRARY["Default Cable"])


def format_complex_ohm(z_value: complex) -> str:
    return f"{z_value.real:.4f} + j{z_value.imag:.4f}"


def calc_fault_values(v_ll_kv: float, z1_total: complex, z2_total: complex, z0_total: complex) -> dict:
    z1_magnitude = abs(z1_total)
    z_ll = z1_total + z2_total
    z_lg = z1_total + z2_total + z0_total

    if z1_magnitude <= 0 or v_ll_kv <= 0:
        return {
            "Z1_total_ohm": 0.0,
            "Z2_total_ohm": 0.0,
            "Z0_total_ohm": 0.0,
            "I_3ph_kA": math.inf,
            "I_LL_kA": math.inf,
            "I_LG_kA": math.inf,
            "Fault_MVA": math.inf,
        }

    i_3ph_ka = v_ll_kv / (math.sqrt(3) * z1_magnitude)
    i_ll_ka = (v_ll_kv / abs(z_ll)) if abs(z_ll) > 0 else math.inf
    i_lg_ka = (math.sqrt(3) * v_ll_kv / abs(z_lg)) if abs(z_lg) > 0 else math.inf
    fault_mva = math.sqrt(3) * v_ll_kv * i_3ph_ka

    return {
        "Z1_total_ohm": abs(z1_total),
        "Z2_total_ohm": abs(z2_total),
        "Z0_total_ohm": abs(z0_total),
        "I_3ph_kA": i_3ph_ka,
        "I_LL_kA": i_ll_ka,
        "I_LG_kA": i_lg_ka,
        "Fault_MVA": fault_mva,
    }


def calc_relay_time_s(
    current_a: float,
    pickup_a: float,
    tms: float,
    curve_name: str,
    inst_pickup_a: float,
) -> float:
    if pickup_a <= 0 or tms <= 0 or current_a <= pickup_a:
        return math.inf

    if inst_pickup_a > 0 and current_a >= inst_pickup_a:
        return 0.05

    k, alpha = CURVE_CONSTANTS.get(curve_name, CURVE_CONSTANTS["IEC Standard Inverse"])
    ratio = current_a / pickup_a
    denominator = ratio**alpha - 1

    if denominator <= 0:
        return math.inf

    return tms * k / denominator


def calc_commissioning_time_s(
    current_a: float,
    pickup_a: float,
    curve_name: str,
    tms: float,
    definite_time_s: float,
    inst_pickup_a: float,
) -> float:
    if pickup_a <= 0 or current_a <= pickup_a:
        return math.inf

    if inst_pickup_a > 0 and current_a >= inst_pickup_a:
        return 0.05

    if curve_name == "Definite Time":
        return definite_time_s if definite_time_s > 0 else math.inf

    return calc_relay_time_s(
        current_a=current_a,
        pickup_a=pickup_a,
        tms=tms,
        curve_name=curve_name,
        inst_pickup_a=0.0,
    )


def build_oc_ef_test_points(
    pickup_secondary_a: float,
    curve_name: str,
    tms: float,
    definite_time_s: float,
    inst_pickup_a: float,
    ct_primary_a: float,
    ct_secondary_a: float,
    test_multipliers: Optional[list[float]] = None,
) -> pd.DataFrame:
    multipliers = test_multipliers or COMMISSIONING_TEST_MULTIPLIERS
    ct_ratio = (ct_primary_a / ct_secondary_a) if ct_secondary_a > 0 else math.nan

    rows = []
    for index, multiplier in enumerate(multipliers, start=1):
        injection_current_secondary_a = pickup_secondary_a * multiplier
        injection_current_primary_a = (
            injection_current_secondary_a * ct_ratio if math.isfinite(ct_ratio) else math.nan
        )
        expected_time_s = calc_commissioning_time_s(
            current_a=injection_current_secondary_a,
            pickup_a=pickup_secondary_a,
            curve_name=curve_name,
            tms=tms,
            definite_time_s=definite_time_s,
            inst_pickup_a=inst_pickup_a,
        )
        operating_mode = "Definite Time" if curve_name == "Definite Time" else "Inverse"
        if inst_pickup_a > 0 and injection_current_secondary_a >= inst_pickup_a:
            operating_mode = "High-set / instantaneous"

        rows.append(
            {
                "Point": f"P{index}",
                "Purpose": f"{multiplier:.2f} x pickup",
                "Injection_Current_Secondary_A": injection_current_secondary_a,
                "Injection_Current_Primary_A": injection_current_primary_a,
                "Expected_Operate_s": expected_time_s,
                "Expected_Mode": operating_mode,
            }
        )

    return pd.DataFrame(rows)


def sel787_operate_threshold(
    restraint_current_a: float,
    pickup_a: float,
    slope1_pct: float,
    slope2_pct: float,
    breakpoint_a: float,
) -> float:
    restraint_current_a = max(restraint_current_a, 0.0)
    pickup_a = max(pickup_a, 0.0)
    slope1 = max(slope1_pct, 0.0) / 100.0
    slope2 = max(slope2_pct, 0.0) / 100.0
    breakpoint_a = max(breakpoint_a, 0.0)

    if restraint_current_a <= breakpoint_a:
        return pickup_a + slope1 * restraint_current_a

    return pickup_a + slope1 * breakpoint_a + slope2 * (restraint_current_a - breakpoint_a)


def sel787_current_pair_from_operate_restraint(
    operate_current_a: float,
    restraint_current_a: float,
) -> tuple[float, float]:
    side_1_current_a = max(restraint_current_a + (operate_current_a / 2.0), 0.0)
    side_2_current_a = max(restraint_current_a - (operate_current_a / 2.0), 0.0)
    return side_1_current_a, side_2_current_a


def build_sel787_test_points(
    pickup_a: float,
    slope1_pct: float,
    slope2_pct: float,
    breakpoint_a: float,
    unrestrained_a: float,
) -> pd.DataFrame:
    low_bias_restraint_a = (breakpoint_a * 0.5) if breakpoint_a > 0 else max(1.0, pickup_a)
    high_bias_restraint_a = (breakpoint_a * 1.5) if breakpoint_a > 0 else max(4.0, pickup_a * 4.0)
    stability_restraint_a = max(high_bias_restraint_a, pickup_a * 4.0, 2.0)

    rows = []

    stability_side_1_a, stability_side_2_a = sel787_current_pair_from_operate_restraint(
        operate_current_a=0.0,
        restraint_current_a=stability_restraint_a,
    )
    rows.append(
        {
            "Test": "Balanced stability",
            "Purpose": "Through-current stability check",
            "Irest_A": stability_restraint_a,
            "Iop_A": 0.0,
            "Side1_Secondary_A": stability_side_1_a,
            "Side2_Secondary_A": stability_side_2_a,
            "Expected_Result": "No trip",
            "Characteristic_Threshold_A": sel787_operate_threshold(
                stability_restraint_a,
                pickup_a,
                slope1_pct,
                slope2_pct,
                breakpoint_a,
            ),
            "Notes": "Inject equal compensated currents on both sides.",
        }
    )

    for test_name, purpose, restraint_current_a in [
        ("Low-bias operate", "Pickup / slope-1 operate check", low_bias_restraint_a),
        ("High-bias operate", "Slope-2 operate check", high_bias_restraint_a),
    ]:
        threshold_a = sel787_operate_threshold(
            restraint_current_a,
            pickup_a,
            slope1_pct,
            slope2_pct,
            breakpoint_a,
        )
        operate_current_a = max(threshold_a * 1.10, pickup_a * 1.05)
        note_text = "Inject currents in opposing sense to create operate current."
        if unrestrained_a > 0 and operate_current_a >= unrestrained_a:
            note_text += " Requested point may also operate the unrestrained element."

        side_1_current_a, side_2_current_a = sel787_current_pair_from_operate_restraint(
            operate_current_a=operate_current_a,
            restraint_current_a=restraint_current_a,
        )
        rows.append(
            {
                "Test": test_name,
                "Purpose": purpose,
                "Irest_A": restraint_current_a,
                "Iop_A": operate_current_a,
                "Side1_Secondary_A": side_1_current_a,
                "Side2_Secondary_A": side_2_current_a,
                "Expected_Result": "Trip",
                "Characteristic_Threshold_A": threshold_a,
                "Notes": note_text,
            }
        )

    return pd.DataFrame(rows)


def translay_trip_threshold(
    through_current_a: float,
    pickup_a: float,
    slope_pct: float,
) -> float:
    through_current_a = max(through_current_a, 0.0)
    pickup_a = max(pickup_a, 0.0)
    slope = max(slope_pct, 0.0) / 100.0
    return pickup_a + slope * through_current_a


def translay_local_remote_from_through_spill(
    through_current_a: float,
    spill_current_a: float,
) -> tuple[float, float]:
    local_current_a = max(through_current_a + (spill_current_a / 2.0), 0.0)
    remote_current_a = max(through_current_a - (spill_current_a / 2.0), 0.0)
    return local_current_a, remote_current_a


def build_translay_test_points(
    pickup_a: float,
    slope_pct: float,
    operate_time_s: float,
    high_set_spill_a: float,
) -> pd.DataFrame:
    pickup_a = max(pickup_a, 0.01)
    slope_pct = max(slope_pct, 0.0)
    operate_time_s = max(operate_time_s, 0.01)
    high_set_spill_a = max(high_set_spill_a, 0.0)

    low_through_a = max(1.0, pickup_a * 4.0)
    high_through_a = max(2.0, low_through_a * 2.0)
    stability_through_a = max(2.0, high_through_a)

    test_plan_rows = []

    stability_threshold_a = translay_trip_threshold(stability_through_a, pickup_a, slope_pct)
    stability_local_a, stability_remote_a = translay_local_remote_from_through_spill(
        through_current_a=stability_through_a,
        spill_current_a=0.0,
    )
    test_plan_rows.append(
        {
            "Point": "T1",
            "Purpose": "Through-current stability",
            "Through_Current_A": stability_through_a,
            "Spill_Current_A": 0.0,
            "Local_Current_A": stability_local_a,
            "Remote_Current_A": stability_remote_a,
            "Trip_Threshold_A": stability_threshold_a,
            "Expected_Result": "No trip",
            "Expected_Operate_s": math.inf,
            "Expected_Mode": "Stability",
        }
    )

    for point_name, purpose, through_current_a in [
        ("T2", "Low-bias operate", low_through_a),
        ("T3", "High-bias operate", high_through_a),
    ]:
        threshold_a = translay_trip_threshold(through_current_a, pickup_a, slope_pct)
        spill_current_a = max(threshold_a * 1.10, pickup_a * 1.05)
        local_current_a, remote_current_a = translay_local_remote_from_through_spill(
            through_current_a=through_current_a,
            spill_current_a=spill_current_a,
        )

        expected_mode = "Restrained spill"
        expected_operate_s = operate_time_s
        if high_set_spill_a > 0 and spill_current_a >= high_set_spill_a:
            expected_mode = "High-set spill"
            expected_operate_s = 0.05

        test_plan_rows.append(
            {
                "Point": point_name,
                "Purpose": purpose,
                "Through_Current_A": through_current_a,
                "Spill_Current_A": spill_current_a,
                "Local_Current_A": local_current_a,
                "Remote_Current_A": remote_current_a,
                "Trip_Threshold_A": threshold_a,
                "Expected_Result": "Trip",
                "Expected_Operate_s": expected_operate_s,
                "Expected_Mode": expected_mode,
            }
        )

    return pd.DataFrame(test_plan_rows)


def format_time_value(time_s: float) -> str:
    return f"{time_s:.3f}" if math.isfinite(time_s) else "No trip"


def arc_flash_category(incident_energy_cal_cm2: float) -> str:
    if incident_energy_cal_cm2 < 1.2:
        return "Below 1.2 cal/cm² (typically no arc-rated PPE threshold reached)"
    if incident_energy_cal_cm2 < 4:
        return "PPE Category 1 range"
    if incident_energy_cal_cm2 < 8:
        return "PPE Category 2 range"
    if incident_energy_cal_cm2 < 25:
        return "PPE Category 3 range"
    if incident_energy_cal_cm2 < 40:
        return "PPE Category 4 range"
    return "Above 40 cal/cm² (danger zone, additional controls required)"


def sanitize_relay_settings(settings_df: pd.DataFrame) -> pd.DataFrame:
    if settings_df.empty:
        return pd.DataFrame(columns=RELAY_EXPORT_COLUMNS)

    required_columns = set(RELAY_EXPORT_COLUMNS)
    if not required_columns.issubset(set(settings_df.columns)):
        return pd.DataFrame(columns=RELAY_EXPORT_COLUMNS)

    relay_df = settings_df.copy()
    relay_df["Order"] = pd.to_numeric(relay_df["Order"], errors="coerce")
    relay_df["Pickup_A"] = pd.to_numeric(relay_df["Pickup_A"], errors="coerce")
    relay_df["TMS"] = pd.to_numeric(relay_df["TMS"], errors="coerce")
    relay_df["Inst_A"] = pd.to_numeric(relay_df["Inst_A"], errors="coerce").fillna(0.0)
    relay_df["Curve"] = relay_df["Curve"].where(
        relay_df["Curve"].isin(CURVE_CONSTANTS.keys()),
        "IEC Standard Inverse",
    )

    relay_df = relay_df.dropna(subset=["Device", "Order", "Pickup_A", "TMS"])
    relay_df["Order"] = relay_df["Order"].astype(int)
    relay_df["Device"] = relay_df["Device"].astype(str)
    return relay_df[RELAY_EXPORT_COLUMNS]


def default_relay_settings_df() -> pd.DataFrame:
    return pd.DataFrame(BASELINE_RELAY_SETTINGS)[RELAY_EXPORT_COLUMNS]


def default_network_elements_df() -> pd.DataFrame:
    return pd.DataFrame(DEFAULT_NETWORK_ELEMENTS)[NETWORK_ELEMENT_COLUMNS]


def sanitize_network_elements(elements_df: pd.DataFrame) -> pd.DataFrame:
    if elements_df.empty:
        return pd.DataFrame(columns=NETWORK_ELEMENT_COLUMNS)

    network_df = elements_df.copy()
    for column in NETWORK_ELEMENT_COLUMNS:
        if column not in network_df.columns:
            network_df[column] = 0.0 if column not in {"Name", "Element_Type", "Data_Mode", "Cable_Type"} else ""

    numeric_columns = [
        "Order",
        "Voltage_kV",
        "Parallel_Count",
        "Source_SC_MVA",
        "XR",
        "Reactance_Ohm",
        "Length_km",
        "HV_kV",
        "LV_kV",
        "Rating_MVA",
        "Z_pct",
        "R1_ohm",
        "X1_ohm",
        "R0_ohm",
        "X0_ohm",
    ]
    for column in numeric_columns:
        network_df[column] = pd.to_numeric(network_df[column], errors="coerce").fillna(0.0)

    network_df["Order"] = network_df["Order"].replace(0.0, pd.NA)
    network_df["Order"] = network_df["Order"].fillna(pd.Series(range(1, len(network_df) + 1), index=network_df.index))
    network_df["Order"] = network_df["Order"].astype(int)
    network_df["Parallel_Count"] = network_df["Parallel_Count"].clip(lower=1).round().astype(int)

    network_df["Name"] = network_df["Name"].fillna("").astype(str).str.strip()
    network_df["Element_Type"] = network_df["Element_Type"].fillna("Busbar").astype(str).str.strip()
    network_df["Data_Mode"] = network_df["Data_Mode"].fillna("Common Data").astype(str).str.strip()
    network_df["Cable_Type"] = network_df["Cable_Type"].fillna("Default Cable").astype(str).str.strip()

    network_df["Element_Type"] = network_df["Element_Type"].where(
        network_df["Element_Type"].isin(NETWORK_ELEMENT_TYPE_OPTIONS),
        "Busbar",
    )
    network_df["Data_Mode"] = network_df["Data_Mode"].where(
        network_df["Data_Mode"].isin(NETWORK_ELEMENT_DATA_MODE_OPTIONS),
        "Common Data",
    )
    all_cable_options = set(HV_CABLE_OPTIONS) | set(LV_CABLE_OPTIONS)
    network_df["Cable_Type"] = network_df["Cable_Type"].where(
        network_df["Cable_Type"].isin(all_cable_options),
        "Default Cable",
    )

    network_df = network_df.sort_values(["Order", "Name"]).reset_index(drop=True)
    for index, row in network_df.iterrows():
        if not row["Name"]:
            network_df.at[index, "Name"] = f"{row['Element_Type']} {int(row['Order'])}"

    return network_df[NETWORK_ELEMENT_COLUMNS]


def refer_impedance_to_voltage(z_value: complex, from_kv: float, to_kv: float) -> complex:
    if from_kv <= 0 or to_kv <= 0 or math.isclose(from_kv, to_kv, rel_tol=1e-9, abs_tol=1e-9):
        return z_value
    return z_value * ((to_kv / from_kv) ** 2)


def direct_sequence_impedance_from_row(
    row: pd.Series,
    z0_multiplier_fallback: float = 1.0,
    default_z0_to_z1: bool = False,
) -> tuple[complex, complex, complex]:
    z1 = complex(float(row["R1_ohm"]), float(row["X1_ohm"]))
    z2 = z1
    raw_z0 = complex(float(row["R0_ohm"]), float(row["X0_ohm"]))
    if abs(raw_z0) > 0:
        z0 = raw_z0
    elif default_z0_to_z1:
        z0 = z1 * z0_multiplier_fallback
    else:
        z0 = z1
    return z1, z2, z0


def build_single_line_diagram_dot(network_df: pd.DataFrame) -> str:
    lines = [
        "digraph NetworkSingleLine {",
        "rankdir=LR;",
        'graph [fontname="Helvetica", labelloc="t"];',
        'node [fontname="Helvetica", shape=box, style="rounded,filled", fillcolor="#F7F9FC"];',
        'edge [fontname="Helvetica"];',
    ]

    previous_node_id = None
    for index, row in network_df.iterrows():
        node_id = f"n{index}"
        element_type = str(row["Element_Type"])
        voltage_kv = float(row["Voltage_kV"])
        parallel_count = int(row["Parallel_Count"])
        label_lines = [str(row["Name"]), element_type]
        if element_type == "Transformer":
            label_lines.append(f"{float(row['HV_kV']):.1f}/{float(row['LV_kV']):.1f} kV")
        elif voltage_kv > 0:
            label_lines.append(f"{voltage_kv:.1f} kV")
        if parallel_count > 1:
            label_lines.append(f"x{parallel_count} parallel")

        shape = "box"
        fillcolor = "#F7F9FC"
        if element_type == "Source":
            shape = "ellipse"
            fillcolor = "#E8F3FF"
        elif element_type == "Busbar":
            shape = "component"
            fillcolor = "#FFF4D6"
        elif element_type == "Transformer":
            shape = "hexagon"
            fillcolor = "#E8F7E8"

        lines.append(
            f'{node_id} [label="{"\\n".join(label_lines)}", shape={shape}, fillcolor="{fillcolor}"];'
        )
        if previous_node_id is not None:
            lines.append(f"{previous_node_id} -> {node_id};")
        previous_node_id = node_id

    lines.append("}")
    return "\n".join(lines)


def calculate_line_builder_network(network_elements_df: pd.DataFrame) -> dict:
    network_df = sanitize_network_elements(network_elements_df)

    if network_df.empty:
        return {
            "network_df": network_df,
            "cable_df": pd.DataFrame(),
            "impedance_df": pd.DataFrame(),
            "fault_df": pd.DataFrame(),
            "diagram_dot": build_single_line_diagram_dot(default_network_elements_df()),
            "warnings": ["Add at least one source and one busbar in the line builder to calculate fault levels."],
            "max_fault_current_ka": 0.0,
        }

    current_base_kv: float | None = None
    cumulative_z1 = complex(0.0, 0.0)
    cumulative_z2 = complex(0.0, 0.0)
    cumulative_z0 = complex(0.0, 0.0)

    cable_rows: list[dict] = []
    impedance_rows: list[dict] = []
    fault_rows: list[dict] = []
    warnings: list[str] = []

    for _, row in network_df.iterrows():
        name = str(row["Name"])
        element_type = str(row["Element_Type"])
        data_mode = str(row["Data_Mode"])
        voltage_kv = float(row["Voltage_kV"])
        parallel_count = max(int(row["Parallel_Count"]), 1)

        element_z1 = complex(0.0, 0.0)
        element_z2 = complex(0.0, 0.0)
        element_z0 = complex(0.0, 0.0)
        effective_voltage_kv = current_base_kv if current_base_kv is not None else voltage_kv

        if element_type == "Source":
            if data_mode == "Common Data":
                if voltage_kv <= 0 or float(row["Source_SC_MVA"]) <= 0:
                    warnings.append(f"Source '{name}' needs voltage and fault level input.")
                source_z = split_rx_from_z_magnitude((voltage_kv**2) / max(float(row["Source_SC_MVA"]), 1e-9), float(row["XR"]))
                element_z1, element_z2, element_z0 = (
                    source_z,
                    source_z,
                    source_z * SOURCE_Z0_FACTOR,
                )
            else:
                element_z1, element_z2, element_z0 = direct_sequence_impedance_from_row(
                    row,
                    z0_multiplier_fallback=SOURCE_Z0_FACTOR,
                    default_z0_to_z1=True,
                )

            current_base_kv = voltage_kv if voltage_kv > 0 else current_base_kv
            effective_voltage_kv = current_base_kv if current_base_kv is not None else voltage_kv
            cumulative_z1 += element_z1 / parallel_count
            cumulative_z2 += element_z2 / parallel_count
            cumulative_z0 += element_z0 / parallel_count

        elif element_type == "Busbar":
            if current_base_kv is None and voltage_kv > 0:
                current_base_kv = voltage_kv
            elif current_base_kv is not None and voltage_kv > 0 and not math.isclose(current_base_kv, voltage_kv):
                cumulative_z1 = refer_impedance_to_voltage(cumulative_z1, current_base_kv, voltage_kv)
                cumulative_z2 = refer_impedance_to_voltage(cumulative_z2, current_base_kv, voltage_kv)
                cumulative_z0 = refer_impedance_to_voltage(cumulative_z0, current_base_kv, voltage_kv)
                current_base_kv = voltage_kv

            bus_voltage_kv = current_base_kv if current_base_kv is not None else voltage_kv
            effective_voltage_kv = bus_voltage_kv
            if bus_voltage_kv > 0:
                fault_result = calc_fault_values(bus_voltage_kv, cumulative_z1, cumulative_z2, cumulative_z0)
                fault_rows.append(
                    {
                        "Bus": name,
                        "Z1_total_ohm": fault_result["Z1_total_ohm"],
                        "Z2_total_ohm": fault_result["Z2_total_ohm"],
                        "Z0_total_ohm": fault_result["Z0_total_ohm"],
                        "I_3ph_kA": fault_result["I_3ph_kA"],
                        "I_LL_kA": fault_result["I_LL_kA"],
                        "I_LG_kA": fault_result["I_LG_kA"],
                        "Fault_MVA": fault_result["Fault_MVA"],
                    }
                )
            else:
                warnings.append(f"Busbar '{name}' needs a valid voltage input.")

        elif element_type == "Transformer":
            hv_kv = float(row["HV_kV"]) if float(row["HV_kV"]) > 0 else current_base_kv or voltage_kv
            lv_kv = float(row["LV_kV"]) if float(row["LV_kV"]) > 0 else voltage_kv or hv_kv

            if current_base_kv is None:
                current_base_kv = hv_kv
            elif hv_kv > 0 and not math.isclose(current_base_kv, hv_kv):
                cumulative_z1 = refer_impedance_to_voltage(cumulative_z1, current_base_kv, hv_kv)
                cumulative_z2 = refer_impedance_to_voltage(cumulative_z2, current_base_kv, hv_kv)
                cumulative_z0 = refer_impedance_to_voltage(cumulative_z0, current_base_kv, hv_kv)
                current_base_kv = hv_kv

            if data_mode == "Common Data":
                transformer_z_lv = split_rx_from_z_magnitude(
                    (float(row["Z_pct"]) / 100.0) * ((lv_kv**2) / max(float(row["Rating_MVA"]), 1e-9)),
                    float(row["XR"]),
                )
                element_z1, element_z2, element_z0 = (
                    transformer_z_lv,
                    transformer_z_lv,
                    transformer_z_lv * TRANSFORMER_Z0_FACTOR,
                )
            else:
                element_z1, element_z2, element_z0 = direct_sequence_impedance_from_row(
                    row,
                    z0_multiplier_fallback=TRANSFORMER_Z0_FACTOR,
                    default_z0_to_z1=True,
                )

            cumulative_z1 = refer_impedance_to_voltage(cumulative_z1, current_base_kv, lv_kv)
            cumulative_z2 = refer_impedance_to_voltage(cumulative_z2, current_base_kv, lv_kv)
            cumulative_z0 = refer_impedance_to_voltage(cumulative_z0, current_base_kv, lv_kv)
            current_base_kv = lv_kv
            effective_voltage_kv = lv_kv

            cumulative_z1 += element_z1 / parallel_count
            cumulative_z2 += element_z2 / parallel_count
            cumulative_z0 += element_z0 / parallel_count

        else:
            if current_base_kv is None:
                current_base_kv = voltage_kv
            elif voltage_kv > 0 and not math.isclose(current_base_kv, voltage_kv):
                cumulative_z1 = refer_impedance_to_voltage(cumulative_z1, current_base_kv, voltage_kv)
                cumulative_z2 = refer_impedance_to_voltage(cumulative_z2, current_base_kv, voltage_kv)
                cumulative_z0 = refer_impedance_to_voltage(cumulative_z0, current_base_kv, voltage_kv)
                current_base_kv = voltage_kv

            effective_voltage_kv = current_base_kv if current_base_kv is not None else voltage_kv

            if element_type == "Reactor" and data_mode == "Common Data":
                element_z1 = complex(0.0, float(row["Reactance_Ohm"]))
                element_z2 = element_z1
                element_z0 = element_z1
            elif element_type == "Cable" and data_mode == "Common Data":
                cable_data = cable_sequence_impedance_per_km(str(row["Cable_Type"]))
                element_z1 = float(row["Length_km"]) * cable_data["z1"]
                element_z2 = float(row["Length_km"]) * cable_data["z2"]
                element_z0 = float(row["Length_km"]) * cable_data["z0"]
                cable_rows.append(
                    {
                        "Element": name,
                        "Voltage_kV": effective_voltage_kv,
                        "Cable_Type": str(row["Cable_Type"]),
                        "Parallel_Count": parallel_count,
                        "Per_Cable_Z1": format_complex_ohm(element_z1),
                        "Per_Cable_Z2": format_complex_ohm(element_z2),
                        "Per_Cable_Z0": format_complex_ohm(element_z0),
                        "Equivalent_Z1": format_complex_ohm(element_z1 / parallel_count),
                        "Equivalent_Z2": format_complex_ohm(element_z2 / parallel_count),
                        "Equivalent_Z0": format_complex_ohm(element_z0 / parallel_count),
                    }
                )
            else:
                z0_default_to_z1 = element_type != "Source"
                element_z1, element_z2, element_z0 = direct_sequence_impedance_from_row(
                    row,
                    z0_multiplier_fallback=1.0,
                    default_z0_to_z1=z0_default_to_z1,
                )

            cumulative_z1 += element_z1 / parallel_count
            cumulative_z2 += element_z2 / parallel_count
            cumulative_z0 += element_z0 / parallel_count

        impedance_rows.append(
            {
                "Order": int(row["Order"]),
                "Name": name,
                "Type": element_type,
                "Voltage_kV": effective_voltage_kv,
                "Parallel_Count": parallel_count,
                "Element_Z1": format_complex_ohm(element_z1 / parallel_count if parallel_count > 0 else element_z1),
                "Element_Z2": format_complex_ohm(element_z2 / parallel_count if parallel_count > 0 else element_z2),
                "Element_Z0": format_complex_ohm(element_z0 / parallel_count if parallel_count > 0 else element_z0),
                "Running_Z1": format_complex_ohm(cumulative_z1),
                "Running_Z2": format_complex_ohm(cumulative_z2),
                "Running_Z0": format_complex_ohm(cumulative_z0),
                "|Running_Z1|": abs(cumulative_z1),
                "|Running_Z2|": abs(cumulative_z2),
                "|Running_Z0|": abs(cumulative_z0),
            }
        )

    fault_df = pd.DataFrame(fault_rows)
    return {
        "network_df": network_df,
        "cable_df": pd.DataFrame(cable_rows),
        "impedance_df": pd.DataFrame(impedance_rows),
        "fault_df": fault_df,
        "diagram_dot": build_single_line_diagram_dot(network_df),
        "warnings": warnings,
        "max_fault_current_ka": float(fault_df["I_3ph_kA"].max()) if not fault_df.empty else 0.0,
    }


def parse_bool_like(raw_value: object) -> bool:
    if isinstance(raw_value, bool):
        return raw_value

    normalized_value = str(raw_value).strip().lower()
    if normalized_value in {"1", "true", "yes", "y", "on"}:
        return True
    if normalized_value in {"0", "false", "no", "n", "off"}:
        return False

    raise ValueError(f"Could not parse boolean value from {raw_value!r}.")


def legacy_enclosure_type_to_arc_enclosure(enclosure_type: object) -> str:
    normalized_value = str(enclosure_type).strip().lower()
    if normalized_value == "open air":
        return "Unenclosed"
    return "Enclosed"


def coerce_study_case_value(key: str, raw_value: object, default_value: object) -> object:
    if raw_value is None:
        return default_value

    try:
        if bool(pd.isna(raw_value)):
            return default_value
    except Exception:
        pass

    try:
        if key in STUDY_CASE_BOOL_FIELDS:
            parsed_value = parse_bool_like(raw_value)
        elif key in STUDY_CASE_INT_FIELDS:
            parsed_value = int(float(raw_value))
        elif isinstance(default_value, float):
            parsed_value = float(raw_value)
        else:
            parsed_value = str(raw_value).strip()
    except (ValueError, TypeError):
        return default_value

    if key in STUDY_CASE_BOUNDS:
        lower_bound, upper_bound = STUDY_CASE_BOUNDS[key]
        parsed_value = max(lower_bound, min(parsed_value, upper_bound))
        if key in STUDY_CASE_INT_FIELDS:
            parsed_value = int(parsed_value)

    if key in STUDY_CASE_OPTION_VALUES and parsed_value not in STUDY_CASE_OPTION_VALUES[key]:
        return default_value

    return parsed_value


def load_persisted_state() -> dict:
    if not APP_STATE_FILE.exists():
        return {}

    try:
        with APP_STATE_FILE.open("r", encoding="utf-8") as state_file:
            state_data = json.load(state_file)
        if isinstance(state_data, dict):
            return state_data
    except (OSError, json.JSONDecodeError):
        return {}

    return {}


def persist_last_entered_state() -> None:
    study_case_data = {
        key: st.session_state.get(key, default_value)
        for key, default_value in DEFAULT_STUDY_CASE.items()
    }

    relay_df = sanitize_relay_settings(pd.DataFrame(st.session_state.get("relay_settings", pd.DataFrame())))
    if relay_df.empty:
        relay_df = default_relay_settings_df()

    network_elements_df = sanitize_network_elements(
        pd.DataFrame(st.session_state.get("network_elements", pd.DataFrame()))
    )
    if network_elements_df.empty:
        network_elements_df = default_network_elements_df()

    payload = {
        "study_case": study_case_data,
        "relay_settings": relay_df[RELAY_EXPORT_COLUMNS].to_dict(orient="records"),
        "network_elements": network_elements_df[NETWORK_ELEMENT_COLUMNS].to_dict(orient="records"),
    }

    try:
        APP_STATE_FILE.parent.mkdir(parents=True, exist_ok=True)
        with APP_STATE_FILE.open("w", encoding="utf-8") as state_file:
            json.dump(payload, state_file, indent=2)
    except OSError:
        return


def initialize_app_state() -> None:
    persisted_state = load_persisted_state()
    persisted_study_case = persisted_state.get("study_case", {}) if isinstance(persisted_state, dict) else {}
    if isinstance(persisted_study_case, dict) and "arc_enclosure" not in persisted_study_case:
        legacy_enclosure_type = persisted_study_case.get("enclosure_type")
        if legacy_enclosure_type is not None:
            persisted_study_case = persisted_study_case.copy()
            persisted_study_case["arc_enclosure"] = legacy_enclosure_type_to_arc_enclosure(
                legacy_enclosure_type
            )

    for key, default_value in DEFAULT_STUDY_CASE.items():
        if key not in st.session_state:
            st.session_state[key] = coerce_study_case_value(
                key,
                persisted_study_case.get(key, default_value),
                default_value,
            )

    if "relay_settings" not in st.session_state:
        relay_rows = persisted_state.get("relay_settings", BASELINE_RELAY_SETTINGS)
        relay_df = pd.DataFrame(relay_rows)
        sanitized_df = sanitize_relay_settings(relay_df)
        if sanitized_df.empty:
            sanitized_df = default_relay_settings_df()
        st.session_state.relay_settings = sanitized_df[RELAY_EXPORT_COLUMNS]

    if "network_elements" not in st.session_state:
        network_rows = persisted_state.get("network_elements", DEFAULT_NETWORK_ELEMENTS)
        network_df = pd.DataFrame(network_rows)
        sanitized_network_df = sanitize_network_elements(network_df)
        if sanitized_network_df.empty:
            sanitized_network_df = default_network_elements_df()
        st.session_state.network_elements = sanitized_network_df[NETWORK_ELEMENT_COLUMNS]


def load_network_preset() -> None:
    for key, value in DEFAULT_NETWORK_PRESET.items():
        st.session_state[key] = value
    st.session_state.network_elements = default_network_elements_df()


def restore_original_defaults() -> None:
    load_network_preset()
    st.session_state.relay_settings = default_relay_settings_df()
    st.session_state.pop("relay_settings_editor", None)
    st.session_state.pop("network_elements_editor", None)


def serialize_study_case_csv() -> bytes:
    rows = [
        {
            "parameter": key,
            "value": st.session_state.get(key, default_value),
        }
        for key, default_value in DEFAULT_STUDY_CASE.items()
    ]
    return pd.DataFrame(rows).to_csv(index=False).encode("utf-8")


def normalize_relay_column_name(column_name: str) -> str:
    return (
        str(column_name)
        .strip()
        .lower()
        .replace(" ", "")
        .replace("-", "")
        .replace("_", "")
        .replace("(", "")
        .replace(")", "")
        .replace("/", "")
    )


def parse_study_case_csv(upload_bytes: bytes) -> tuple[bool, str]:
    try:
        case_df = pd.read_csv(io.BytesIO(upload_bytes))
    except Exception as exc:
        return False, f"Study case import failed: {exc}"

    if case_df.empty:
        return False, "Study case import failed: file is empty."

    normalized_columns = {str(col).strip().lower(): col for col in case_df.columns}
    parameter_column = normalized_columns.get("parameter")
    value_column = normalized_columns.get("value")

    if parameter_column is None or value_column is None:
        if len(case_df.columns) < 2:
            return False, "Study case import failed: expected columns 'parameter' and 'value'."
        parameter_column, value_column = case_df.columns[0], case_df.columns[1]

    parameter_map = {}
    for _, row in case_df.iterrows():
        parameter_key = str(row[parameter_column]).strip()
        if parameter_key:
            parameter_map[parameter_key] = row[value_column]

    if "arc_enclosure" not in parameter_map and "enclosure_type" in parameter_map:
        parameter_map["arc_enclosure"] = legacy_enclosure_type_to_arc_enclosure(
            parameter_map["enclosure_type"]
        )

    updates: dict[str, object] = {}

    for key, default_value in DEFAULT_STUDY_CASE.items():
        if key not in parameter_map:
            continue

        raw_value = parameter_map[key]
        if pd.isna(raw_value):
            continue

        try:
            if key in STUDY_CASE_BOOL_FIELDS:
                parsed_value = parse_bool_like(raw_value)
            elif key in STUDY_CASE_INT_FIELDS:
                parsed_value = int(float(raw_value))
            elif isinstance(default_value, float):
                parsed_value = float(raw_value)
            else:
                parsed_value = str(raw_value).strip()
        except (ValueError, TypeError):
            continue

        if key in STUDY_CASE_BOUNDS:
            lower_bound, upper_bound = STUDY_CASE_BOUNDS[key]
            parsed_value = max(lower_bound, min(parsed_value, upper_bound))
            if key in STUDY_CASE_INT_FIELDS:
                parsed_value = int(parsed_value)

        if key in STUDY_CASE_OPTION_VALUES and parsed_value not in STUDY_CASE_OPTION_VALUES[key]:
            continue

        updates[key] = parsed_value

    if not updates:
        return False, "Study case import failed: no recognized parameters found."

    for key, value in updates.items():
        st.session_state[key] = value

    return True, f"Loaded {len(updates)} study parameters from CSV."


def parse_relay_settings_csv(upload_bytes: bytes) -> Tuple[Optional[pd.DataFrame], str]:
    try:
        relay_df = pd.read_csv(io.BytesIO(upload_bytes))
    except Exception as exc:
        return None, f"Relay settings import failed: {exc}"

    if relay_df.empty:
        return None, "Relay settings import failed: file is empty."

    rename_map = {}
    for column_name in relay_df.columns:
        normalized_name = normalize_relay_column_name(column_name)
        target_column = RELAY_COLUMN_ALIASES.get(normalized_name)
        if target_column is not None and target_column not in rename_map.values():
            rename_map[column_name] = target_column

    relay_df = relay_df.rename(columns=rename_map)
    missing_columns = [column for column in RELAY_EXPORT_COLUMNS if column not in relay_df.columns]

    if missing_columns:
        missing_list = ", ".join(missing_columns)
        return None, f"Relay settings import failed: missing columns [{missing_list}]."

    sanitized_df = sanitize_relay_settings(relay_df[RELAY_EXPORT_COLUMNS])
    if sanitized_df.empty:
        return None, "Relay settings import failed: no valid rows after validation."

    return sanitized_df, f"Loaded {len(sanitized_df)} relay settings rows from CSV."


def initialize_commissioning_widget_state() -> None:
    for key, value in COMMISSIONING_WIDGET_DEFAULTS.items():
        if key not in st.session_state:
            st.session_state[key] = value


st.set_page_config(
    page_title=f"Protection {APP_REVISION}",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded",
)

initialize_app_state()
initialize_commissioning_widget_state()

st.title("🛡️ Protection")
st.caption(
    f"{APP_REVISION} · ETAP-style protection study MVP: network parameters, fault levels, protection grading, commissioning worksheets, and arc-flash screening."
)

with st.sidebar:
    st.header("Study Case CSV")
    st.download_button(
        "Export study case CSV",
        data=serialize_study_case_csv(),
        file_name="study_case.csv",
        mime="text/csv",
        use_container_width=True,
    )

    study_case_upload = st.file_uploader(
        "Import study case CSV",
        type=["csv"],
        key="study_case_upload",
    )
    if study_case_upload is not None:
        upload_bytes = study_case_upload.getvalue()
        upload_digest = hashlib.md5(upload_bytes).hexdigest()
        if st.session_state.get("study_case_upload_digest") != upload_digest:
            import_ok, import_message = parse_study_case_csv(upload_bytes)
            st.session_state["study_case_upload_digest"] = upload_digest
            st.session_state["study_case_upload_status"] = "success" if import_ok else "error"
            st.session_state["study_case_upload_message"] = import_message

    if st.session_state.get("study_case_upload_message"):
        if st.session_state.get("study_case_upload_status") == "success":
            st.success(st.session_state["study_case_upload_message"])
        else:
            st.error(st.session_state["study_case_upload_message"])

    st.markdown("---")
    st.header("Study Controls")
    project_name = st.text_input("Project name", key="project_name")
    grading_margin_s = st.number_input(
        "Target grading margin (s)",
        min_value=0.10,
        max_value=1.00,
        step=0.05,
        key="grading_margin_s",
    )

    if st.button("Restore original defaults", use_container_width=True):
        restore_original_defaults()
        persist_last_entered_state()
        st.session_state["study_case_upload_status"] = "success"
        st.session_state["study_case_upload_message"] = "Restored original default network and relay settings."
        st.rerun()

    st.caption(f"Active study: {project_name}")

network_tab, fault_tab, protection_tab, arc_tab, full_arc_tab, oc_ef_commissioning_tab, sel787_tab, translay_tab, formula_tab = st.tabs(
    [
        "1) Network Inputs",
        "2) Fault Levels",
        "3) Protection & Grading",
        "4) Arc-Flash Estimate",
        "5) Full Arc-Flash Calc",
        "6) OC/EF Commissioning",
        "7) SEL-787 Differential",
        "8) Translay Commissioning",
        "9) Formula Reference",
    ]
)

frequency_hz = st.session_state.get("frequency_hz", 50)
source_v_kv = float(st.session_state.get("source_v_kv", DEFAULT_NETWORK_PRESET["source_v_kv"]))
transformer_hv_kv = float(
    st.session_state.get("transformer_hv_kv", DEFAULT_NETWORK_PRESET["transformer_hv_kv"])
)
transformer_lv_kv = float(
    st.session_state.get("transformer_lv_kv", DEFAULT_NETWORK_PRESET["transformer_lv_kv"])
)
max_fault_current_ka = 0.0

with network_tab:
    st.subheader("Power Network Parameters")
    network_input_mode = st.radio(
        "Network input mode",
        options=NETWORK_INPUT_MODE_OPTIONS,
        key="network_input_mode",
        horizontal=True,
    )

    if network_input_mode == "Template Inputs":
        st.caption("All impedances are converted to the selected study voltage base before fault calculations.")
        st.markdown("#### One-line arrangement (default template)")
        st.caption(" → ".join(NETWORK_ARRANGEMENT))
        st.info(
            "Legacy cable sequence data uses best-available typical values for old assets. "
            "Where detailed test data is unavailable, Z2 is taken equal to Z1 and typical Z0 values are used."
        )
        if st.session_state.get("duplicate_source_to_first_busbar_circuit", False):
            st.caption(
                "Optional duplicate circuit enabled: reactor, 33kV cable, transformer, and 11kV transformer cable "
                "are modeled as two identical circuits in parallel from the 33kV source bus to the first 11kV busbar. "
                "The upstream source equivalent remains common."
            )
        if st.session_state.get("parallel_lv_transformer_cable", False):
            st.caption(
                "11kV transformer cable is modeled as two cables in parallel (impedance halved for this segment)."
            )
        if st.session_state.get("parallel_feeder_cable", False):
            st.caption(
                "11kV feeder cable is modeled as two cables in parallel (impedance halved for this segment)."
            )

        col1, col2, col3 = st.columns(3)

        with col1:
            st.markdown("**33kV Source Side**")
            source_v_kv = st.number_input(
                "Source nominal voltage (kV)",
                min_value=1.0,
                max_value=400.0,
                step=0.1,
                key="source_v_kv",
            )
            source_sc_mva = st.number_input(
                "Upstream source short-circuit level (MVA)",
                min_value=10.0,
                max_value=50000.0,
                step=10.0,
                key="source_sc_mva",
            )
            source_xr = st.number_input(
                "Source X/R ratio",
                min_value=0.0,
                max_value=50.0,
                step=0.5,
                key="source_xr",
            )
            reactor_ohm = st.number_input(
                "33kV reactor reactance (Ω)",
                min_value=0.0,
                max_value=50.0,
                step=0.01,
                key="reactor_ohm",
            )
            hv_cable_length_km = st.number_input(
                "33kV cable length (km)",
                min_value=0.0,
                max_value=100.0,
                step=0.1,
                key="hv_cable_length_km",
            )
            hv_cable_type = st.selectbox(
                "33kV cable type",
                options=HV_CABLE_OPTIONS,
                key="hv_cable_type",
            )

        with col2:
            st.markdown("**Transformer & 11kV Link**")
            transformer_hv_kv = st.number_input(
                "Transformer HV voltage (kV)",
                min_value=1.0,
                max_value=400.0,
                step=0.1,
                key="transformer_hv_kv",
            )
            transformer_lv_kv = st.number_input(
                "Transformer LV voltage (kV)",
                min_value=0.4,
                max_value=66.0,
                step=0.1,
                key="transformer_lv_kv",
            )
            transformer_mva = st.number_input(
                "Transformer rating (MVA)",
                min_value=0.1,
                max_value=1000.0,
                step=0.1,
                key="transformer_mva",
            )
            transformer_z_pct = st.number_input(
                "Transformer impedance (%Z)",
                min_value=1.0,
                max_value=25.0,
                step=0.1,
                key="transformer_z_pct",
            )
            transformer_xr = st.number_input(
                "Transformer X/R ratio",
                min_value=0.0,
                max_value=50.0,
                step=0.5,
                key="transformer_xr",
            )
            lv_cable_length_km = st.number_input(
                "11kV cable length from transformer (km)",
                min_value=0.0,
                max_value=20.0,
                step=0.01,
                key="lv_cable_length_km",
            )
            lv_cable_type = st.selectbox(
                "11kV transformer cable type",
                options=LV_CABLE_OPTIONS,
                key="lv_cable_type",
            )
            parallel_lv_transformer_cable = st.checkbox(
                "Parallel 11kV transformer cable (2 in parallel)",
                key="parallel_lv_transformer_cable",
            )
            duplicate_source_to_first_busbar_circuit = st.checkbox(
                "Duplicate source-to-1st-11kV-busbar circuit in parallel",
                key="duplicate_source_to_first_busbar_circuit",
            )

        with col3:
            st.markdown("**11kV Study Side**")
            v_ll_kv = st.number_input(
                "Study bus nominal voltage (kV)",
                min_value=0.4,
                max_value=66.0,
                step=0.1,
                key="v_ll_kv",
            )
            frequency_hz = st.selectbox("Frequency (Hz)", options=[50, 60], key="frequency_hz")
            feeder_length_km = st.number_input(
                "Feeder length (km)",
                min_value=0.0,
                max_value=100.0,
                step=0.1,
                key="feeder_length_km",
            )
            feeder_cable_type = st.selectbox(
                "11kV feeder cable type",
                options=LV_CABLE_OPTIONS,
                key="feeder_cable_type",
            )
            parallel_feeder_cable = st.checkbox(
                "Parallel 11kV feeder cable (2 in parallel)",
                key="parallel_feeder_cable",
            )
    else:
        st.caption(
            "Enter one network element per row. Busbar rows become fault locations. "
            "Use common data for typical source/cable/reactor/transformer inputs, or direct sequence impedance where available."
        )
        st.info(
            "Parallel circuits are modeled with `Parallel_Count`. Each row also carries its own voltage input so the builder can step across voltage levels line by line."
        )
        builder_col1, builder_col2 = st.columns([5, 1])
        with builder_col2:
            frequency_hz = st.selectbox("Frequency (Hz)", options=[50, 60], key="frequency_hz")

        network_elements_editor_df = pd.DataFrame(st.session_state.network_elements).reset_index(drop=True)
        if network_elements_editor_df.empty:
            network_elements_editor_df = default_network_elements_df()
        else:
            for column_name in NETWORK_ELEMENT_COLUMNS:
                if column_name not in network_elements_editor_df.columns:
                    if column_name in {"Name", "Element_Type", "Data_Mode", "Cable_Type"}:
                        network_elements_editor_df[column_name] = ""
                    else:
                        network_elements_editor_df[column_name] = 0.0
            network_elements_editor_df = network_elements_editor_df[NETWORK_ELEMENT_COLUMNS]

        edited_network_elements_df = st.data_editor(
            network_elements_editor_df,
            num_rows="dynamic",
            key="network_elements_editor",
            use_container_width=True,
            height=380,
            column_config={
                "Order": st.column_config.NumberColumn("Order", min_value=1, step=1),
                "Name": st.column_config.TextColumn("Name"),
                "Element_Type": st.column_config.SelectboxColumn(
                    "Type", options=NETWORK_ELEMENT_TYPE_OPTIONS
                ),
                "Data_Mode": st.column_config.SelectboxColumn(
                    "Data mode", options=NETWORK_ELEMENT_DATA_MODE_OPTIONS
                ),
                "Voltage_kV": st.column_config.NumberColumn("Voltage (kV)", min_value=0.0),
                "Parallel_Count": st.column_config.NumberColumn("Parallel", min_value=1, step=1),
                "Source_SC_MVA": st.column_config.NumberColumn("Source SC (MVA)", min_value=0.0),
                "XR": st.column_config.NumberColumn("X/R", min_value=0.0),
                "Reactance_Ohm": st.column_config.NumberColumn("Reactor X (Ω)", min_value=0.0),
                "Length_km": st.column_config.NumberColumn("Length (km)", min_value=0.0),
                "Cable_Type": st.column_config.SelectboxColumn(
                    "Cable type", options=sorted(set(HV_CABLE_OPTIONS) | set(LV_CABLE_OPTIONS))
                ),
                "HV_kV": st.column_config.NumberColumn("HV (kV)", min_value=0.0),
                "LV_kV": st.column_config.NumberColumn("LV (kV)", min_value=0.0),
                "Rating_MVA": st.column_config.NumberColumn("Rating (MVA)", min_value=0.0),
                "Z_pct": st.column_config.NumberColumn("%Z", min_value=0.0),
                "R1_ohm": st.column_config.NumberColumn("R1 (Ω)", min_value=0.0),
                "X1_ohm": st.column_config.NumberColumn("X1 (Ω)", min_value=0.0),
                "R0_ohm": st.column_config.NumberColumn("R0 (Ω)", min_value=0.0),
                "X0_ohm": st.column_config.NumberColumn("X0 (Ω)", min_value=0.0),
            },
        )
        st.session_state.network_elements = pd.DataFrame(edited_network_elements_df).reset_index(drop=True)

cable_library_df = pd.DataFrame(
    [
        {
            "Cable Type": cable_type,
            "Z1": format_complex_ohm(cable_data["z1"]),
            "Z2": format_complex_ohm(cable_data["z2"]),
            "Z0": format_complex_ohm(cable_data["z0"]),
            "Data basis": cable_data["note"],
        }
        for cable_type, cable_data in CABLE_SEQUENCE_LIBRARY.items()
    ]
)

if network_input_mode == "Template Inputs":
    source_z_hv = split_rx_from_z_magnitude((source_v_kv**2) / source_sc_mva, source_xr)
    source_to_study_factor = (v_ll_kv / source_v_kv) ** 2 if source_v_kv > 0 else 0.0
    source_z1 = source_z_hv * source_to_study_factor
    source_z2 = source_z1
    source_z0 = source_z1 * SOURCE_Z0_FACTOR

    reactor_z1 = complex(0.0, reactor_ohm) * source_to_study_factor
    reactor_z2 = reactor_z1
    reactor_z0 = reactor_z1

    hv_cable_data = cable_sequence_impedance_per_km(hv_cable_type)
    hv_cable_z1 = hv_cable_length_km * hv_cable_data["z1"] * source_to_study_factor
    hv_cable_z2 = hv_cable_length_km * hv_cable_data["z2"] * source_to_study_factor
    hv_cable_z0 = hv_cable_length_km * hv_cable_data["z0"] * source_to_study_factor

    transformer_z_lv = split_rx_from_z_magnitude(
        (transformer_z_pct / 100) * ((transformer_lv_kv**2) / transformer_mva),
        transformer_xr,
    )
    lv_to_study_factor = (v_ll_kv / transformer_lv_kv) ** 2 if transformer_lv_kv > 0 else 0.0
    transformer_z1 = transformer_z_lv * lv_to_study_factor
    transformer_z2 = transformer_z1
    transformer_z0 = transformer_z1 * TRANSFORMER_Z0_FACTOR

    lv_cable_data = cable_sequence_impedance_per_km(lv_cable_type)
    lv_cable_z1 = lv_cable_length_km * lv_cable_data["z1"] * lv_to_study_factor
    lv_cable_z2 = lv_cable_length_km * lv_cable_data["z2"] * lv_to_study_factor
    lv_cable_z0 = lv_cable_length_km * lv_cable_data["z0"] * lv_to_study_factor

    lv_transformer_cable_count = 2 if parallel_lv_transformer_cable else 1
    lv_cable_equiv_z1 = lv_cable_z1 / lv_transformer_cable_count
    lv_cable_equiv_z2 = lv_cable_z2 / lv_transformer_cable_count
    lv_cable_equiv_z0 = lv_cable_z0 / lv_transformer_cable_count

    single_source_to_first_busbar_circuit_z1 = reactor_z1 + hv_cable_z1 + transformer_z1 + lv_cable_equiv_z1
    single_source_to_first_busbar_circuit_z2 = reactor_z2 + hv_cable_z2 + transformer_z2 + lv_cable_equiv_z2
    single_source_to_first_busbar_circuit_z0 = reactor_z0 + hv_cable_z0 + transformer_z0 + lv_cable_equiv_z0

    parallel_circuit_count = 2 if duplicate_source_to_first_busbar_circuit else 1
    source_to_first_busbar_equiv_z1 = single_source_to_first_busbar_circuit_z1 / parallel_circuit_count
    source_to_first_busbar_equiv_z2 = single_source_to_first_busbar_circuit_z2 / parallel_circuit_count
    source_to_first_busbar_equiv_z0 = single_source_to_first_busbar_circuit_z0 / parallel_circuit_count

    feeder_cable_data = cable_sequence_impedance_per_km(feeder_cable_type)
    feeder_z1 = feeder_length_km * feeder_cable_data["z1"]
    feeder_z2 = feeder_length_km * feeder_cable_data["z2"]
    feeder_z0 = feeder_length_km * feeder_cable_data["z0"]

    feeder_cable_count = 2 if parallel_feeder_cable else 1
    feeder_equiv_z1 = feeder_z1 / feeder_cable_count
    feeder_equiv_z2 = feeder_z2 / feeder_cable_count
    feeder_equiv_z0 = feeder_z0 / feeder_cable_count

    z1_incomer = source_z1 + source_to_first_busbar_equiv_z1
    z2_incomer = source_z2 + source_to_first_busbar_equiv_z2
    z0_incomer = source_z0 + source_to_first_busbar_equiv_z0

    z1_load = z1_incomer + feeder_equiv_z1
    z2_load = z2_incomer + feeder_equiv_z2
    z0_load = z0_incomer + feeder_equiv_z0

    incomer_fault = calc_fault_values(v_ll_kv, z1_incomer, z2_incomer, z0_incomer)
    load_fault = calc_fault_values(v_ll_kv, z1_load, z2_load, z0_load)

    cable_df = pd.DataFrame(
        [
            {
                "Segment": "33kV cable",
                "Type": hv_cable_type,
                "Z1": format_complex_ohm(hv_cable_data["z1"]),
                "Z2": format_complex_ohm(hv_cable_data["z2"]),
                "Z0": format_complex_ohm(hv_cable_data["z0"]),
            },
            {
                "Segment": "11kV transformer cable",
                "Type": lv_cable_type,
                "Z1": format_complex_ohm(lv_cable_data["z1"]),
                "Z2": format_complex_ohm(lv_cable_data["z2"]),
                "Z0": format_complex_ohm(lv_cable_data["z0"]),
            },
            {
                "Segment": "11kV feeder cable",
                "Type": feeder_cable_type,
                "Z1": format_complex_ohm(feeder_cable_data["z1"]),
                "Z2": format_complex_ohm(feeder_cable_data["z2"]),
                "Z0": format_complex_ohm(feeder_cable_data["z0"]),
            },
        ]
    )

    impedance_rows = [
        {
            "Element": "Source equivalent (referred)",
            "Z1 (Ω)": format_complex_ohm(source_z1),
            "Z2 (Ω)": format_complex_ohm(source_z2),
            "Z0 (Ω)": format_complex_ohm(source_z0),
            "|Z1|": abs(source_z1),
            "|Z2|": abs(source_z2),
            "|Z0|": abs(source_z0),
        },
        {
            "Element": "33kV reactor (per circuit, referred)",
            "Z1 (Ω)": format_complex_ohm(reactor_z1),
            "Z2 (Ω)": format_complex_ohm(reactor_z2),
            "Z0 (Ω)": format_complex_ohm(reactor_z0),
            "|Z1|": abs(reactor_z1),
            "|Z2|": abs(reactor_z2),
            "|Z0|": abs(reactor_z0),
        },
        {
            "Element": "33kV cable (per circuit, referred)",
            "Z1 (Ω)": format_complex_ohm(hv_cable_z1),
            "Z2 (Ω)": format_complex_ohm(hv_cable_z2),
            "Z0 (Ω)": format_complex_ohm(hv_cable_z0),
            "|Z1|": abs(hv_cable_z1),
            "|Z2|": abs(hv_cable_z2),
            "|Z0|": abs(hv_cable_z0),
        },
        {
            "Element": "Transformer equivalent (per circuit, referred)",
            "Z1 (Ω)": format_complex_ohm(transformer_z1),
            "Z2 (Ω)": format_complex_ohm(transformer_z2),
            "Z0 (Ω)": format_complex_ohm(transformer_z0),
            "|Z1|": abs(transformer_z1),
            "|Z2|": abs(transformer_z2),
            "|Z0|": abs(transformer_z0),
        },
        {
            "Element": "11kV cable from transformer (per cable)",
            "Z1 (Ω)": format_complex_ohm(lv_cable_z1),
            "Z2 (Ω)": format_complex_ohm(lv_cable_z2),
            "Z0 (Ω)": format_complex_ohm(lv_cable_z0),
            "|Z1|": abs(lv_cable_z1),
            "|Z2|": abs(lv_cable_z2),
            "|Z0|": abs(lv_cable_z0),
        },
        {
            "Element": "Single circuit subtotal: 33kV source bus to 1st 11kV busbar",
            "Z1 (Ω)": format_complex_ohm(single_source_to_first_busbar_circuit_z1),
            "Z2 (Ω)": format_complex_ohm(single_source_to_first_busbar_circuit_z2),
            "Z0 (Ω)": format_complex_ohm(single_source_to_first_busbar_circuit_z0),
            "|Z1|": abs(single_source_to_first_busbar_circuit_z1),
            "|Z2|": abs(single_source_to_first_busbar_circuit_z2),
            "|Z0|": abs(single_source_to_first_busbar_circuit_z0),
        },
    ]
    if parallel_lv_transformer_cable:
        impedance_rows.append(
            {
                "Element": "Equivalent of 2 parallel 11kV transformer cables (per source circuit)",
                "Z1 (Ω)": format_complex_ohm(lv_cable_equiv_z1),
                "Z2 (Ω)": format_complex_ohm(lv_cable_equiv_z2),
                "Z0 (Ω)": format_complex_ohm(lv_cable_equiv_z0),
                "|Z1|": abs(lv_cable_equiv_z1),
                "|Z2|": abs(lv_cable_equiv_z2),
                "|Z0|": abs(lv_cable_equiv_z0),
            }
        )
    if duplicate_source_to_first_busbar_circuit:
        impedance_rows.append(
            {
                "Element": "Equivalent of 2 parallel source-to-1st-11kV-busbar circuits",
                "Z1 (Ω)": format_complex_ohm(source_to_first_busbar_equiv_z1),
                "Z2 (Ω)": format_complex_ohm(source_to_first_busbar_equiv_z2),
                "Z0 (Ω)": format_complex_ohm(source_to_first_busbar_equiv_z0),
                "|Z1|": abs(source_to_first_busbar_equiv_z1),
                "|Z2|": abs(source_to_first_busbar_equiv_z2),
                "|Z0|": abs(source_to_first_busbar_equiv_z0),
            }
        )
    impedance_rows.append(
        {
            "Element": "11kV feeder (per cable)",
            "Z1 (Ω)": format_complex_ohm(feeder_z1),
            "Z2 (Ω)": format_complex_ohm(feeder_z2),
            "Z0 (Ω)": format_complex_ohm(feeder_z0),
            "|Z1|": abs(feeder_z1),
            "|Z2|": abs(feeder_z2),
            "|Z0|": abs(feeder_z0),
        }
    )
    if parallel_feeder_cable:
        impedance_rows.append(
            {
                "Element": "Equivalent of 2 parallel 11kV feeder cables",
                "Z1 (Ω)": format_complex_ohm(feeder_equiv_z1),
                "Z2 (Ω)": format_complex_ohm(feeder_equiv_z2),
                "Z0 (Ω)": format_complex_ohm(feeder_equiv_z0),
                "|Z1|": abs(feeder_equiv_z1),
                "|Z2|": abs(feeder_equiv_z2),
                "|Z0|": abs(feeder_equiv_z0),
            }
        )
    impedance_rows.extend(
        [
            {
                "Element": "Total to 11kV transformer busbar",
                "Z1 (Ω)": format_complex_ohm(z1_incomer),
                "Z2 (Ω)": format_complex_ohm(z2_incomer),
                "Z0 (Ω)": format_complex_ohm(z0_incomer),
                "|Z1|": abs(z1_incomer),
                "|Z2|": abs(z2_incomer),
                "|Z0|": abs(z0_incomer),
            },
            {
                "Element": "Total to 11kV remote busbar",
                "Z1 (Ω)": format_complex_ohm(z1_load),
                "Z2 (Ω)": format_complex_ohm(z2_load),
                "Z0 (Ω)": format_complex_ohm(z0_load),
                "|Z1|": abs(z1_load),
                "|Z2|": abs(z2_load),
                "|Z0|": abs(z0_load),
            },
        ]
    )
    impedance_df = pd.DataFrame(impedance_rows)
    fault_df = pd.DataFrame(
        [
            {
                "Bus": "11kV Transformer Busbar",
                "Z1_total_ohm": incomer_fault["Z1_total_ohm"],
                "Z2_total_ohm": incomer_fault["Z2_total_ohm"],
                "Z0_total_ohm": incomer_fault["Z0_total_ohm"],
                "I_3ph_kA": incomer_fault["I_3ph_kA"],
                "I_LL_kA": incomer_fault["I_LL_kA"],
                "I_LG_kA": incomer_fault["I_LG_kA"],
                "Fault_MVA": incomer_fault["Fault_MVA"],
            },
            {
                "Bus": "11kV Remote Busbar",
                "Z1_total_ohm": load_fault["Z1_total_ohm"],
                "Z2_total_ohm": load_fault["Z2_total_ohm"],
                "Z0_total_ohm": load_fault["Z0_total_ohm"],
                "I_3ph_kA": load_fault["I_3ph_kA"],
                "I_LL_kA": load_fault["I_LL_kA"],
                "I_LG_kA": load_fault["I_LG_kA"],
                "Fault_MVA": load_fault["Fault_MVA"],
            },
        ]
    )
    network_model_warnings = []
    network_diagram_dot = build_single_line_diagram_dot(default_network_elements_df())
    max_fault_current_ka = float(fault_df["I_3ph_kA"].max()) if not fault_df.empty else 0.0
else:
    builder_results = calculate_line_builder_network(pd.DataFrame(st.session_state.network_elements))
    cable_df = builder_results["cable_df"]
    impedance_df = builder_results["impedance_df"]
    fault_df = builder_results["fault_df"]
    network_model_warnings = builder_results["warnings"]
    network_diagram_dot = builder_results["diagram_dot"]
    max_fault_current_ka = float(builder_results["max_fault_current_ka"])

with network_tab:
    if network_input_mode == "Template Inputs":
        st.markdown("#### Selected Cable Sequence Data (Ω/km)")
        st.dataframe(cable_df, use_container_width=True)
        st.markdown("#### Cable Library (Best Available Legacy Data, Ω/km)")
        st.dataframe(cable_library_df, use_container_width=True)
        st.markdown("#### Sequence Impedance Build-Up (Referred to Study Voltage)")
        st.dataframe(
            impedance_df.style.format({"|Z1|": "{:.4f}", "|Z2|": "{:.4f}", "|Z0|": "{:.4f}"}),
            use_container_width=True,
        )
        st.caption(
            f"Nominal frequency: {frequency_hz} Hz | Source base: {source_v_kv:.1f} kV | Transformer: {transformer_hv_kv:.1f}/{transformer_lv_kv:.1f} kV"
        )
        st.caption(
            f"Assumed source Z0/Z1 = {SOURCE_Z0_FACTOR:.1f} and transformer Z0/Z1 = {TRANSFORMER_Z0_FACTOR:.1f}."
        )
    else:
        if network_model_warnings:
            for warning_message in network_model_warnings:
                st.warning(warning_message)

        st.markdown("#### Single-Line Diagram")
        st.graphviz_chart(network_diagram_dot, use_container_width=True)

        st.markdown("#### Builder Row Sequence Data")
        if cable_df.empty:
            st.info("No cable rows are currently defined in the line-by-line builder.")
        else:
            st.dataframe(cable_df, use_container_width=True)

        st.markdown("#### Cable Library (Best Available Legacy Data, Ω/km)")
        st.dataframe(cable_library_df, use_container_width=True)

        st.markdown("#### Running Sequence Impedance Build-Up")
        if impedance_df.empty:
            st.info("Add network rows to calculate impedance build-up.")
        else:
            st.dataframe(
                impedance_df.style.format(
                    {
                        "Voltage_kV": "{:.2f}",
                        "|Running_Z1|": "{:.4f}",
                        "|Running_Z2|": "{:.4f}",
                        "|Running_Z0|": "{:.4f}",
                    }
                ),
                use_container_width=True,
            )

if not fault_df.empty:
    available_study_buses = fault_df["Bus"].tolist()
    if st.session_state.get("study_bus") not in available_study_buses:
        st.session_state["study_bus"] = available_study_buses[0]
else:
    available_study_buses = []
    st.session_state["study_bus"] = ""

with fault_tab:
    st.subheader("Calculated Fault Levels")
    if fault_df.empty:
        st.info(
            "No fault locations are currently available. Add at least one Busbar row in the line-by-line builder or use the template inputs."
        )
    else:
        st.dataframe(
            fault_df.style.format(
                {
                    "Z1_total_ohm": "{:.4f}",
                    "Z2_total_ohm": "{:.4f}",
                    "Z0_total_ohm": "{:.4f}",
                    "I_3ph_kA": "{:.2f}",
                    "I_LL_kA": "{:.2f}",
                    "I_LG_kA": "{:.2f}",
                    "Fault_MVA": "{:.1f}",
                }
            ),
            use_container_width=True,
        )

    study_col1, study_col2 = st.columns(2)
    with study_col1:
        if available_study_buses:
            study_bus = st.selectbox(
                "Select fault location for protection checks",
                options=available_study_buses,
                key="study_bus",
            )
        else:
            study_bus = ""
            st.text_input(
                "Select fault location for protection checks",
                value="No busbars available",
                disabled=True,
            )
    with study_col2:
        fault_type = st.selectbox(
            "Select fault type",
            options=["3-Phase", "Line-Line", "Line-Ground"],
            key="fault_type",
        )

fault_column_map = {
    "3-Phase": "I_3ph_kA",
    "Line-Line": "I_LL_kA",
    "Line-Ground": "I_LG_kA",
}

if available_study_buses:
    selected_fault_current_ka = float(
        fault_df.loc[fault_df["Bus"] == study_bus, fault_column_map[fault_type]].iloc[0]
    )
else:
    selected_fault_current_ka = 0.0
selected_fault_current_a = selected_fault_current_ka * 1000

with fault_tab:
    metric_col1, metric_col2 = st.columns(2)
    with metric_col1:
        st.metric("Selected fault current", f"{selected_fault_current_ka:.2f} kA")
    with metric_col2:
        st.metric("Selected fault type", fault_type)

with protection_tab:
    st.subheader("Protection Device Settings")

    relay_export_df = sanitize_relay_settings(pd.DataFrame(st.session_state.relay_settings))
    if relay_export_df.empty:
        relay_export_df = pd.DataFrame(columns=RELAY_EXPORT_COLUMNS)

    relay_io_col1, relay_io_col2 = st.columns(2)
    with relay_io_col1:
        st.download_button(
            "Export relay settings CSV",
            data=relay_export_df[RELAY_EXPORT_COLUMNS].to_csv(index=False).encode("utf-8"),
            file_name="relay_settings.csv",
            mime="text/csv",
            use_container_width=True,
        )
    with relay_io_col2:
        relay_upload = st.file_uploader(
            "Import relay settings CSV",
            type=["csv"],
            key="relay_settings_upload",
        )

    if relay_upload is not None:
        relay_upload_bytes = relay_upload.getvalue()
        relay_upload_digest = hashlib.md5(relay_upload_bytes).hexdigest()
        if st.session_state.get("relay_upload_digest") != relay_upload_digest:
            imported_relay_df, relay_message = parse_relay_settings_csv(relay_upload_bytes)
            st.session_state["relay_upload_digest"] = relay_upload_digest
            if imported_relay_df is not None:
                st.session_state.relay_settings = imported_relay_df.reset_index(drop=True)
                st.session_state.pop("relay_settings_editor", None)
                st.session_state["relay_upload_status"] = "success"
            else:
                st.session_state["relay_upload_status"] = "error"
            st.session_state["relay_upload_message"] = relay_message

    if st.session_state.get("relay_upload_message"):
        if st.session_state.get("relay_upload_status") == "success":
            st.success(st.session_state["relay_upload_message"])
        else:
            st.error(st.session_state["relay_upload_message"])

    relay_editor_df = pd.DataFrame(st.session_state.relay_settings).reset_index(drop=True)
    if relay_editor_df.empty:
        relay_editor_df = default_relay_settings_df()

    edited_df = st.data_editor(
        relay_editor_df,
        num_rows="dynamic",
        key="relay_settings_editor",
        use_container_width=True,
        column_config={
            "Order": st.column_config.NumberColumn("Order", min_value=1, step=1),
            "Device": st.column_config.TextColumn("Device"),
            "Curve": st.column_config.SelectboxColumn("Curve", options=list(CURVE_CONSTANTS.keys())),
            "Pickup_A": st.column_config.NumberColumn("Pickup (A)", min_value=1.0),
            "TMS": st.column_config.NumberColumn("TMS", min_value=0.01, max_value=1.0),
            "Inst_A": st.column_config.NumberColumn("Instantaneous pickup (A)", min_value=0.0),
        },
    )
    edited_df = pd.DataFrame(edited_df).reset_index(drop=True)
    st.session_state.relay_settings = edited_df

    relay_df = sanitize_relay_settings(pd.DataFrame(edited_df))

    if relay_df.empty:
        st.warning("Add at least one valid protection device setting to run grading checks.")
        relay_results_df = pd.DataFrame()
    elif max_fault_current_ka <= 0:
        st.warning(
            "No fault current results are available yet. Add at least one valid fault location to run grading checks."
        )
        relay_results_df = pd.DataFrame()
    else:
        min_pickup = max(float(relay_df["Pickup_A"].min()) * 1.05, 1.0)
        max_fault_current_a = max_fault_current_ka * 1000
        max_current = max(
            max_fault_current_a * 2.0,
            float(relay_df["Inst_A"].max()) * 1.20,
            min_pickup * 5,
        )

        slider_min_a = max(1, int(math.floor(min_pickup)))
        slider_max_a = max(slider_min_a + 1, int(math.ceil(max_current)))
        slider_default_a = int(min(max(selected_fault_current_a, slider_min_a), slider_max_a))
        slider_step_a = max(1, int((slider_max_a - slider_min_a) / 200))

        grading_current_a = float(
            st.slider(
                "Grading current on TCC graph (A)",
                min_value=slider_min_a,
                max_value=slider_max_a,
                value=slider_default_a,
                step=slider_step_a,
            )
        )
        st.caption(
            f"Relay operating times and grading margins are evaluated at {grading_current_a / 1000:.3f} kA."
        )

        relay_results = []
        for _, row in relay_df.iterrows():
            operate_time_s = calc_relay_time_s(
                current_a=grading_current_a,
                pickup_a=float(row["Pickup_A"]),
                tms=float(row["TMS"]),
                curve_name=str(row["Curve"]),
                inst_pickup_a=float(row["Inst_A"]),
            )
            relay_results.append(
                {
                    "Order": int(row["Order"]),
                    "Device": str(row["Device"]),
                    "Curve": str(row["Curve"]),
                    "Pickup_A": float(row["Pickup_A"]),
                    "TMS": float(row["TMS"]),
                    "Inst_A": float(row["Inst_A"]),
                    "Operate_s": operate_time_s,
                    "Operate_display": format_time_value(operate_time_s),
                }
            )

        relay_results_df = pd.DataFrame(relay_results).sort_values("Order").reset_index(drop=True)

        st.markdown("#### Trip Time Readout at Slider Current")
        trip_readout_df = relay_results_df[
            ["Order", "Device", "Curve", "Pickup_A", "TMS", "Inst_A", "Operate_s", "Operate_display"]
        ].copy()
        trip_readout_df["Operate_s"] = trip_readout_df["Operate_s"].apply(
            lambda value: value if math.isfinite(float(value)) else math.nan
        )
        st.dataframe(
            trip_readout_df.style.format(
                {
                    "Pickup_A": "{:.1f}",
                    "TMS": "{:.3f}",
                    "Inst_A": "{:.1f}",
                    "Operate_s": "{:.4f}",
                }
            ),
            use_container_width=True,
        )

        grading_rows = []
        for index in range(len(relay_results_df) - 1):
            downstream = relay_results_df.iloc[index]
            upstream = relay_results_df.iloc[index + 1]
            downstream_time_s = float(downstream["Operate_s"])
            upstream_time_s = float(upstream["Operate_s"])

            if not (math.isfinite(downstream_time_s) and math.isfinite(upstream_time_s)):
                margin_s = math.nan
                status = "CHECK"
            else:
                margin_s = upstream_time_s - downstream_time_s
                status = "PASS" if margin_s >= grading_margin_s else "FAIL"

            grading_rows.append(
                {
                    "Downstream": downstream["Device"],
                    "Upstream": upstream["Device"],
                    "Margin_s": margin_s,
                    "Target_s": grading_margin_s,
                    "Status": status,
                }
            )

        grading_df = pd.DataFrame(grading_rows)
        st.markdown("#### Margin Readout at Slider Current")
        st.dataframe(
            grading_df.style.format({"Margin_s": "{:.3f}", "Target_s": "{:.3f}"}),
            use_container_width=True,
        )

        margin_readout = []
        for _, margin_row in grading_df.iterrows():
            margin_value = float(margin_row["Margin_s"])
            if math.isfinite(margin_value) and margin_row["Status"] in ["PASS", "FAIL"]:
                margin_readout.append(
                    f"{margin_row['Downstream']} → {margin_row['Upstream']}: {margin_value:.3f} s ({margin_row['Status']})"
                )
            else:
                margin_readout.append(
                    f"{margin_row['Downstream']} → {margin_row['Upstream']}: CHECK"
                )
        if margin_readout:
            st.caption(" | ".join(margin_readout))

        graded_df = grading_df[grading_df["Status"].isin(["PASS", "FAIL"])]
        if not graded_df.empty and (graded_df["Status"] == "FAIL").any():
            st.error("One or more relay pairs do not meet the grading margin target.")
        elif not graded_df.empty:
            st.success("All checked relay pairs meet the grading margin target.")
        else:
            st.info("No relay pairs qualified for grading checks at this selected current.")

        current_points = [
            min_pickup * ((max_current / min_pickup) ** (idx / 79))
            for idx in range(80)
        ]
        tcc_df = pd.DataFrame({"Current_A": current_points})

        for _, row in relay_df.sort_values("Order").iterrows():
            times = [
                calc_relay_time_s(
                    current_a=value,
                    pickup_a=float(row["Pickup_A"]),
                    tms=float(row["TMS"]),
                    curve_name=str(row["Curve"]),
                    inst_pickup_a=0.0,
                )
                for value in current_points
            ]
            cleaned_times = [time if math.isfinite(time) and time <= 30 else None for time in times]
            tcc_df[str(row["Device"])] = cleaned_times

        st.markdown("#### Time-Current Curves (Approximate)")
        tcc_long_df = tcc_df.melt(
            id_vars="Current_A",
            var_name="Device",
            value_name="Operate_s",
        ).dropna()
        tcc_long_df = tcc_long_df[(tcc_long_df["Current_A"] > 0) & (tcc_long_df["Operate_s"] > 0)]

        if tcc_long_df.empty:
            st.warning("No valid relay curve points available for log-log plotting.")
        else:
            x_scale = alt.Scale(type="log", domain=[min_pickup * 0.9, max_current])

            idmt_layer = (
                alt.Chart(tcc_long_df)
                .mark_line()
                .encode(
                    x=alt.X(
                        "Current_A:Q",
                        title="Current (A)",
                        scale=x_scale,
                        axis=alt.Axis(grid=True),
                    ),
                    y=alt.Y(
                        "Operate_s:Q",
                        title="Operating Time (s)",
                        scale=alt.Scale(type="log"),
                        axis=alt.Axis(grid=True),
                    ),
                    color=alt.Color("Device:N", title="Device"),
                    tooltip=[
                        alt.Tooltip("Device:N", title="Device"),
                        alt.Tooltip("Current_A:Q", title="Current (A)", format=".2f"),
                        alt.Tooltip("Operate_s:Q", title="Time (s)", format=".4f"),
                    ],
                )
            )

            inst_df = (
                relay_df[relay_df["Inst_A"] > 0][["Device", "Inst_A"]]
                .rename(columns={"Inst_A": "Current_A"})
                .copy()
            )

            grading_current_df = pd.DataFrame(
                {"Current_A": [grading_current_a], "Marker": ["Selected grading current"]}
            )

            layers = [idmt_layer]

            if not inst_df.empty:
                inst_layer = (
                    alt.Chart(inst_df)
                    .mark_rule(strokeDash=[4, 4], strokeWidth=1.5)
                    .encode(
                        x=alt.X("Current_A:Q", scale=x_scale),
                        color=alt.Color("Device:N", title="Device"),
                        tooltip=[
                            alt.Tooltip("Device:N", title="Device"),
                            alt.Tooltip("Current_A:Q", title="Inst. Pickup (A)", format=".0f"),
                        ],
                    )
                )
                layers.append(inst_layer)

            grading_current_layer = (
                alt.Chart(grading_current_df)
                .mark_rule(strokeWidth=2)
                .encode(
                    x=alt.X("Current_A:Q", scale=x_scale),
                    tooltip=[
                        alt.Tooltip("Marker:N", title="Marker"),
                        alt.Tooltip("Current_A:Q", title="Current (A)", format=".0f"),
                    ],
                )
            )
            layers.append(grading_current_layer)

            tcc_chart = (
                alt.layer(*layers)
                .properties(height=420)
                .configure_axis(grid=True)
            )

            st.altair_chart(tcc_chart, use_container_width=True)

with arc_tab:
    st.subheader("Arc-Flash Screening Estimate")
    st.caption(
        "This is a simplified 11kV switchgear screening model for early design checks. "
        "Select electrode orientation and whether the switchgear is enclosed or unenclosed. "
        "Use a full IEEE 1584 / NFPA 70E study for engineering sign-off."
    )

    relay_df_for_arc = sanitize_relay_settings(pd.DataFrame(st.session_state.relay_settings))
    suggested_clearing_time = 0.20
    if not relay_df_for_arc.empty:
        trip_times = [
            calc_relay_time_s(
                current_a=selected_fault_current_a,
                pickup_a=float(row["Pickup_A"]),
                tms=float(row["TMS"]),
                curve_name=str(row["Curve"]),
                inst_pickup_a=float(row["Inst_A"]),
            )
            for _, row in relay_df_for_arc.iterrows()
        ]
        finite_times = [time for time in trip_times if math.isfinite(time)]
        if finite_times:
            suggested_clearing_time = max(0.05, min(finite_times))

    st.caption(f"Suggested clearing time from fastest relay at selected fault: {suggested_clearing_time:.3f} s")

    arc_col1, arc_col2, arc_col3 = st.columns(3)
    with arc_col1:
        bolted_fault_current_ka = st.number_input(
            "Bolted fault current (kA)",
            min_value=0.1,
            max_value=200.0,
            step=0.1,
            key="bolted_fault_current_ka",
        )
        arc_current_factor = st.slider(
            "Arc current factor",
            0.50,
            1.00,
            step=0.01,
            key="arc_current_factor",
        )

    with arc_col2:
        clearing_time_s = st.number_input(
            "Clearing time (s)",
            min_value=0.01,
            max_value=5.00,
            step=0.01,
            key="clearing_time_s",
        )
        working_distance_mm = st.number_input(
            "Working distance (mm)",
            min_value=200,
            max_value=2000,
            step=25,
            key="working_distance_mm",
        )

    with arc_col3:
        arc_orientation = st.selectbox(
            "Electrode orientation",
            options=["Horizontal", "Vertical"],
            key="arc_orientation",
        )
        arc_enclosure = st.selectbox(
            "Switchgear construction",
            options=["Enclosed", "Unenclosed"],
            key="arc_enclosure",
        )
        configuration_factor = ARC_FLASH_SWITCHGEAR_FACTOR_MAP[(arc_orientation, arc_enclosure)]

    st.caption(
        f"11kV switchgear configuration: {arc_orientation.lower()} electrodes, {arc_enclosure.lower()} switchgear "
        f"(screening factor {configuration_factor:.2f})."
    )

    arc_current_ka = bolted_fault_current_ka * arc_current_factor
    distance_multiplier = (610 / max(working_distance_mm, 1)) ** 1.5
    incident_energy_cal_cm2 = (
        0.2
        * arc_current_ka
        * clearing_time_s
        * configuration_factor
        * distance_multiplier
    )
    arc_flash_boundary_mm = 610 * ((max(incident_energy_cal_cm2, 0.001) / 1.2) ** (1 / 1.5))

    result_col1, result_col2, result_col3 = st.columns(3)
    with result_col1:
        st.metric("Arcing current", f"{arc_current_ka:.2f} kA")
    with result_col2:
        st.metric("Incident energy", f"{incident_energy_cal_cm2:.2f} cal/cm²")
    with result_col3:
        st.metric("Arc flash boundary", f"{arc_flash_boundary_mm:.0f} mm")

    st.info(arc_flash_category(incident_energy_cal_cm2))

with full_arc_tab:
    st.subheader("Full Arc-Flash Calculation (Detailed Screening)")
    st.caption(
        "Expanded arc-flash workflow using min/nominal/max arcing-current scenarios and relay-based clearing times. "
        "This remains a detailed screening model and is not a substitute for a full IEEE 1584 study."
    )

    full_col1, full_col2, full_col3 = st.columns(3)
    with full_col1:
        bolted_fault_current_ka_full = st.number_input(
            "Bolted fault current (kA) - full calc",
            min_value=0.1,
            max_value=200.0,
            value=float(st.session_state.get("bolted_fault_current_ka", 10.0)),
            step=0.1,
            key="full_arc_bolted_fault_current_ka",
        )
        arc_current_factor_full = st.slider(
            "Arc current factor - full calc",
            0.50,
            1.00,
            value=float(st.session_state.get("arc_current_factor", 0.85)),
            step=0.01,
            key="full_arc_arc_current_factor",
        )
        arc_current_variation_pct = st.slider(
            "Arcing current variation (%)",
            min_value=0,
            max_value=30,
            value=15,
            step=1,
            key="full_arc_current_variation_pct",
        )

    with full_col2:
        clearing_time_method = st.selectbox(
            "Clearing time method",
            options=["Manual fixed", "Relay-derived (capped)"],
            key="full_arc_clearing_time_method",
        )
        if clearing_time_method == "Manual fixed":
            manual_clearing_time_s = st.number_input(
                "Applied clearing time (s)",
                min_value=0.01,
                max_value=5.00,
                value=float(st.session_state.get("clearing_time_s", 0.20)),
                step=0.01,
                key="full_arc_manual_clearing_time_s",
            )
            clearing_time_cap_s = float(manual_clearing_time_s)
        else:
            manual_clearing_time_s = math.nan
            clearing_time_cap_s = st.number_input(
                "Clearing time cap (s)",
                min_value=0.01,
                max_value=5.00,
                value=float(st.session_state.get("clearing_time_s", 0.20)),
                step=0.01,
                key="full_arc_clearing_time_cap_s",
            )
        working_distance_mm_full = st.number_input(
            "Working distance (mm) - full calc",
            min_value=200,
            max_value=2000,
            value=int(st.session_state.get("working_distance_mm", 910)),
            step=25,
            key="full_arc_working_distance_mm",
        )

    with full_col3:
        default_orientation = str(st.session_state.get("arc_orientation", "Vertical"))
        orientation_options = ["Horizontal", "Vertical"]
        full_arc_orientation = st.selectbox(
            "Electrode orientation - full calc",
            options=orientation_options,
            index=orientation_options.index(default_orientation) if default_orientation in orientation_options else 1,
            key="full_arc_orientation",
        )
        default_enclosure = str(st.session_state.get("arc_enclosure", "Enclosed"))
        enclosure_options = ["Enclosed", "Unenclosed"]
        full_arc_enclosure = st.selectbox(
            "Switchgear construction - full calc",
            options=enclosure_options,
            index=enclosure_options.index(default_enclosure) if default_enclosure in enclosure_options else 0,
            key="full_arc_enclosure",
        )

    full_config_factor = ARC_FLASH_SWITCHGEAR_FACTOR_MAP[(full_arc_orientation, full_arc_enclosure)]
    variation_factor = float(arc_current_variation_pct) / 100.0
    arc_current_nominal_ka = bolted_fault_current_ka_full * arc_current_factor_full
    arc_current_min_ka = arc_current_nominal_ka * (1.0 - variation_factor)
    arc_current_max_ka = arc_current_nominal_ka * (1.0 + variation_factor)
    distance_multiplier_full = (610 / max(working_distance_mm_full, 1)) ** 1.5

    relay_df_for_full_arc = sanitize_relay_settings(pd.DataFrame(st.session_state.relay_settings))

    scenario_rows = []
    for scenario_name, arc_current_ka_scenario in [
        ("Min arcing current", arc_current_min_ka),
        ("Nominal arcing current", arc_current_nominal_ka),
        ("Max arcing current", arc_current_max_ka),
    ]:
        if clearing_time_method == "Manual fixed":
            applied_clearing_time_s = float(manual_clearing_time_s)
            clearing_source = "Manual"
        else:
            relay_clearing_time_s = math.inf
            if not relay_df_for_full_arc.empty:
                scenario_trip_times = [
                    calc_relay_time_s(
                        current_a=arc_current_ka_scenario * 1000,
                        pickup_a=float(row["Pickup_A"]),
                        tms=float(row["TMS"]),
                        curve_name=str(row["Curve"]),
                        inst_pickup_a=float(row["Inst_A"]),
                    )
                    for _, row in relay_df_for_full_arc.iterrows()
                ]
                finite_trip_times = [time for time in scenario_trip_times if math.isfinite(time)]
                if finite_trip_times:
                    relay_clearing_time_s = max(0.05, min(finite_trip_times))

            if math.isfinite(relay_clearing_time_s):
                applied_clearing_time_s = min(relay_clearing_time_s, clearing_time_cap_s)
                clearing_source = "Relay (capped)" if relay_clearing_time_s > clearing_time_cap_s else "Relay"
            else:
                applied_clearing_time_s = clearing_time_cap_s
                clearing_source = "Cap (no relay trip)"

        incident_energy_scenario = (
            0.2
            * arc_current_ka_scenario
            * applied_clearing_time_s
            * full_config_factor
            * distance_multiplier_full
        )
        arc_flash_boundary_scenario_mm = 610 * (
            (max(incident_energy_scenario, 0.001) / 1.2) ** (1 / 1.5)
        )

        scenario_rows.append(
            {
                "Scenario": scenario_name,
                "Arc_Current_kA": arc_current_ka_scenario,
                "Clearing_s": applied_clearing_time_s,
                "Clearing_Source": clearing_source,
                "Incident_Energy_cal_cm2": incident_energy_scenario,
                "Arc_Flash_Boundary_mm": arc_flash_boundary_scenario_mm,
            }
        )

    full_arc_df = pd.DataFrame(scenario_rows)
    worst_row = full_arc_df.loc[full_arc_df["Incident_Energy_cal_cm2"].idxmax()]

    st.caption(
        f"Configuration: {full_arc_orientation.lower()} electrodes, {full_arc_enclosure.lower()} switchgear "
        f"(factor {full_config_factor:.2f}); working distance {working_distance_mm_full} mm."
    )
    if clearing_time_method == "Manual fixed":
        st.caption(f"Clearing time mode: manual fixed at {manual_clearing_time_s:.3f} s.")
    else:
        st.caption(f"Clearing time mode: relay-derived with cap {clearing_time_cap_s:.3f} s.")

    full_result_col1, full_result_col2, full_result_col3, full_result_col4 = st.columns(4)
    with full_result_col1:
        st.metric("Nominal arcing current", f"{arc_current_nominal_ka:.2f} kA")
    with full_result_col2:
        st.metric("Worst incident energy", f"{float(worst_row['Incident_Energy_cal_cm2']):.2f} cal/cm²")
    with full_result_col3:
        st.metric("Worst arc flash boundary", f"{float(worst_row['Arc_Flash_Boundary_mm']):.0f} mm")
    with full_result_col4:
        st.metric("Controlling scenario", str(worst_row["Scenario"]))

    st.dataframe(
        full_arc_df.style.format(
            {
                "Arc_Current_kA": "{:.2f}",
                "Clearing_s": "{:.3f}",
                "Incident_Energy_cal_cm2": "{:.2f}",
                "Arc_Flash_Boundary_mm": "{:.0f}",
            }
        ),
        use_container_width=True,
    )

    st.info(arc_flash_category(float(worst_row["Incident_Energy_cal_cm2"])))

with oc_ef_commissioning_tab:
    st.subheader("OC / EF Commissioning Worksheet")
    st.caption(
        "Record relay settings, calculate expected operating times, and generate three suggested secondary injection points for commissioning."
    )

    oc_settings_col1, oc_settings_col2, oc_settings_col3 = st.columns(3)
    with oc_settings_col1:
        oc_relay_model = st.selectbox(
            "Relay model",
            options=OC_EF_RELAY_MODELS,
            key="oc_ef_relay_model",
        )
        oc_element_type = st.selectbox(
            "Element type",
            options=["OC", "EF"],
            key="oc_ef_element_type",
        )
        oc_device_name = st.text_input(
            "Relay / feeder reference",
            key="oc_ef_device_name",
        )
        oc_curve_name = st.selectbox(
            "Characteristic",
            options=COMMISSIONING_CURVE_OPTIONS,
            key="oc_ef_curve_name",
        )

    with oc_settings_col2:
        oc_ct_primary_a = st.number_input(
            "CT primary (A)",
            min_value=1.0,
            step=1.0,
            key="oc_ef_ct_primary_a",
        )
        oc_ct_secondary_a = st.selectbox(
            "CT secondary (A)",
            options=[1.0, 5.0],
            key="oc_ef_ct_secondary_a",
        )
        oc_pickup_secondary_a = st.number_input(
            "Pickup setting (A secondary)",
            min_value=0.05,
            step=0.05,
            key="oc_ef_pickup_secondary_a",
        )
        oc_tms = st.number_input(
            "TMS / time dial",
            min_value=0.01,
            max_value=2.00,
            step=0.01,
            key="oc_ef_tms",
        )

    with oc_settings_col3:
        oc_definite_time_s = st.number_input(
            "Definite time setting (s)",
            min_value=0.00,
            max_value=30.00,
            step=0.01,
            key="oc_ef_definite_time_s",
        )
        oc_inst_pickup_a = st.number_input(
            "High-set / instantaneous (A secondary)",
            min_value=0.00,
            step=0.10,
            key="oc_ef_inst_pickup_a",
        )
        oc_notes = st.text_area(
            "Commissioning notes",
            key="oc_ef_notes",
            height=96,
        )

    st.info(OC_EF_MODEL_NOTES[oc_relay_model])

    oc_pickup_primary_a = oc_pickup_secondary_a * (oc_ct_primary_a / oc_ct_secondary_a)
    oc_inst_primary_a = (
        oc_inst_pickup_a * (oc_ct_primary_a / oc_ct_secondary_a) if oc_inst_pickup_a > 0 else math.nan
    )
    oc_settings_df = pd.DataFrame(
        [
            {
                "Relay_Model": oc_relay_model,
                "Element_Type": oc_element_type,
                "Device": oc_device_name,
                "Characteristic": oc_curve_name,
                "CT_Primary_A": oc_ct_primary_a,
                "CT_Secondary_A": oc_ct_secondary_a,
                "Pickup_Secondary_A": oc_pickup_secondary_a,
                "Pickup_Primary_A": oc_pickup_primary_a,
                "TMS": oc_tms,
                "Definite_Time_s": oc_definite_time_s,
                "High_Set_Secondary_A": oc_inst_pickup_a,
                "High_Set_Primary_A": oc_inst_primary_a,
                "Notes": oc_notes,
            }
        ]
    )

    oc_test_points_df = build_oc_ef_test_points(
        pickup_secondary_a=oc_pickup_secondary_a,
        curve_name=oc_curve_name,
        tms=oc_tms,
        definite_time_s=oc_definite_time_s,
        inst_pickup_a=oc_inst_pickup_a,
        ct_primary_a=oc_ct_primary_a,
        ct_secondary_a=oc_ct_secondary_a,
    )
    oc_test_points_df["Expected_Operate_Display"] = oc_test_points_df["Expected_Operate_s"].apply(format_time_value)

    oc_metric_col1, oc_metric_col2, oc_metric_col3 = st.columns(3)
    with oc_metric_col1:
        st.metric("Pickup primary", f"{oc_pickup_primary_a:.1f} A")
    with oc_metric_col2:
        high_set_text = f"{oc_inst_primary_a:.1f} A" if math.isfinite(oc_inst_primary_a) else "Not enabled"
        st.metric("High-set primary", high_set_text)
    with oc_metric_col3:
        st.metric("Suggested injections", f"{len(oc_test_points_df)} points")

    oc_curve_max_secondary_a = max(
        float(oc_test_points_df["Injection_Current_Secondary_A"].max()) * 1.3,
        oc_pickup_secondary_a * 8.0,
        (oc_inst_pickup_a * 1.2) if oc_inst_pickup_a > 0 else 0.0,
    )
    oc_curve_points = [
        oc_pickup_secondary_a * (oc_curve_max_secondary_a / oc_pickup_secondary_a) ** (index / 79)
        for index in range(80)
    ]
    oc_curve_df = pd.DataFrame(
        {
            "Current_Secondary_A": oc_curve_points,
            "Expected_Operate_s": [
                calc_commissioning_time_s(
                    current_a=current_value,
                    pickup_a=oc_pickup_secondary_a,
                    curve_name=oc_curve_name,
                    tms=oc_tms,
                    definite_time_s=oc_definite_time_s,
                    inst_pickup_a=oc_inst_pickup_a,
                )
                for current_value in oc_curve_points
            ],
        }
    )
    oc_curve_df = oc_curve_df.replace([math.inf, -math.inf], math.nan).dropna()

    st.markdown("#### Recorded Settings")
    st.dataframe(
        oc_settings_df.style.format(
            {
                "CT_Primary_A": "{:.1f}",
                "CT_Secondary_A": "{:.1f}",
                "Pickup_Secondary_A": "{:.2f}",
                "Pickup_Primary_A": "{:.1f}",
                "TMS": "{:.3f}",
                "Definite_Time_s": "{:.3f}",
                "High_Set_Secondary_A": "{:.2f}",
                "High_Set_Primary_A": "{:.1f}",
            }
        ),
        use_container_width=True,
    )

    oc_download_col1, oc_download_col2 = st.columns(2)
    with oc_download_col1:
        st.download_button(
            "Export OC/EF settings CSV",
            data=oc_settings_df.to_csv(index=False).encode("utf-8"),
            file_name="oc_ef_commissioning_settings.csv",
            mime="text/csv",
            use_container_width=True,
        )
    with oc_download_col2:
        st.download_button(
            "Export OC/EF test points CSV",
            data=oc_test_points_df.to_csv(index=False).encode("utf-8"),
            file_name="oc_ef_commissioning_test_points.csv",
            mime="text/csv",
            use_container_width=True,
        )

    st.markdown("#### Suggested Three-Point Injection Plan")
    st.dataframe(
        oc_test_points_df.style.format(
            {
                "Injection_Current_Secondary_A": "{:.2f}",
                "Injection_Current_Primary_A": "{:.1f}",
                "Expected_Operate_s": "{:.3f}",
            }
        ),
        use_container_width=True,
    )
    if oc_inst_pickup_a > 0:
        st.caption(
            f"High-set is enabled at {oc_inst_pickup_a:.2f} A secondary. Add a separate pickup check around 0.95x / 1.05x high-set if site scope includes the instantaneous element."
        )

    st.markdown("#### Expected Operating Curve")
    if oc_curve_df.empty:
        st.warning("No valid OC/EF curve points are available for the selected settings.")
    else:
        oc_curve_chart = alt.Chart(oc_curve_df).mark_line().encode(
            x=alt.X(
                "Current_Secondary_A:Q",
                title="Secondary injection current (A)",
                scale=alt.Scale(type="log"),
                axis=alt.Axis(grid=True),
            ),
            y=alt.Y(
                "Expected_Operate_s:Q",
                title="Expected operating time (s)",
                scale=alt.Scale(type="log"),
                axis=alt.Axis(grid=True),
            ),
            tooltip=[
                alt.Tooltip("Current_Secondary_A:Q", title="Secondary current (A)", format=".2f"),
                alt.Tooltip("Expected_Operate_s:Q", title="Expected time (s)", format=".3f"),
            ],
        )
        oc_point_chart = alt.Chart(oc_test_points_df).mark_point(size=90, filled=True).encode(
            x=alt.X("Injection_Current_Secondary_A:Q", scale=alt.Scale(type="log")),
            y=alt.Y("Expected_Operate_s:Q", scale=alt.Scale(type="log")),
            color=alt.Color("Point:N", legend=None),
            tooltip=[
                alt.Tooltip("Point:N"),
                alt.Tooltip("Purpose:N"),
                alt.Tooltip("Injection_Current_Secondary_A:Q", title="Secondary current (A)", format=".2f"),
                alt.Tooltip("Expected_Operate_s:Q", title="Expected time (s)", format=".3f"),
                alt.Tooltip("Expected_Mode:N", title="Mode"),
            ],
        )
        st.altair_chart(
            alt.layer(oc_curve_chart, oc_point_chart).properties(height=420),
            use_container_width=True,
        )

with sel787_tab:
    st.subheader("SEL-787 Differential Commissioning Worksheet")
    st.caption(
        "Record differential settings, plot the expected restrained characteristic, and generate suggested secondary injection points."
    )
    st.info(SEL787_DEFAULT_NOTES)

    sel_settings_col1, sel_settings_col2, sel_settings_col3 = st.columns(3)
    with sel_settings_col1:
        sel_asset_name = st.text_input(
            "Transformer / relay reference",
            key="sel787_asset_name",
        )
        sel_w1_ct_primary_a = st.number_input(
            "Winding 1 CT primary (A)",
            min_value=1.0,
            step=1.0,
            key="sel787_w1_ct_primary_a",
        )
        sel_w1_ct_secondary_a = st.selectbox(
            "Winding 1 CT secondary (A)",
            options=[1.0, 5.0],
            key="sel787_w1_ct_secondary_a",
        )
        sel_w2_ct_primary_a = st.number_input(
            "Winding 2 CT primary (A)",
            min_value=1.0,
            step=1.0,
            key="sel787_w2_ct_primary_a",
        )

    with sel_settings_col2:
        sel_w2_ct_secondary_a = st.selectbox(
            "Winding 2 CT secondary (A)",
            options=[1.0, 5.0],
            key="sel787_w2_ct_secondary_a",
        )
        sel_pickup_a = st.number_input(
            "Differential pickup (A secondary)",
            min_value=0.05,
            step=0.05,
            key="sel787_pickup_a",
        )
        sel_slope1_pct = st.number_input(
            "Slope 1 (%)",
            min_value=0.0,
            max_value=100.0,
            step=1.0,
            key="sel787_slope1_pct",
        )
        sel_slope2_pct = st.number_input(
            "Slope 2 (%)",
            min_value=0.0,
            max_value=100.0,
            step=1.0,
            key="sel787_slope2_pct",
        )

    with sel_settings_col3:
        sel_breakpoint_a = st.number_input(
            "Slope breakpoint (A restraint)",
            min_value=0.0,
            step=0.10,
            key="sel787_breakpoint_a",
        )
        sel_unrestrained_a = st.number_input(
            "Unrestrained 87U pickup (A secondary)",
            min_value=0.0,
            step=0.10,
            key="sel787_unrestrained_a",
        )
        sel_h2_block_pct = st.number_input(
            "2nd harmonic block (%)",
            min_value=0.0,
            max_value=100.0,
            step=1.0,
            key="sel787_h2_block_pct",
        )
        sel_h5_block_pct = st.number_input(
            "5th harmonic block (%)",
            min_value=0.0,
            max_value=100.0,
            step=1.0,
            key="sel787_h5_block_pct",
        )

    sel_settings_df = pd.DataFrame(
        [
            {
                "Relay_Model": "SEL-787",
                "Asset": sel_asset_name,
                "W1_CT_Primary_A": sel_w1_ct_primary_a,
                "W1_CT_Secondary_A": sel_w1_ct_secondary_a,
                "W2_CT_Primary_A": sel_w2_ct_primary_a,
                "W2_CT_Secondary_A": sel_w2_ct_secondary_a,
                "Diff_Pickup_A": sel_pickup_a,
                "Slope1_pct": sel_slope1_pct,
                "Slope2_pct": sel_slope2_pct,
                "Breakpoint_A": sel_breakpoint_a,
                "Unrestrained_87U_A": sel_unrestrained_a,
                "2nd_Harmonic_Block_pct": sel_h2_block_pct,
                "5th_Harmonic_Block_pct": sel_h5_block_pct,
            }
        ]
    )

    sel_test_points_df = build_sel787_test_points(
        pickup_a=sel_pickup_a,
        slope1_pct=sel_slope1_pct,
        slope2_pct=sel_slope2_pct,
        breakpoint_a=sel_breakpoint_a,
        unrestrained_a=sel_unrestrained_a,
    )

    sel_metric_col1, sel_metric_col2, sel_metric_col3 = st.columns(3)
    with sel_metric_col1:
        st.metric("Differential pickup", f"{sel_pickup_a:.2f} A")
    with sel_metric_col2:
        st.metric("Breakpoint restraint", f"{sel_breakpoint_a:.2f} A")
    with sel_metric_col3:
        unrestrained_text = f"{sel_unrestrained_a:.2f} A" if sel_unrestrained_a > 0 else "Disabled"
        st.metric("87U pickup", unrestrained_text)

    st.markdown("#### Recorded Settings")
    st.dataframe(
        sel_settings_df.style.format(
            {
                "W1_CT_Primary_A": "{:.1f}",
                "W1_CT_Secondary_A": "{:.1f}",
                "W2_CT_Primary_A": "{:.1f}",
                "W2_CT_Secondary_A": "{:.1f}",
                "Diff_Pickup_A": "{:.2f}",
                "Slope1_pct": "{:.1f}",
                "Slope2_pct": "{:.1f}",
                "Breakpoint_A": "{:.2f}",
                "Unrestrained_87U_A": "{:.2f}",
                "2nd_Harmonic_Block_pct": "{:.1f}",
                "5th_Harmonic_Block_pct": "{:.1f}",
            }
        ),
        use_container_width=True,
    )

    sel_download_col1, sel_download_col2 = st.columns(2)
    with sel_download_col1:
        st.download_button(
            "Export SEL-787 settings CSV",
            data=sel_settings_df.to_csv(index=False).encode("utf-8"),
            file_name="sel_787_settings.csv",
            mime="text/csv",
            use_container_width=True,
        )
    with sel_download_col2:
        st.download_button(
            "Export SEL-787 test points CSV",
            data=sel_test_points_df.to_csv(index=False).encode("utf-8"),
            file_name="sel_787_test_points.csv",
            mime="text/csv",
            use_container_width=True,
        )

    st.markdown("#### Suggested Three-Point Injection Plan")
    st.dataframe(
        sel_test_points_df.style.format(
            {
                "Irest_A": "{:.2f}",
                "Iop_A": "{:.2f}",
                "Side1_Secondary_A": "{:.2f}",
                "Side2_Secondary_A": "{:.2f}",
                "Characteristic_Threshold_A": "{:.2f}",
            }
        ),
        use_container_width=True,
    )

    sel_restraint_limit_a = max(sel_breakpoint_a * 2.5, sel_pickup_a * 10.0, 10.0)
    sel_characteristic_df = pd.DataFrame(
        {
            "Irest_A": [sel_restraint_limit_a * index / 100 for index in range(101)],
        }
    )
    sel_characteristic_df["Trip_Threshold_A"] = sel_characteristic_df["Irest_A"].apply(
        lambda restraint_value: sel787_operate_threshold(
            restraint_value,
            sel_pickup_a,
            sel_slope1_pct,
            sel_slope2_pct,
            sel_breakpoint_a,
        )
    )

    st.markdown("#### Expected Differential Characteristic")
    sel_chart_layers = [
        alt.Chart(sel_characteristic_df)
        .mark_line()
        .encode(
            x=alt.X("Irest_A:Q", title="Restraint current (A secondary)", axis=alt.Axis(grid=True)),
            y=alt.Y("Trip_Threshold_A:Q", title="Operate current (A secondary)", axis=alt.Axis(grid=True)),
            tooltip=[
                alt.Tooltip("Irest_A:Q", title="Irest (A)", format=".2f"),
                alt.Tooltip("Trip_Threshold_A:Q", title="Trip threshold (A)", format=".2f"),
            ],
        )
    ]

    if sel_unrestrained_a > 0:
        sel_unrestrained_df = pd.DataFrame(
            {
                "Irest_A": [0.0, sel_restraint_limit_a],
                "Trip_Threshold_A": [sel_unrestrained_a, sel_unrestrained_a],
                "Legend": ["87U", "87U"],
            }
        )
        sel_chart_layers.append(
            alt.Chart(sel_unrestrained_df)
            .mark_rule(strokeDash=[4, 4], color="#D62728")
            .encode(
                y=alt.Y("Trip_Threshold_A:Q"),
                tooltip=[alt.Tooltip("Trip_Threshold_A:Q", title="87U pickup (A)", format=".2f")],
            )
        )

    sel_chart_layers.append(
        alt.Chart(sel_test_points_df)
        .mark_point(size=90, filled=True)
        .encode(
            x=alt.X("Irest_A:Q"),
            y=alt.Y("Iop_A:Q"),
            color=alt.Color("Expected_Result:N", title="Expected"),
            tooltip=[
                alt.Tooltip("Test:N"),
                alt.Tooltip("Purpose:N"),
                alt.Tooltip("Irest_A:Q", title="Irest (A)", format=".2f"),
                alt.Tooltip("Iop_A:Q", title="Iop (A)", format=".2f"),
                alt.Tooltip("Expected_Result:N", title="Expected"),
            ],
        )
    )

    st.altair_chart(alt.layer(*sel_chart_layers).properties(height=420), use_container_width=True)

with translay_tab:
    st.subheader("Translay Commissioning Worksheet")
    st.caption(
        "Record translay settings, calculate expected restrained-spill characteristic, and generate three suggested injection points."
    )

    translay_relay_model_col1, translay_relay_model_col2 = st.columns([4, 1])
    with translay_relay_model_col1:
        translay_relay_model = st.selectbox(
            "Translay relay model",
            options=TRANSLAY_RELAY_OPTIONS,
            key="translay_relay_model",
        )
    with translay_relay_model_col2:
        st.write("")
        st.write("")
        if st.button("Apply model defaults", key="translay_apply_defaults", use_container_width=True):
            selected_defaults = TRANSLAY_MODEL_DEFAULTS[translay_relay_model]
            st.session_state["translay_pickup_a"] = selected_defaults["pickup_a"]
            st.session_state["translay_slope_pct"] = selected_defaults["slope_pct"]
            st.session_state["translay_operate_time_s"] = selected_defaults["operate_time_s"]
            st.session_state["translay_high_set_spill_a"] = selected_defaults["high_set_spill_a"]
            st.rerun()

    translay_col1, translay_col2, translay_col3 = st.columns(3)
    with translay_col1:
        translay_circuit_name = st.text_input(
            "Protected circuit reference",
            key="translay_circuit_name",
        )
        translay_local_ct_primary_a = st.number_input(
            "Local CT primary (A)",
            min_value=1.0,
            step=1.0,
            key="translay_local_ct_primary_a",
        )
        translay_local_ct_secondary_a = st.selectbox(
            "Local CT secondary (A)",
            options=[1.0, 5.0],
            key="translay_local_ct_secondary_a",
        )

    with translay_col2:
        translay_remote_ct_primary_a = st.number_input(
            "Remote CT primary (A)",
            min_value=1.0,
            step=1.0,
            key="translay_remote_ct_primary_a",
        )
        translay_remote_ct_secondary_a = st.selectbox(
            "Remote CT secondary (A)",
            options=[1.0, 5.0],
            key="translay_remote_ct_secondary_a",
        )
        translay_pickup_a = st.number_input(
            "Spill pickup (A secondary)",
            min_value=0.02,
            step=0.01,
            key="translay_pickup_a",
        )

    with translay_col3:
        translay_slope_pct = st.number_input(
            "Restraint slope (%)",
            min_value=0.0,
            max_value=100.0,
            step=1.0,
            key="translay_slope_pct",
        )
        translay_operate_time_s = st.number_input(
            "Expected operate time (s)",
            min_value=0.01,
            max_value=10.0,
            step=0.01,
            key="translay_operate_time_s",
        )
        translay_high_set_spill_a = st.number_input(
            "High-set spill element (A secondary)",
            min_value=0.0,
            step=0.01,
            key="translay_high_set_spill_a",
        )
        translay_notes = st.text_area(
            "Commissioning notes",
            key="translay_notes",
            height=90,
        )

    st.info(TRANSLAY_MODEL_NOTES[translay_relay_model])

    translay_local_ratio = translay_local_ct_primary_a / translay_local_ct_secondary_a
    translay_remote_ratio = translay_remote_ct_primary_a / translay_remote_ct_secondary_a

    translay_settings_df = pd.DataFrame(
        [
            {
                "Relay_Model": translay_relay_model,
                "Circuit": translay_circuit_name,
                "Local_CT_Primary_A": translay_local_ct_primary_a,
                "Local_CT_Secondary_A": translay_local_ct_secondary_a,
                "Remote_CT_Primary_A": translay_remote_ct_primary_a,
                "Remote_CT_Secondary_A": translay_remote_ct_secondary_a,
                "Spill_Pickup_A": translay_pickup_a,
                "Slope_pct": translay_slope_pct,
                "Expected_Operate_s": translay_operate_time_s,
                "High_Set_Spill_A": translay_high_set_spill_a,
                "Notes": translay_notes,
            }
        ]
    )

    translay_test_points_df = build_translay_test_points(
        pickup_a=translay_pickup_a,
        slope_pct=translay_slope_pct,
        operate_time_s=translay_operate_time_s,
        high_set_spill_a=translay_high_set_spill_a,
    )
    translay_test_points_df["Local_Current_Primary_A"] = (
        translay_test_points_df["Local_Current_A"] * translay_local_ratio
    )
    translay_test_points_df["Remote_Current_Primary_A"] = (
        translay_test_points_df["Remote_Current_A"] * translay_remote_ratio
    )
    translay_test_points_df["Expected_Operate_Display"] = translay_test_points_df[
        "Expected_Operate_s"
    ].apply(format_time_value)

    translay_metric_col1, translay_metric_col2, translay_metric_col3 = st.columns(3)
    with translay_metric_col1:
        st.metric("Spill pickup", f"{translay_pickup_a:.2f} A")
    with translay_metric_col2:
        st.metric("Restraint slope", f"{translay_slope_pct:.1f} %")
    with translay_metric_col3:
        st.metric("Suggested injections", f"{len(translay_test_points_df)} points")

    st.markdown("#### Recorded Settings")
    st.dataframe(
        translay_settings_df.style.format(
            {
                "Local_CT_Primary_A": "{:.1f}",
                "Local_CT_Secondary_A": "{:.1f}",
                "Remote_CT_Primary_A": "{:.1f}",
                "Remote_CT_Secondary_A": "{:.1f}",
                "Spill_Pickup_A": "{:.2f}",
                "Slope_pct": "{:.1f}",
                "Expected_Operate_s": "{:.3f}",
                "High_Set_Spill_A": "{:.2f}",
            }
        ),
        use_container_width=True,
    )

    translay_download_col1, translay_download_col2 = st.columns(2)
    with translay_download_col1:
        st.download_button(
            "Export Translay settings CSV",
            data=translay_settings_df.to_csv(index=False).encode("utf-8"),
            file_name="translay_commissioning_settings.csv",
            mime="text/csv",
            use_container_width=True,
        )
    with translay_download_col2:
        st.download_button(
            "Export Translay test points CSV",
            data=translay_test_points_df.to_csv(index=False).encode("utf-8"),
            file_name="translay_commissioning_test_points.csv",
            mime="text/csv",
            use_container_width=True,
        )

    st.markdown("#### Suggested Three-Point Injection Plan")
    st.dataframe(
        translay_test_points_df.style.format(
            {
                "Through_Current_A": "{:.2f}",
                "Spill_Current_A": "{:.2f}",
                "Local_Current_A": "{:.2f}",
                "Remote_Current_A": "{:.2f}",
                "Local_Current_Primary_A": "{:.1f}",
                "Remote_Current_Primary_A": "{:.1f}",
                "Trip_Threshold_A": "{:.2f}",
                "Expected_Operate_s": "{:.3f}",
            }
        ),
        use_container_width=True,
    )
    st.caption(
        "Injection values are listed at relay secondary current level and converted to equivalent local/remote primary values using the entered CT ratios."
    )

    translay_max_through_a = max(
        float(translay_test_points_df["Through_Current_A"].max()) * 1.4,
        translay_pickup_a * 20.0,
        5.0,
    )
    translay_characteristic_df = pd.DataFrame(
        {
            "Through_Current_A": [translay_max_through_a * index / 100 for index in range(101)],
        }
    )
    translay_characteristic_df["Trip_Threshold_A"] = translay_characteristic_df[
        "Through_Current_A"
    ].apply(
        lambda through_value: translay_trip_threshold(
            through_current_a=through_value,
            pickup_a=translay_pickup_a,
            slope_pct=translay_slope_pct,
        )
    )

    st.markdown("#### Expected Translay Characteristic")
    translay_chart_layers = [
        alt.Chart(translay_characteristic_df)
        .mark_line()
        .encode(
            x=alt.X("Through_Current_A:Q", title="Through current (A secondary)", axis=alt.Axis(grid=True)),
            y=alt.Y("Trip_Threshold_A:Q", title="Spill operate threshold (A secondary)", axis=alt.Axis(grid=True)),
            tooltip=[
                alt.Tooltip("Through_Current_A:Q", title="Through current (A)", format=".2f"),
                alt.Tooltip("Trip_Threshold_A:Q", title="Threshold (A)", format=".2f"),
            ],
        )
    ]

    if translay_high_set_spill_a > 0:
        translay_high_set_df = pd.DataFrame(
            {
                "High_Set_Spill_A": [translay_high_set_spill_a],
            }
        )
        translay_chart_layers.append(
            alt.Chart(translay_high_set_df)
            .mark_rule(strokeDash=[4, 4], color="#D62728")
            .encode(
                y=alt.Y("High_Set_Spill_A:Q"),
                tooltip=[alt.Tooltip("High_Set_Spill_A:Q", title="High-set spill (A)", format=".2f")],
            )
        )

    translay_chart_layers.append(
        alt.Chart(translay_test_points_df)
        .mark_point(size=90, filled=True)
        .encode(
            x=alt.X("Through_Current_A:Q"),
            y=alt.Y("Spill_Current_A:Q"),
            color=alt.Color("Expected_Result:N", title="Expected"),
            tooltip=[
                alt.Tooltip("Point:N"),
                alt.Tooltip("Purpose:N"),
                alt.Tooltip("Through_Current_A:Q", title="Through current (A)", format=".2f"),
                alt.Tooltip("Spill_Current_A:Q", title="Spill current (A)", format=".2f"),
                alt.Tooltip("Expected_Result:N", title="Expected"),
                alt.Tooltip("Expected_Mode:N", title="Mode"),
            ],
        )
    )

    st.altair_chart(alt.layer(*translay_chart_layers).properties(height=420), use_container_width=True)

with formula_tab:
    st.subheader("Formula Reference")
    st.caption(
        "Equations used in the app model, including per-unit/base conversion, sequence fault calculations, "
        "relay timing, grading, commissioning calculations, and simplified arc-flash screening."
    )

    st.markdown("#### Per-Unit / Base Conversion")
    st.latex(r"Z_{\mathrm{base}} = \frac{V_{LL}^2}{S_{3\phi}}")
    st.latex(r"\lvert Z_{\mathrm{source,HV}}\rvert = \frac{V_{\mathrm{source}}^2}{S_{\mathrm{SC}}}")
    st.latex(r"Z_{\mathrm{referred}} = Z_{\mathrm{original}} \left(\frac{V_{\mathrm{study}}}{V_{\mathrm{original}}}\right)^2")
    st.caption("Use kV and MVA consistently; resulting impedance is in ohms.")

    st.markdown("#### R/X Split From Magnitude and X/R")
    st.latex(r"R = \frac{\lvert Z \rvert}{\sqrt{1 + (X/R)^2}}")
    st.latex(r"X = R\,(X/R)")
    st.latex(r"Z = R + jX")

    st.markdown("#### Sequence and Network Build-Up")
    st.latex(r"Z_{2} = Z_{1}")
    st.latex(r"Z_{0,\mathrm{source}} = k_{\mathrm{source}} Z_{1} \quad (k_{\mathrm{source}} = %.1f)" % SOURCE_Z0_FACTOR)
    st.latex(r"Z_{0,\mathrm{tr}} = k_{\mathrm{tr}} Z_{1} \quad (k_{\mathrm{tr}} = %.1f)" % TRANSFORMER_Z0_FACTOR)
    st.latex(r"Z_{1,\mathrm{cable}} = L\,z_1,\; Z_{2,\mathrm{cable}} = L\,z_2,\; Z_{0,\mathrm{cable}} = L\,z_0")
    st.latex(r"Z_{\mathrm{parallel,eq}} = \frac{Z}{n}")
    st.latex(
        r"Z_{\mathrm{source\to 1st\,11kV\,bus,single}} = Z_{\mathrm{reactor}} + Z_{33\mathrm{kV\,cable}} + Z_{\mathrm{tr}} + Z_{11\mathrm{kV\,tr\,cable(eq)}}"
    )
    st.latex(r"Z_{\mathrm{source\to 1st\,11kV\,bus,eq}} = \frac{Z_{\mathrm{source\to 1st\,11kV\,bus,single}}}{n_{\mathrm{source\,paths}}}")
    st.latex(r"Z_{\mathrm{incomer}} = Z_{\mathrm{source}} + Z_{\mathrm{source\to 1st\,11kV\,bus,eq}}")
    st.latex(r"Z_{\mathrm{remote}} = Z_{\mathrm{incomer}} + Z_{\mathrm{feeder(eq)}}")

    st.markdown("#### Fault Current Equations")
    st.latex(r"I_{3\phi} = \frac{V_{LL}}{\sqrt{3}\,\lvert Z_1 \rvert}")
    st.latex(r"I_{LL} = \frac{V_{LL}}{\lvert Z_1 + Z_2 \rvert}")
    st.latex(r"I_{LG} = \frac{\sqrt{3}\,V_{LL}}{\lvert Z_1 + Z_2 + Z_0 \rvert}")
    st.latex(r"\mathrm{Fault\,MVA} = \sqrt{3}\,V_{LL}\,I_{3\phi}")

    st.markdown("#### Relay Timing and Grading")
    st.latex(r"t_{\mathrm{IDMT}} = \frac{\mathrm{TMS}\,k}{\left(\frac{I}{I_p}\right)^{\alpha} - 1}")
    st.caption(
        "Relay logic used: no trip if I <= pickup or invalid settings; instantaneous element trips at 0.05 s when enabled and I >= Inst_A."
    )
    st.latex(r"t_{\mathrm{definite}} = t_{\mathrm{set}} \quad \mathrm{for}\ I > I_p")
    st.latex(r"\Delta t = t_{\mathrm{upstream}} - t_{\mathrm{downstream}}")
    st.latex(r"\mathrm{PASS\ if\ } \Delta t \ge t_{\mathrm{grading\,target}}")

    st.markdown("#### Commissioning Conversion, SEL-787, and Translay")
    st.latex(r"I_{\mathrm{primary}} = I_{\mathrm{secondary}}\,\frac{CT_{\mathrm{primary}}}{CT_{\mathrm{secondary}}}")
    st.latex(r"I_{\mathrm{inj}} = m\,I_p")
    st.latex(r"I_{\mathrm{op}} = \lvert I_1 - I_2 \rvert")
    st.latex(r"I_{\mathrm{rest}} = \frac{\lvert I_1 \rvert + \lvert I_2 \rvert}{2}")
    st.latex(r"I_{\mathrm{trip}} = I_{\mathrm{pickup}} + S_1 I_{\mathrm{rest}} \quad \mathrm{for}\ I_{\mathrm{rest}} \le I_{\mathrm{break}}")
    st.latex(r"I_{\mathrm{trip}} = I_{\mathrm{pickup}} + S_1 I_{\mathrm{break}} + S_2\left(I_{\mathrm{rest}} - I_{\mathrm{break}}\right) \quad \mathrm{for}\ I_{\mathrm{rest}} > I_{\mathrm{break}}")
    st.latex(r"I_{\mathrm{spill}} = \lvert I_{\mathrm{local}} - I_{\mathrm{remote}} \rvert")
    st.latex(r"I_{\mathrm{through}} = \frac{\lvert I_{\mathrm{local}} \rvert + \lvert I_{\mathrm{remote}} \rvert}{2}")
    st.latex(r"I_{\mathrm{trip,translay}} = I_{\mathrm{pickup}} + S\,I_{\mathrm{through}}")
    st.latex(r"\mathrm{Trip\ if}\ I_{\mathrm{spill}} \ge I_{\mathrm{trip,translay}}")

    st.markdown("#### Simplified 11kV Arc-Flash Screening")
    st.latex(r"I_{\mathrm{arc}} = I_{\mathrm{bolted}}\,f_{\mathrm{arc}}")
    st.latex(r"E_{\mathrm{cal/cm^2}} = 0.2\,I_{\mathrm{arc}}\,t_{\mathrm{clear}}\,C_{\mathrm{config}}\left(\frac{610}{D_{\mathrm{mm}}}\right)^{1.5}")
    st.latex(r"\mathrm{AFB}_{\mathrm{mm}} = 610\left(\frac{E_{\mathrm{cal/cm^2}}}{1.2}\right)^{\frac{1}{1.5}}")
    st.caption("This arc-flash section is an early-stage screening approximation, not a full IEEE 1584 study.")

persist_last_entered_state()
