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

RELAY_EXPORT_COLUMNS = ["Order", "Device", "Curve", "Pickup_A", "TMS", "Inst_A"]
APP_STATE_FILE = Path(".streamlit/last_entered_state.json")

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
    "frequency_hz": {50, 60},
    "study_bus": {"11kV Transformer Busbar", "11kV Remote Busbar"},
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

    payload = {
        "study_case": study_case_data,
        "relay_settings": relay_df[RELAY_EXPORT_COLUMNS].to_dict(orient="records"),
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


def load_network_preset() -> None:
    for key, value in DEFAULT_NETWORK_PRESET.items():
        st.session_state[key] = value


def restore_original_defaults() -> None:
    load_network_preset()
    st.session_state.relay_settings = default_relay_settings_df()
    st.session_state.pop("relay_settings_editor", None)


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


st.set_page_config(
    page_title="Protection",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded",
)

initialize_app_state()

st.title("🛡️ Protection")
st.caption(
    "ETAP-style protection study MVP: network parameters, fault levels, protection grading, and arc-flash screening."
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

network_tab, fault_tab, protection_tab, arc_tab, full_arc_tab, formula_tab = st.tabs(
    [
        "1) Network Inputs",
        "2) Fault Levels",
        "3) Protection & Grading",
        "4) Arc-Flash Estimate",
        "5) Full Arc-Flash Calc",
        "6) Formula Reference",
    ]
)

with network_tab:
    st.subheader("Power Network Parameters")
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

if st.session_state.get("study_bus") not in STUDY_CASE_OPTION_VALUES["study_bus"]:
    st.session_state["study_bus"] = "11kV Transformer Busbar"

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

with network_tab:
    st.markdown("#### Selected Cable Sequence Data (Ω/km)")
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
    st.dataframe(cable_df, use_container_width=True)

    st.markdown("#### Cable Library (Best Available Legacy Data, Ω/km)")
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
    st.dataframe(cable_library_df, use_container_width=True)

    st.markdown("#### Sequence Impedance Build-Up (Referred to Study Voltage)")
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

with fault_tab:
    st.subheader("Calculated Fault Levels")
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
        study_bus = st.selectbox(
            "Select fault location for protection checks",
            options=fault_df["Bus"].tolist(),
            key="study_bus",
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

selected_fault_current_ka = float(
    fault_df.loc[fault_df["Bus"] == study_bus, fault_column_map[fault_type]].iloc[0]
)
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
    else:
        min_pickup = max(float(relay_df["Pickup_A"].min()) * 1.05, 1.0)
        max_fault_current_a = incomer_fault["I_3ph_kA"] * 1000
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

with formula_tab:
    st.subheader("Formula Reference")
    st.caption(
        "Equations used in the app model, including per-unit/base conversion, sequence fault calculations, "
        "relay timing, grading, and simplified arc-flash screening."
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
    st.latex(r"\Delta t = t_{\mathrm{upstream}} - t_{\mathrm{downstream}}")
    st.latex(r"\mathrm{PASS\ if\ } \Delta t \ge t_{\mathrm{grading\,target}}")

    st.markdown("#### Simplified 11kV Arc-Flash Screening")
    st.latex(r"I_{\mathrm{arc}} = I_{\mathrm{bolted}}\,f_{\mathrm{arc}}")
    st.latex(r"E_{\mathrm{cal/cm^2}} = 0.2\,I_{\mathrm{arc}}\,t_{\mathrm{clear}}\,C_{\mathrm{config}}\left(\frac{610}{D_{\mathrm{mm}}}\right)^{1.5}")
    st.latex(r"\mathrm{AFB}_{\mathrm{mm}} = 610\left(\frac{E_{\mathrm{cal/cm^2}}}{1.2}\right)^{\frac{1}{1.5}}")
    st.caption("This arc-flash section is an early-stage screening approximation, not a full IEEE 1584 study.")

persist_last_entered_state()
