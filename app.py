import hashlib
import io
import math
from typing import Optional, Tuple

import pandas as pd
import streamlit as st


CURVE_CONSTANTS = {
    "IEC Standard Inverse": (0.14, 0.02),
    "IEC Very Inverse": (13.5, 1.0),
    "IEC Extremely Inverse": (80.0, 2.0),
    "IEC Long-Time Inverse": (120.0, 1.0),
}

RELAY_EXPORT_COLUMNS = ["Order", "Device", "Curve", "Pickup_A", "TMS", "Inst_A"]

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
    "hv_cable_r_ohm_per_km": 0.09,
    "hv_cable_x_ohm_per_km": 0.10,
    "transformer_hv_kv": 33.0,
    "transformer_lv_kv": 11.0,
    "transformer_mva": 13.5,
    "transformer_z_pct": 7.2,
    "transformer_xr": 8.0,
    "lv_cable_length_km": 0.01,
    "lv_cable_r_ohm_per_km": 0.08,
    "lv_cable_x_ohm_per_km": 0.10,
    "feeder_length_km": 0.5,
    "feeder_r_ohm_per_km": 0.08,
    "feeder_x_ohm_per_km": 0.10,
    "ll_factor": 0.87,
    "lg_factor": 0.95,
    "study_bus": "11kV Transformer Busbar",
    "fault_type": "3-Phase",
    "bolted_fault_current_ka": 10.0,
    "arc_current_factor": 0.85,
    "clearing_time_s": 0.20,
    "working_distance_mm": 455,
    "enclosure_type": "Switchboard",
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
    "enclosure_type": {"Open air", "Switchboard", "Metal-clad cubicle"},
}

STUDY_CASE_BOUNDS = {
    "grading_margin_s": (0.10, 1.00),
    "v_ll_kv": (0.4, 66.0),
    "source_v_kv": (1.0, 400.0),
    "source_sc_mva": (10.0, 50000.0),
    "source_xr": (0.0, 50.0),
    "reactor_ohm": (0.0, 50.0),
    "hv_cable_length_km": (0.0, 100.0),
    "hv_cable_r_ohm_per_km": (0.0, 5.0),
    "hv_cable_x_ohm_per_km": (0.0, 5.0),
    "transformer_hv_kv": (1.0, 400.0),
    "transformer_lv_kv": (0.4, 66.0),
    "transformer_mva": (0.1, 1000.0),
    "transformer_z_pct": (1.0, 25.0),
    "transformer_xr": (0.0, 50.0),
    "lv_cable_length_km": (0.0, 20.0),
    "lv_cable_r_ohm_per_km": (0.0, 5.0),
    "lv_cable_x_ohm_per_km": (0.0, 5.0),
    "feeder_length_km": (0.0, 100.0),
    "feeder_r_ohm_per_km": (0.0, 5.0),
    "feeder_x_ohm_per_km": (0.0, 5.0),
    "ll_factor": (0.50, 1.00),
    "lg_factor": (0.50, 1.20),
    "bolted_fault_current_ka": (0.1, 200.0),
    "arc_current_factor": (0.50, 1.00),
    "clearing_time_s": (0.01, 2.00),
    "working_distance_mm": (200, 2000),
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


def calc_fault_values(v_ll_kv: float, z_total: complex, ll_factor: float, lg_factor: float) -> dict:
    z_magnitude = abs(z_total)
    if z_magnitude <= 0 or v_ll_kv <= 0:
        return {
            "Z_total_ohm": 0.0,
            "I_3ph_kA": math.inf,
            "I_LL_kA": math.inf,
            "I_LG_kA": math.inf,
            "Fault_MVA": math.inf,
        }

    i_3ph_ka = v_ll_kv / (math.sqrt(3) * z_magnitude)
    i_ll_ka = i_3ph_ka * ll_factor
    i_lg_ka = i_3ph_ka * lg_factor
    fault_mva = math.sqrt(3) * v_ll_kv * i_3ph_ka

    return {
        "Z_total_ohm": z_magnitude,
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
        return settings_df

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
    return relay_df


def initialize_study_case_state() -> None:
    for key, default_value in DEFAULT_STUDY_CASE.items():
        if key not in st.session_state:
            st.session_state[key] = default_value


def load_network_preset() -> None:
    for key, value in DEFAULT_NETWORK_PRESET.items():
        st.session_state[key] = value


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

    updates: dict[str, object] = {}

    for key, default_value in DEFAULT_STUDY_CASE.items():
        if key not in parameter_map:
            continue

        raw_value = parameter_map[key]
        if pd.isna(raw_value):
            continue

        try:
            if key in STUDY_CASE_INT_FIELDS:
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

initialize_study_case_state()

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

    if st.button("Load 33/11kV default network", use_container_width=True):
        load_network_preset()
        st.session_state["study_case_upload_status"] = "success"
        st.session_state["study_case_upload_message"] = "Loaded 33/11kV default network parameters."
        st.rerun()

    st.caption(f"Active study: {project_name}")

if "relay_settings" not in st.session_state:
    st.session_state.relay_settings = pd.DataFrame(
        [
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
    )

network_tab, fault_tab, protection_tab, arc_tab = st.tabs(
    [
        "1) Network Inputs",
        "2) Fault Levels",
        "3) Protection & Grading",
        "4) Arc-Flash Estimate",
    ]
)

with network_tab:
    st.subheader("Power Network Parameters")
    st.caption("All impedances are converted to the selected study voltage base before fault calculations.")
    st.markdown("#### One-line arrangement (default template)")
    st.caption(" → ".join(NETWORK_ARRANGEMENT))
    st.info(
        "Cable conductor type is represented through entered R/X values. "
        "Update cable Ω/km values from your manufacturer data for final studies."
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
        hv_cable_r_ohm_per_km = st.number_input(
            "33kV cable R (Ω/km)",
            min_value=0.0,
            max_value=5.0,
            step=0.01,
            key="hv_cable_r_ohm_per_km",
        )
        hv_cable_x_ohm_per_km = st.number_input(
            "33kV cable X (Ω/km)",
            min_value=0.0,
            max_value=5.0,
            step=0.01,
            key="hv_cable_x_ohm_per_km",
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
        lv_cable_r_ohm_per_km = st.number_input(
            "11kV cable R (Ω/km)",
            min_value=0.0,
            max_value=5.0,
            step=0.01,
            key="lv_cable_r_ohm_per_km",
        )
        lv_cable_x_ohm_per_km = st.number_input(
            "11kV cable X (Ω/km)",
            min_value=0.0,
            max_value=5.0,
            step=0.01,
            key="lv_cable_x_ohm_per_km",
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
        feeder_r_ohm_per_km = st.number_input(
            "Feeder R (Ω/km)",
            min_value=0.0,
            max_value=5.0,
            step=0.01,
            key="feeder_r_ohm_per_km",
        )
        feeder_x_ohm_per_km = st.number_input(
            "Feeder X (Ω/km)",
            min_value=0.0,
            max_value=5.0,
            step=0.01,
            key="feeder_x_ohm_per_km",
        )

    st.markdown("#### Fault Type Factors")
    factor_col1, factor_col2 = st.columns(2)
    with factor_col1:
        ll_factor = st.slider(
            "Line-Line current factor (vs 3-phase)",
            0.50,
            1.00,
            step=0.01,
            key="ll_factor",
        )
    with factor_col2:
        lg_factor = st.slider(
            "Line-Ground current factor (vs 3-phase)",
            0.50,
            1.20,
            step=0.01,
            key="lg_factor",
        )

if st.session_state.get("study_bus") not in STUDY_CASE_OPTION_VALUES["study_bus"]:
    st.session_state["study_bus"] = "11kV Transformer Busbar"

source_z_hv = split_rx_from_z_magnitude((source_v_kv**2) / source_sc_mva, source_xr)
source_to_study_factor = (v_ll_kv / source_v_kv) ** 2 if source_v_kv > 0 else 0.0
source_z = source_z_hv * source_to_study_factor

reactor_z_hv = complex(0.0, reactor_ohm)
reactor_z = reactor_z_hv * source_to_study_factor

hv_cable_z_hv = complex(
    hv_cable_length_km * hv_cable_r_ohm_per_km,
    hv_cable_length_km * hv_cable_x_ohm_per_km,
)
hv_cable_z = hv_cable_z_hv * source_to_study_factor

transformer_z_lv = split_rx_from_z_magnitude(
    (transformer_z_pct / 100) * ((transformer_lv_kv**2) / transformer_mva),
    transformer_xr,
)
lv_to_study_factor = (v_ll_kv / transformer_lv_kv) ** 2 if transformer_lv_kv > 0 else 0.0
transformer_z = transformer_z_lv * lv_to_study_factor

lv_cable_z_lv = complex(
    lv_cable_length_km * lv_cable_r_ohm_per_km,
    lv_cable_length_km * lv_cable_x_ohm_per_km,
)
lv_cable_z = lv_cable_z_lv * lv_to_study_factor

feeder_z = complex(
    feeder_length_km * feeder_r_ohm_per_km,
    feeder_length_km * feeder_x_ohm_per_km,
)

z_incomer = source_z + reactor_z + hv_cable_z + transformer_z + lv_cable_z
z_load = z_incomer + feeder_z

incomer_fault = calc_fault_values(v_ll_kv, z_incomer, ll_factor, lg_factor)
load_fault = calc_fault_values(v_ll_kv, z_load, ll_factor, lg_factor)

with network_tab:
    st.markdown("#### Equivalent Impedance Build-Up")
    impedance_df = pd.DataFrame(
        [
            {
                "Element": "Source equivalent (referred)",
                "R (Ω)": source_z.real,
                "X (Ω)": source_z.imag,
                "|Z| (Ω)": abs(source_z),
            },
            {
                "Element": "33kV reactor (referred)",
                "R (Ω)": reactor_z.real,
                "X (Ω)": reactor_z.imag,
                "|Z| (Ω)": abs(reactor_z),
            },
            {
                "Element": "33kV cable (referred)",
                "R (Ω)": hv_cable_z.real,
                "X (Ω)": hv_cable_z.imag,
                "|Z| (Ω)": abs(hv_cable_z),
            },
            {
                "Element": "Transformer equivalent (referred)",
                "R (Ω)": transformer_z.real,
                "X (Ω)": transformer_z.imag,
                "|Z| (Ω)": abs(transformer_z),
            },
            {
                "Element": "11kV cable from transformer",
                "R (Ω)": lv_cable_z.real,
                "X (Ω)": lv_cable_z.imag,
                "|Z| (Ω)": abs(lv_cable_z),
            },
            {
                "Element": "11kV feeder",
                "R (Ω)": feeder_z.real,
                "X (Ω)": feeder_z.imag,
                "|Z| (Ω)": abs(feeder_z),
            },
            {
                "Element": "Total to 11kV transformer busbar",
                "R (Ω)": z_incomer.real,
                "X (Ω)": z_incomer.imag,
                "|Z| (Ω)": abs(z_incomer),
            },
            {
                "Element": "Total to 11kV remote busbar",
                "R (Ω)": z_load.real,
                "X (Ω)": z_load.imag,
                "|Z| (Ω)": abs(z_load),
            },
        ]
    )
    st.dataframe(
        impedance_df.style.format({"R (Ω)": "{:.4f}", "X (Ω)": "{:.4f}", "|Z| (Ω)": "{:.4f}"}),
        use_container_width=True,
    )
    st.caption(
        f"Nominal frequency: {frequency_hz} Hz | Source base: {source_v_kv:.1f} kV | Transformer: {transformer_hv_kv:.1f}/{transformer_lv_kv:.1f} kV"
    )

fault_df = pd.DataFrame(
    [
        {
            "Bus": "11kV Transformer Busbar",
            "Z_total_ohm": incomer_fault["Z_total_ohm"],
            "I_3ph_kA": incomer_fault["I_3ph_kA"],
            "I_LL_kA": incomer_fault["I_LL_kA"],
            "I_LG_kA": incomer_fault["I_LG_kA"],
            "Fault_MVA": incomer_fault["Fault_MVA"],
        },
        {
            "Bus": "11kV Remote Busbar",
            "Z_total_ohm": load_fault["Z_total_ohm"],
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
                "Z_total_ohm": "{:.4f}",
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
                st.session_state.relay_settings = imported_relay_df
                st.session_state["relay_upload_status"] = "success"
            else:
                st.session_state["relay_upload_status"] = "error"
            st.session_state["relay_upload_message"] = relay_message

    if st.session_state.get("relay_upload_message"):
        if st.session_state.get("relay_upload_status") == "success":
            st.success(st.session_state["relay_upload_message"])
        else:
            st.error(st.session_state["relay_upload_message"])

    edited_df = st.data_editor(
        st.session_state.relay_settings,
        num_rows="dynamic",
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
    st.session_state.relay_settings = edited_df

    relay_df = sanitize_relay_settings(pd.DataFrame(edited_df))

    if relay_df.empty:
        st.warning("Add at least one valid protection device setting to run grading checks.")
        relay_results_df = pd.DataFrame()
    else:
        relay_results = []
        for _, row in relay_df.iterrows():
            operate_time_s = calc_relay_time_s(
                current_a=selected_fault_current_a,
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

        st.markdown("#### Relay Operating Times at Selected Fault")
        st.dataframe(
            relay_results_df[
                ["Order", "Device", "Curve", "Pickup_A", "TMS", "Inst_A", "Operate_display"]
            ],
            use_container_width=True,
        )

        grading_rows = []
        for index in range(len(relay_results_df) - 1):
            downstream = relay_results_df.iloc[index]
            upstream = relay_results_df.iloc[index + 1]

            if math.isfinite(float(downstream["Operate_s"])) and math.isfinite(float(upstream["Operate_s"])):
                margin_s = float(upstream["Operate_s"]) - float(downstream["Operate_s"])
                status = "PASS" if margin_s >= grading_margin_s else "FAIL"
            else:
                margin_s = math.nan
                status = "CHECK"

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
        st.markdown("#### Grading Margins")
        st.dataframe(
            grading_df.style.format({"Margin_s": "{:.3f}", "Target_s": "{:.3f}"}),
            use_container_width=True,
        )

        if not grading_df.empty and (grading_df["Status"] == "FAIL").any():
            st.error("One or more relay pairs do not meet the grading margin target.")
        elif not grading_df.empty:
            st.success("All checked relay pairs meet the grading margin target.")

        min_pickup = max(float(relay_df["Pickup_A"].min()) * 1.05, 1.0)
        max_current = max(
            selected_fault_current_a * 1.30,
            float(relay_df["Inst_A"].max()) * 1.20,
            min_pickup * 5,
        )
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
                    inst_pickup_a=float(row["Inst_A"]),
                )
                for value in current_points
            ]
            cleaned_times = [time if math.isfinite(time) and time <= 30 else None for time in times]
            tcc_df[str(row["Device"])] = cleaned_times

        st.markdown("#### Time-Current Curves (Approximate)")
        st.line_chart(tcc_df.set_index("Current_A"), use_container_width=True)

with arc_tab:
    st.subheader("Arc-Flash Screening Estimate")
    st.caption(
        "This is a simplified screening model for early design checks. "
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
            max_value=2.00,
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
        enclosure_type = st.selectbox(
            "Enclosure",
            options=["Open air", "Switchboard", "Metal-clad cubicle"],
            key="enclosure_type",
        )
        enclosure_factor_map = {
            "Open air": 1.0,
            "Switchboard": 1.3,
            "Metal-clad cubicle": 1.5,
        }
        enclosure_factor = enclosure_factor_map[enclosure_type]

    arc_current_ka = bolted_fault_current_ka * arc_current_factor
    distance_multiplier = (610 / max(working_distance_mm, 1)) ** 1.5
    incident_energy_cal_cm2 = 0.2 * arc_current_ka * clearing_time_s * enclosure_factor * distance_multiplier
    arc_flash_boundary_mm = 610 * ((max(incident_energy_cal_cm2, 0.001) / 1.2) ** (1 / 1.5))

    result_col1, result_col2, result_col3 = st.columns(3)
    with result_col1:
        st.metric("Arcing current", f"{arc_current_ka:.2f} kA")
    with result_col2:
        st.metric("Incident energy", f"{incident_energy_cal_cm2:.2f} cal/cm²")
    with result_col3:
        st.metric("Arc flash boundary", f"{arc_flash_boundary_mm:.0f} mm")

    st.info(arc_flash_category(incident_energy_cal_cm2))
