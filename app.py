import math

import pandas as pd
import streamlit as st


CURVE_CONSTANTS = {
    "IEC Standard Inverse": (0.14, 0.02),
    "IEC Very Inverse": (13.5, 1.0),
    "IEC Extremely Inverse": (80.0, 2.0),
    "IEC Long-Time Inverse": (120.0, 1.0),
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


st.set_page_config(
    page_title="Protection",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.title("🛡️ Protection")
st.caption(
    "ETAP-style protection study MVP: network parameters, fault levels, protection grading, and arc-flash screening."
)

with st.sidebar:
    st.header("Study Controls")
    project_name = st.text_input("Project name", value="Protection")
    grading_margin_s = st.number_input(
        "Target grading margin (s)",
        min_value=0.10,
        max_value=1.00,
        value=0.30,
        step=0.05,
    )
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
    col1, col2, col3 = st.columns(3)

    with col1:
        v_ll_kv = st.number_input(
            "System voltage at study bus (kV)",
            min_value=0.4,
            max_value=66.0,
            value=11.0,
            step=0.1,
        )
        frequency_hz = st.selectbox("Frequency (Hz)", options=[50, 60], index=0)
        source_sc_mva = st.number_input(
            "Upstream source short-circuit level (MVA)",
            min_value=10.0,
            max_value=50000.0,
            value=500.0,
            step=10.0,
        )
        source_xr = st.number_input("Source X/R ratio", min_value=0.0, max_value=50.0, value=10.0, step=0.5)

    with col2:
        transformer_mva = st.number_input(
            "Transformer rating (MVA)",
            min_value=0.1,
            max_value=1000.0,
            value=10.0,
            step=0.1,
        )
        transformer_z_pct = st.number_input(
            "Transformer impedance (%Z)",
            min_value=1.0,
            max_value=25.0,
            value=8.0,
            step=0.1,
        )
        transformer_xr = st.number_input(
            "Transformer X/R ratio",
            min_value=0.0,
            max_value=50.0,
            value=8.0,
            step=0.5,
        )

    with col3:
        feeder_length_km = st.number_input(
            "Feeder length (km)",
            min_value=0.0,
            max_value=100.0,
            value=0.5,
            step=0.1,
        )
        feeder_r_ohm_per_km = st.number_input(
            "Feeder R (Ω/km)",
            min_value=0.0,
            max_value=5.0,
            value=0.08,
            step=0.01,
        )
        feeder_x_ohm_per_km = st.number_input(
            "Feeder X (Ω/km)",
            min_value=0.0,
            max_value=5.0,
            value=0.10,
            step=0.01,
        )

    st.markdown("#### Fault Type Factors")
    factor_col1, factor_col2 = st.columns(2)
    with factor_col1:
        ll_factor = st.slider("Line-Line current factor (vs 3-phase)", 0.50, 1.00, 0.87, 0.01)
    with factor_col2:
        lg_factor = st.slider("Line-Ground current factor (vs 3-phase)", 0.50, 1.20, 0.95, 0.01)

source_z_mag = (v_ll_kv**2) / source_sc_mva
source_z = split_rx_from_z_magnitude(source_z_mag, source_xr)

transformer_z_mag = (transformer_z_pct / 100) * ((v_ll_kv**2) / transformer_mva)
transformer_z = split_rx_from_z_magnitude(transformer_z_mag, transformer_xr)

feeder_z = complex(
    feeder_length_km * feeder_r_ohm_per_km,
    feeder_length_km * feeder_x_ohm_per_km,
)

z_incomer = source_z + transformer_z
z_load = z_incomer + feeder_z

incomer_fault = calc_fault_values(v_ll_kv, z_incomer, ll_factor, lg_factor)
load_fault = calc_fault_values(v_ll_kv, z_load, ll_factor, lg_factor)

with network_tab:
    st.markdown("#### Equivalent Impedance Build-Up")
    impedance_df = pd.DataFrame(
        [
            {
                "Element": "Source equivalent",
                "R (Ω)": source_z.real,
                "X (Ω)": source_z.imag,
                "|Z| (Ω)": abs(source_z),
            },
            {
                "Element": "Transformer equivalent",
                "R (Ω)": transformer_z.real,
                "X (Ω)": transformer_z.imag,
                "|Z| (Ω)": abs(transformer_z),
            },
            {
                "Element": "Feeder",
                "R (Ω)": feeder_z.real,
                "X (Ω)": feeder_z.imag,
                "|Z| (Ω)": abs(feeder_z),
            },
            {
                "Element": "Total to incomer bus",
                "R (Ω)": z_incomer.real,
                "X (Ω)": z_incomer.imag,
                "|Z| (Ω)": abs(z_incomer),
            },
            {
                "Element": "Total to load bus",
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
    st.caption(f"Nominal frequency selected: {frequency_hz} Hz")

fault_df = pd.DataFrame(
    [
        {
            "Bus": "Incomer Bus",
            "Z_total_ohm": incomer_fault["Z_total_ohm"],
            "I_3ph_kA": incomer_fault["I_3ph_kA"],
            "I_LL_kA": incomer_fault["I_LL_kA"],
            "I_LG_kA": incomer_fault["I_LG_kA"],
            "Fault_MVA": incomer_fault["Fault_MVA"],
        },
        {
            "Bus": "Load Bus",
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
        study_bus = st.selectbox("Select fault location for protection checks", options=fault_df["Bus"].tolist())
    with study_col2:
        fault_type = st.selectbox("Select fault type", options=["3-Phase", "Line-Line", "Line-Ground"])

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
    edited_df = st.data_editor(
        st.session_state.relay_settings,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "Order": st.column_config.NumberColumn("Order", min_value=1, step=1),
            "Device": st.column_config.TextColumn("Device"),
            "Curve": st.column_config.SelectboxColumn("Curve", options=list(CURVE_CONSTANTS.keys())),
            "Pickup_A": st.column_config.NumberColumn("Pickup (A)", min_value=1.0),
            "TMS": st.column_config.NumberColumn("TMS", min_value=0.01),
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

    arc_col1, arc_col2, arc_col3 = st.columns(3)
    with arc_col1:
        bolted_fault_current_ka = st.number_input(
            "Bolted fault current (kA)",
            min_value=0.1,
            max_value=200.0,
            value=float(round(selected_fault_current_ka, 2)),
            step=0.1,
        )
        arc_current_factor = st.slider("Arc current factor", 0.50, 1.00, 0.85, 0.01)

    with arc_col2:
        clearing_time_s = st.number_input(
            "Clearing time (s)",
            min_value=0.01,
            max_value=2.00,
            value=float(round(suggested_clearing_time, 3)),
            step=0.01,
        )
        working_distance_mm = st.number_input(
            "Working distance (mm)",
            min_value=200,
            max_value=2000,
            value=455,
            step=25,
        )

    with arc_col3:
        enclosure_type = st.selectbox(
            "Enclosure",
            options=["Open air", "Switchboard", "Metal-clad cubicle"],
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
