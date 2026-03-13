# Protection

Current release: Rev 0.1.

[![CI](https://github.com/flyingdave/Protection/actions/workflows/ci.yml/badge.svg)](https://github.com/flyingdave/Protection/actions/workflows/ci.yml)

An ETAP-style web GUI prototype built with [Streamlit](https://streamlit.io/) and Python.

## Features

- Accepts detailed network parameters including 33kV source, reactor, HV cable, 33/11kV transformer, LV cable, and feeder.
- Uses selectable site cable types with built-in best-available legacy sequence impedance data (Z1, Z2, Z0).
- Includes a one-click 33/11kV default network template (1850 MVA source, 3.23 Ω reactor, 3 km HV cable, 13.5 MVA transformer).
- Calculates fault levels at 11kV transformer and remote busbars (3-phase, line-line, line-ground).
- Allows entry of protection relay settings (pickup, TMS, IEC inverse curves, instantaneous).
- Displays operating times and grading margins between downstream/upstream devices.
- Provides a simplified arc-flash screening estimate (arcing current, incident energy, boundary).
- Supports CSV import/export for study cases and relay settings.
- Remembers last entered study/relay data as startup defaults.
- Includes a Restore Original Defaults action to return to the built-in base case.

## Scope Note

This tool is an MVP for study/screening workflows. It is **not** a substitute for a full protection and arc-flash study using utility data, detailed sequence network modeling, and formal IEEE 1584 / IEC / NFPA methods.

## Getting Started

### Prerequisites

- Python 3.9+

### Installation

```bash
git clone https://github.com/flyingdave/Protection.git
cd Protection

python -m venv .venv
.venv\Scripts\activate   # Windows
# source .venv/bin/activate  # macOS/Linux

pip install -r requirements.txt
```

### Run

```bash
streamlit run app.py
```

Then open `http://localhost:8501`.

## Study Flow

1. Enter network parameters in **Network Inputs**.
2. Review calculated fault levels in **Fault Levels**.
3. Enter relay settings and check grading in **Protection & Grading**.
4. Run arc-flash screening in **Arc-Flash Estimate**.
