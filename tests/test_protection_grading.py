import contextlib
import importlib.util
import io
import math
import unittest
from pathlib import Path

import pandas as pd

APP_PATH = Path(__file__).resolve().parents[1] / "app.py"
APP_SPEC = importlib.util.spec_from_file_location("app", APP_PATH)
if APP_SPEC is None or APP_SPEC.loader is None:
    raise RuntimeError(f"Unable to load app module from {APP_PATH}")
app = importlib.util.module_from_spec(APP_SPEC)
with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
    APP_SPEC.loader.exec_module(app)


class ProtectionGradingBehaviorTests(unittest.TestCase):
    def test_tcc_curve_stops_after_inst_pickup_and_drops_vertically(self) -> None:
        current_points = [120.0, 200.0, 400.0, 800.0, 1200.0, 1600.0]
        points = app.build_relay_tcc_points(
            current_points=current_points,
            pickup_a=100.0,
            tms=0.1,
            curve_name="IEC Standard Inverse",
            inst_pickup_a=800.0,
        )

        self.assertTrue(points)
        self.assertTrue(all(float(point["Current_A"]) <= 800.0 + 1e-9 for point in points))

        inst_points = [
            point
            for point in points
            if math.isclose(float(point["Current_A"]), 800.0, rel_tol=0.0, abs_tol=1e-9)
        ]
        self.assertGreaterEqual(len(inst_points), 2)

        inst_points_by_sequence = sorted(inst_points, key=lambda point: int(point["Seq"]))
        self.assertGreater(
            float(inst_points_by_sequence[0]["Operate_s"]),
            float(inst_points_by_sequence[-1]["Operate_s"]),
        )
        self.assertTrue(
            math.isclose(
                float(inst_points_by_sequence[-1]["Operate_s"]),
                float(app.TCC_AXIS_FLOOR_S),
                rel_tol=0.0,
                abs_tol=1e-12,
            )
        )

    def test_tcc_curve_without_inst_pickup_keeps_high_current_points(self) -> None:
        current_points = [120.0, 200.0, 400.0, 800.0]
        points = app.build_relay_tcc_points(
            current_points=current_points,
            pickup_a=100.0,
            tms=0.1,
            curve_name="IEC Standard Inverse",
            inst_pickup_a=0.0,
        )

        current_values = [float(point["Current_A"]) for point in points]
        self.assertIn(800.0, current_values)

    def test_normalize_relay_editor_rows_preserves_editable_rows(self) -> None:
        editor_df = pd.DataFrame(
            [
                {
                    "Order": "1",
                    "Device": "Feeder A",
                    "Curve": "IEC Standard Inverse",
                    "Pickup_A": "",
                    "TMS": "0.10",
                    "Inst_A": "",
                }
            ]
        )

        normalized_df = app.normalize_relay_editor_rows(editor_df)

        self.assertListEqual(list(normalized_df.columns), app.RELAY_EXPORT_COLUMNS)
        self.assertEqual(len(normalized_df), 1)
        self.assertEqual(str(normalized_df.loc[0, "Device"]), "Feeder A")


if __name__ == "__main__":
    unittest.main()
