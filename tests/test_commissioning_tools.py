import contextlib
import io
import math
import unittest

with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
    import app


class CommissioningToolTests(unittest.TestCase):
    def test_definite_time_stage_returns_set_time_above_pickup(self) -> None:
        self.assertTrue(
            math.isinf(
                app.calc_commissioning_time_s(
                    current_a=0.95,
                    pickup_a=1.0,
                    curve_name="Definite Time",
                    tms=0.1,
                    definite_time_s=0.35,
                    inst_pickup_a=0.0,
                )
            )
        )
        self.assertAlmostEqual(
            app.calc_commissioning_time_s(
                current_a=2.0,
                pickup_a=1.0,
                curve_name="Definite Time",
                tms=0.1,
                definite_time_s=0.35,
                inst_pickup_a=0.0,
            ),
            0.35,
            places=9,
        )

    def test_high_set_overrides_inverse_curve_time(self) -> None:
        self.assertAlmostEqual(
            app.calc_commissioning_time_s(
                current_a=10.0,
                pickup_a=1.0,
                curve_name="IEC Standard Inverse",
                tms=0.1,
                definite_time_s=0.0,
                inst_pickup_a=8.0,
            ),
            0.05,
            places=9,
        )

    def test_oc_ef_points_follow_pickup_multipliers(self) -> None:
        test_points_df = app.build_oc_ef_test_points(
            pickup_secondary_a=1.0,
            curve_name="IEC Standard Inverse",
            tms=0.1,
            definite_time_s=0.0,
            inst_pickup_a=0.0,
            ct_primary_a=200.0,
            ct_secondary_a=1.0,
        )

        self.assertEqual(len(test_points_df), 3)
        self.assertListEqual(
            [round(value, 6) for value in test_points_df["Injection_Current_Secondary_A"].tolist()],
            [1.5, 3.0, 6.0],
        )
        self.assertTrue((test_points_df["Expected_Operate_s"] > 0).all())

    def test_sel787_threshold_is_piecewise(self) -> None:
        self.assertAlmostEqual(
            app.sel787_operate_threshold(2.0, 0.3, 25.0, 50.0, 4.0),
            0.3 + (0.25 * 2.0),
            places=9,
        )
        self.assertAlmostEqual(
            app.sel787_operate_threshold(6.0, 0.3, 25.0, 50.0, 4.0),
            0.3 + (0.25 * 4.0) + (0.50 * 2.0),
            places=9,
        )

    def test_sel787_test_points_include_stability_and_operate_cases(self) -> None:
        test_points_df = app.build_sel787_test_points(
            pickup_a=0.3,
            slope1_pct=25.0,
            slope2_pct=50.0,
            breakpoint_a=4.0,
            unrestrained_a=8.0,
        )

        self.assertEqual(len(test_points_df), 3)
        self.assertEqual(test_points_df.iloc[0]["Expected_Result"], "No trip")
        self.assertEqual(test_points_df.iloc[1]["Expected_Result"], "Trip")
        self.assertEqual(test_points_df.iloc[2]["Expected_Result"], "Trip")

    def test_translay_options_include_h04_and_newer(self) -> None:
        self.assertIn("H04", app.TRANSLAY_RELAY_OPTIONS)
        self.assertIn("Numerical (Newer)", app.TRANSLAY_RELAY_OPTIONS)

    def test_translay_threshold_increases_with_through_current(self) -> None:
        threshold_low = app.translay_trip_threshold(through_current_a=1.0, pickup_a=0.2, slope_pct=20.0)
        threshold_high = app.translay_trip_threshold(through_current_a=5.0, pickup_a=0.2, slope_pct=20.0)
        self.assertGreater(threshold_high, threshold_low)

    def test_translay_test_points_include_stability_and_operate_checks(self) -> None:
        test_points_df = app.build_translay_test_points(
            pickup_a=0.2,
            slope_pct=20.0,
            operate_time_s=0.1,
            high_set_spill_a=0.0,
        )

        self.assertEqual(len(test_points_df), 3)
        self.assertEqual(test_points_df.iloc[0]["Expected_Result"], "No trip")
        self.assertEqual(test_points_df.iloc[1]["Expected_Result"], "Trip")
        self.assertEqual(test_points_df.iloc[2]["Expected_Result"], "Trip")
        self.assertTrue(math.isinf(float(test_points_df.iloc[0]["Expected_Operate_s"])))
        self.assertGreater(float(test_points_df.iloc[2]["Spill_Current_A"]), 0.0)


if __name__ == "__main__":
    unittest.main()