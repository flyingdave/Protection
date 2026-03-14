import contextlib
import io
import math
import unittest

with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
    import app


class NetworkFaultCalculationTests(unittest.TestCase):
    @staticmethod
    def _compute_network_faults(overrides: dict | None = None) -> dict:
        params = app.DEFAULT_NETWORK_PRESET.copy()
        if overrides:
            params.update(overrides)

        v_ll_kv = float(params["v_ll_kv"])
        source_v_kv = float(params["source_v_kv"])
        source_sc_mva = float(params["source_sc_mva"])
        source_xr = float(params["source_xr"])
        reactor_ohm = float(params["reactor_ohm"])
        hv_cable_length_km = float(params["hv_cable_length_km"])
        hv_cable_type = str(params["hv_cable_type"])
        transformer_lv_kv = float(params["transformer_lv_kv"])
        transformer_mva = float(params["transformer_mva"])
        transformer_z_pct = float(params["transformer_z_pct"])
        transformer_xr = float(params["transformer_xr"])
        lv_cable_length_km = float(params["lv_cable_length_km"])
        lv_cable_type = str(params["lv_cable_type"])
        feeder_length_km = float(params["feeder_length_km"])
        feeder_cable_type = str(params["feeder_cable_type"])

        duplicate_source_to_first_busbar_circuit = bool(
            params.get("duplicate_source_to_first_busbar_circuit", False)
        )
        parallel_lv_transformer_cable = bool(params.get("parallel_lv_transformer_cable", False))
        parallel_feeder_cable = bool(params.get("parallel_feeder_cable", False))

        source_z_hv = app.split_rx_from_z_magnitude((source_v_kv**2) / source_sc_mva, source_xr)
        source_to_study_factor = (v_ll_kv / source_v_kv) ** 2 if source_v_kv > 0 else 0.0
        source_z1 = source_z_hv * source_to_study_factor
        source_z2 = source_z1
        source_z0 = source_z1 * app.SOURCE_Z0_FACTOR

        reactor_z1 = complex(0.0, reactor_ohm) * source_to_study_factor
        reactor_z2 = reactor_z1
        reactor_z0 = reactor_z1

        hv_cable_data = app.cable_sequence_impedance_per_km(hv_cable_type)
        hv_cable_z1 = hv_cable_length_km * hv_cable_data["z1"] * source_to_study_factor
        hv_cable_z2 = hv_cable_length_km * hv_cable_data["z2"] * source_to_study_factor
        hv_cable_z0 = hv_cable_length_km * hv_cable_data["z0"] * source_to_study_factor

        transformer_z_lv = app.split_rx_from_z_magnitude(
            (transformer_z_pct / 100) * ((transformer_lv_kv**2) / transformer_mva),
            transformer_xr,
        )
        lv_to_study_factor = (v_ll_kv / transformer_lv_kv) ** 2 if transformer_lv_kv > 0 else 0.0
        transformer_z1 = transformer_z_lv * lv_to_study_factor
        transformer_z2 = transformer_z1
        transformer_z0 = transformer_z1 * app.TRANSFORMER_Z0_FACTOR

        lv_cable_data = app.cable_sequence_impedance_per_km(lv_cable_type)
        lv_cable_z1 = lv_cable_length_km * lv_cable_data["z1"] * lv_to_study_factor
        lv_cable_z2 = lv_cable_length_km * lv_cable_data["z2"] * lv_to_study_factor
        lv_cable_z0 = lv_cable_length_km * lv_cable_data["z0"] * lv_to_study_factor

        lv_transformer_cable_count = 2 if parallel_lv_transformer_cable else 1
        lv_cable_equiv_z1 = lv_cable_z1 / lv_transformer_cable_count
        lv_cable_equiv_z2 = lv_cable_z2 / lv_transformer_cable_count
        lv_cable_equiv_z0 = lv_cable_z0 / lv_transformer_cable_count

        single_source_to_first_busbar_circuit_z1 = (
            reactor_z1 + hv_cable_z1 + transformer_z1 + lv_cable_equiv_z1
        )
        single_source_to_first_busbar_circuit_z2 = (
            reactor_z2 + hv_cable_z2 + transformer_z2 + lv_cable_equiv_z2
        )
        single_source_to_first_busbar_circuit_z0 = (
            reactor_z0 + hv_cable_z0 + transformer_z0 + lv_cable_equiv_z0
        )

        parallel_circuit_count = 2 if duplicate_source_to_first_busbar_circuit else 1
        source_to_first_busbar_equiv_z1 = single_source_to_first_busbar_circuit_z1 / parallel_circuit_count
        source_to_first_busbar_equiv_z2 = single_source_to_first_busbar_circuit_z2 / parallel_circuit_count
        source_to_first_busbar_equiv_z0 = single_source_to_first_busbar_circuit_z0 / parallel_circuit_count

        feeder_cable_data = app.cable_sequence_impedance_per_km(feeder_cable_type)
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

        incomer_fault = app.calc_fault_values(v_ll_kv, z1_incomer, z2_incomer, z0_incomer)
        load_fault = app.calc_fault_values(v_ll_kv, z1_load, z2_load, z0_load)

        return {
            "z1_incomer": abs(z1_incomer),
            "z1_load": abs(z1_load),
            "incomer_fault": incomer_fault,
            "load_fault": load_fault,
        }

    def test_calc_fault_values_reference_case(self) -> None:
        z1 = complex(0.5, 1.0)
        z2 = complex(0.5, 1.0)
        z0 = complex(1.5, 3.0)
        result = app.calc_fault_values(11.0, z1, z2, z0)

        expected_i3 = 11.0 / (math.sqrt(3) * abs(z1))
        expected_ill = 11.0 / abs(z1 + z2)
        expected_ilg = math.sqrt(3) * 11.0 / abs(z1 + z2 + z0)
        expected_fault_mva = math.sqrt(3) * 11.0 * expected_i3

        self.assertAlmostEqual(result["I_3ph_kA"], expected_i3, places=9)
        self.assertAlmostEqual(result["I_LL_kA"], expected_ill, places=9)
        self.assertAlmostEqual(result["I_LG_kA"], expected_ilg, places=9)
        self.assertAlmostEqual(result["Fault_MVA"], expected_fault_mva, places=9)

    def test_default_network_has_lower_remote_fault_level(self) -> None:
        computed = self._compute_network_faults()
        incomer_fault = computed["incomer_fault"]
        load_fault = computed["load_fault"]

        self.assertGreater(incomer_fault["I_3ph_kA"], load_fault["I_3ph_kA"])
        self.assertGreater(incomer_fault["I_LL_kA"], load_fault["I_LL_kA"])
        self.assertGreater(incomer_fault["I_LG_kA"], load_fault["I_LG_kA"])

    def test_default_network_golden_baseline_fault_levels(self) -> None:
        computed = self._compute_network_faults()
        incomer_fault = computed["incomer_fault"]
        load_fault = computed["load_fault"]

        expected_incomer = {
            "Z1_total_ohm": 1.1075599852793814,
            "Z2_total_ohm": 1.1075599852793814,
            "Z0_total_ohm": 1.3229924325235352,
            "I_3ph_kA": 5.734093905066356,
            "I_LL_kA": 4.965870989472979,
            "I_LG_kA": 5.38816982725064,
            "Fault_MVA": 109.24916176840554,
        }
        expected_load = {
            "Z1_total_ohm": 1.1638046541067235,
            "Z2_total_ohm": 1.1638046541067235,
            "Z0_total_ohm": 1.5097562166473846,
            "I_3ph_kA": 5.456975050473973,
            "I_LL_kA": 4.725879021528331,
            "I_LG_kA": 4.974932699157874,
            "Fault_MVA": 103.96933847362327,
        }

        for key, expected_value in expected_incomer.items():
            self.assertAlmostEqual(incomer_fault[key], expected_value, places=6)

        for key, expected_value in expected_load.items():
            self.assertAlmostEqual(load_fault[key], expected_value, places=6)

    def test_higher_source_sc_mva_increases_fault_current(self) -> None:
        low_source = self._compute_network_faults({"source_sc_mva": 1200.0})
        high_source = self._compute_network_faults({"source_sc_mva": 3000.0})

        self.assertGreater(
            high_source["incomer_fault"]["I_3ph_kA"],
            low_source["incomer_fault"]["I_3ph_kA"],
        )

    def test_parallel_source_path_increases_incomer_fault(self) -> None:
        base = self._compute_network_faults({"duplicate_source_to_first_busbar_circuit": False})
        duplicated = self._compute_network_faults({"duplicate_source_to_first_busbar_circuit": True})

        self.assertLess(duplicated["z1_incomer"], base["z1_incomer"])
        self.assertGreater(
            duplicated["incomer_fault"]["I_3ph_kA"],
            base["incomer_fault"]["I_3ph_kA"],
        )

    def test_parallel_11kv_cables_behave_correctly(self) -> None:
        base = self._compute_network_faults()
        transformer_parallel = self._compute_network_faults({"parallel_lv_transformer_cable": True})
        feeder_parallel = self._compute_network_faults({"parallel_feeder_cable": True})

        self.assertGreater(
            transformer_parallel["incomer_fault"]["I_3ph_kA"],
            base["incomer_fault"]["I_3ph_kA"],
        )
        self.assertAlmostEqual(
            feeder_parallel["incomer_fault"]["I_3ph_kA"],
            base["incomer_fault"]["I_3ph_kA"],
            places=9,
        )
        self.assertGreater(
            feeder_parallel["load_fault"]["I_3ph_kA"],
            base["load_fault"]["I_3ph_kA"],
        )

    def test_default_line_builder_matches_template_fault_levels(self) -> None:
        builder_results = app.calculate_line_builder_network(app.default_network_elements_df())
        builder_fault_df = builder_results["fault_df"]
        template_results = self._compute_network_faults()

        transformer_bus_fault = (
            builder_fault_df.loc[builder_fault_df["Bus"] == "11kV Transformer Busbar"].iloc[0].to_dict()
        )
        remote_bus_fault = (
            builder_fault_df.loc[builder_fault_df["Bus"] == "11kV Remote Busbar"].iloc[0].to_dict()
        )

        for key, expected_value in template_results["incomer_fault"].items():
            self.assertAlmostEqual(transformer_bus_fault[key], expected_value, places=6)

        for key, expected_value in template_results["load_fault"].items():
            self.assertAlmostEqual(remote_bus_fault[key], expected_value, places=6)

    def test_line_builder_busbars_define_fault_locations(self) -> None:
        network_df = app.default_network_elements_df().copy()
        network_df.loc[network_df["Name"] == "11kV Transformer Busbar", "Name"] = "Main Switchboard"
        network_df.loc[network_df["Name"] == "11kV Remote Busbar", "Name"] = "Remote Load Bus"

        builder_results = app.calculate_line_builder_network(network_df)
        fault_buses = builder_results["fault_df"]["Bus"].tolist()

        self.assertEqual(fault_buses, ["Main Switchboard", "Remote Load Bus"])
        self.assertIn("Main Switchboard", builder_results["diagram_dot"])
        self.assertIn("Remote Load Bus", builder_results["diagram_dot"])

    def test_line_builder_parallel_feeder_increases_remote_fault_level(self) -> None:
        base_results = app.calculate_line_builder_network(app.default_network_elements_df())
        parallel_df = app.default_network_elements_df().copy()
        parallel_df.loc[parallel_df["Name"] == "11kV Feeder Cable", "Parallel_Count"] = 2
        parallel_results = app.calculate_line_builder_network(parallel_df)

        base_faults = base_results["fault_df"].set_index("Bus")
        parallel_faults = parallel_results["fault_df"].set_index("Bus")

        self.assertAlmostEqual(
            parallel_faults.loc["11kV Transformer Busbar", "I_3ph_kA"],
            base_faults.loc["11kV Transformer Busbar", "I_3ph_kA"],
            places=9,
        )
        self.assertGreater(
            parallel_faults.loc["11kV Remote Busbar", "I_3ph_kA"],
            base_faults.loc["11kV Remote Busbar", "I_3ph_kA"],
        )


if __name__ == "__main__":
    unittest.main()
