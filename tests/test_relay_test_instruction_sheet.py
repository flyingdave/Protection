import contextlib
import io
import unittest

import pandas as pd

with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
    import app


class RelayTestInstructionSheetTests(unittest.TestCase):
    def test_default_elements_include_required_options(self) -> None:
        default_df = app.default_relay_test_elements_df()
        self.assertListEqual(default_df["Element"].tolist(), app.RELAY_TEST_ELEMENT_OPTIONS)
        self.assertListEqual(list(default_df.columns), app.RELAY_TEST_ELEMENT_COLUMNS)
        self.assertTrue((default_df["Relay_Group"] == app.RELAY_TEST_DEFAULT_GROUP_NAME).all())

    def test_sanitize_sheet_store_normalizes_entries(self) -> None:
        raw_store = {
            " Panel A ": {
                "header": {"substation_name": "North Yard", "panel_number": "P1"},
                "layout_style": "ECNSW Compact",
                "elements": [
                    {
                        "Relay_Group": "Relay A",
                        "Element": "OC/EF",
                        "Settings": "P122",
                        "Test_Requirements": "3 points",
                    }
                ],
            },
            "": {"header": {}, "elements": []},
        }

        sanitized = app.sanitize_relay_test_sheet_store(raw_store)

        self.assertIn("Panel A", sanitized)
        self.assertNotIn("", sanitized)
        self.assertEqual(sanitized["Panel A"]["header"]["substation_name"], "North Yard")
        self.assertEqual(sanitized["Panel A"]["layout_style"], "ECNSW Compact")
        self.assertEqual(sanitized["Panel A"]["elements"][0]["Relay_Group"], "Relay A")
        self.assertEqual(sanitized["Panel A"]["elements"][0]["Element"], "OC/EF")

    def test_invalid_layout_style_defaults_to_grouped_detail(self) -> None:
        sanitized = app.sanitize_relay_test_sheet_store(
            {
                "Panel B": {
                    "header": {},
                    "layout_style": "Unknown Layout",
                    "elements": [],
                }
            }
        )

        self.assertEqual(sanitized["Panel B"]["layout_style"], app.RELAY_TEST_DEFAULT_LAYOUT)

    def test_legacy_relay_details_column_is_mapped_to_settings(self) -> None:
        legacy_df = pd.DataFrame(
            [
                {
                    "Element": "OC/EF",
                    "Relay_Details": "Legacy details",
                    "Test_Requirements": "Legacy test requirements",
                }
            ]
        )

        sanitized = app.sanitize_relay_test_elements(legacy_df)

        self.assertEqual(sanitized.loc[0, "Relay_Group"], app.RELAY_TEST_DEFAULT_GROUP_NAME)
        self.assertEqual(sanitized.loc[0, "Settings"], "Legacy details")
        self.assertEqual(sanitized.loc[0, "Test_Requirements"], "Legacy test requirements")

    def test_build_instruction_text_contains_header_and_element_sections(self) -> None:
        header_data = app.sanitize_relay_test_header(
            {
                "substation_name": "Alpha Substation",
                "panel_number": "Panel-07",
                "feeder_number": "Feeder-3",
                "test_engineer": "A. Engineer",
            }
        )
        elements_df = pd.DataFrame(
            [
                {
                    "Relay_Group": "Relay A",
                    "Element": "OC/EF",
                    "Settings": "Model P122, Pickup 1.0 A sec",
                    "Test_Requirements": "Inject 1.5x, 3x, 6x pickup",
                }
            ]
        )

        text_output = app.build_relay_test_instruction_text(
            panel_id="Panel-07",
            header_data=header_data,
            elements_df=elements_df,
        )

        self.assertIn("RELAY TEST INSTRUCTION SHEET", text_output)
        self.assertIn("Panel Sheet ID: Panel-07", text_output)
        self.assertIn("Substation: Alpha Substation", text_output)
        self.assertIn("Relay Group: Relay A", text_output)
        self.assertIn("Element: OC/EF", text_output)
        self.assertIn("Column 1 - Settings", text_output)
        self.assertIn("Column 2 - Test Requirements", text_output)

    def test_build_instruction_text_ecnsw_compact_layout(self) -> None:
        header_data = app.sanitize_relay_test_header({"substation_name": "Alpha Substation"})
        elements_df = pd.DataFrame(
            [
                {
                    "Relay_Group": "Relay A",
                    "Element": "Translay",
                    "Settings": "Model: H04",
                    "Test_Requirements": "Check operate timing",
                }
            ]
        )

        text_output = app.build_relay_test_instruction_text(
            panel_id="Panel-08",
            header_data=header_data,
            elements_df=elements_df,
            layout_style="ECNSW Compact",
        )

        self.assertIn("Layout Style: ECNSW Compact", text_output)
        self.assertIn("Element                  | Settings                 | Test requirements", text_output)
        self.assertIn("Translay", text_output)
        self.assertIn("Model: H04", text_output)


if __name__ == "__main__":
    unittest.main()
