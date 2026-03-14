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

    def test_sanitize_sheet_store_normalizes_entries(self) -> None:
        raw_store = {
            " Panel A ": {
                "header": {"substation_name": "North Yard", "panel_number": "P1"},
                "elements": [{"Element": "OC/EF", "Relay_Details": "P122", "Test_Requirements": "3 points"}],
            },
            "": {"header": {}, "elements": []},
        }

        sanitized = app.sanitize_relay_test_sheet_store(raw_store)

        self.assertIn("Panel A", sanitized)
        self.assertNotIn("", sanitized)
        self.assertEqual(sanitized["Panel A"]["header"]["substation_name"], "North Yard")
        self.assertEqual(sanitized["Panel A"]["elements"][0]["Element"], "OC/EF")

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
                    "Element": "OC/EF",
                    "Relay_Details": "Model P122, Pickup 1.0 A sec",
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
        self.assertIn("Element: OC/EF", text_output)
        self.assertIn("Left Column - Relay Details", text_output)
        self.assertIn("Right Column - Test Requirements", text_output)


if __name__ == "__main__":
    unittest.main()
