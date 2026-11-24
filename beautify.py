
import argparse
import sys
from pathlib import Path
from copy import copy

import pandas as pd
from openpyxl.utils import get_column_letter


def beautify(input_file, output_file=None):
    """
    Transform the 'result1' sheet into the required Excel format.

    Args:
        input_file: Path to the input Excel file.
        output_file: Optional path for the output Excel file.

    Returns:
        str: Path to the generated Excel file.

    Raises:
        FileNotFoundError: If the input file does not exist.
        ValueError: If required sheets or columns are missing,
                    or numeric data is invalid.
        RuntimeError: If Excel transformation fails.
    """

    input_path = Path(input_file)

    # Validate input file
    if not input_path.is_file():
        raise FileNotFoundError(
            f"Input Excel file not found: {input_file}"
        )

    # Generate default output filename
    if output_file is None:
        output_path = input_path.with_name(
            f"beautified_{input_path.stem}.xlsx"
        )
    else:
        output_path = Path(output_file)

    # Ensure output has .xlsx extension
    if output_path.suffix.lower() != ".xlsx":
        raise ValueError("Output file must have a .xlsx extension.")

    # Prevent overwriting the original file
    if input_path.resolve() == output_path.resolve():
        raise ValueError(
            "Input and output files cannot be the same."
        )

    try:
        # --------------------------------------------------
        # 1. Read Excel file
        # --------------------------------------------------

        workbook = pd.ExcelFile(
            input_path,
            engine="openpyxl"
        )

        if "result1" not in workbook.sheet_names:
            raise ValueError(
                "Sheet 'result1' not found in input Excel file."
            )

        df = pd.read_excel(
            workbook,
            sheet_name="result1"
        )

        workbook.close()

        # --------------------------------------------------
        # 2. Define column mapping
        # --------------------------------------------------

        column_mapping = {
            "dfar_ismt_idnt": "Security Description",

            "dfar_ismt_scnd_desc": "ISIN",

            "dfar_totl_util_qnty":
                "Indicative Value Of Aggregate Holding Of FPIS (INR Cr)",

            "dfar_ismt_issu_qnty":
                "Outstanding Position Of Govt Securities (INR Cr)",

            "dfar_outx_stdg_pect": "Sec Holdings (%)"
        }

        # --------------------------------------------------
        # 3. Validate required columns
        # --------------------------------------------------

        missing_columns = [
            col for col in column_mapping
            if col not in df.columns
        ]

        if missing_columns:
            raise ValueError(
                f"Missing required columns: {missing_columns}"
            )

        # --------------------------------------------------
        # 4. Select and rename columns
        # --------------------------------------------------

        # Selecting only required columns automatically
        # removes dfar_id and any other unwanted columns.

        df = df[list(column_mapping.keys())].copy()

        df.rename(
            columns=column_mapping,
            inplace=True
        )

        # --------------------------------------------------
        # 5. Arrange columns in required order
        # --------------------------------------------------

        output_columns = [
            "ISIN",
            "Security Description",
            "Indicative Value Of Aggregate Holding Of FPIS (INR Cr)",
            "Outstanding Position Of Govt Securities (INR Cr)",
            "Sec Holdings (%)"
        ]

        df = df[output_columns]

        # --------------------------------------------------
        # 6. Convert and validate numeric columns
        # --------------------------------------------------

        numeric_columns = output_columns[2:]

        for col in numeric_columns:

            original_values = df[col]

            converted_values = pd.to_numeric(
                original_values,
                errors="coerce"
            )

            # Detect invalid non-empty numeric values
            invalid_mask = (
                original_values.notna()
                & original_values.astype(str).str.strip().ne("")
                & converted_values.isna()
            )

            if invalid_mask.any():
                invalid_values = (
                    original_values[invalid_mask]
                    .head(5)
                    .tolist()
                )

                raise ValueError(
                    f"Invalid numeric values in '{col}': "
                    f"{invalid_values}"
                )

            df[col] = converted_values

        # --------------------------------------------------
        # 7. Calculate Total row
        # --------------------------------------------------

        total_indicative = df[numeric_columns[0]].sum(
            min_count=1
        )

        total_outstanding = df[numeric_columns[1]].sum(
            min_count=1
        )

        if (
            pd.notna(total_indicative)
            and pd.notna(total_outstanding)
            and total_outstanding != 0
        ):
            total_percentage = (
                total_indicative / total_outstanding
            ) * 100
        else:
            total_percentage = 0.0

        # --------------------------------------------------
        # 8. Append Total row
        # --------------------------------------------------

        total_row = {
            "ISIN": "",
            "Security Description": "Total",

            numeric_columns[0]: total_indicative,

            numeric_columns[1]: total_outstanding,

            numeric_columns[2]: total_percentage
        }

        df = pd.concat(
            [df, pd.DataFrame([total_row])],
            ignore_index=True
        )

        # --------------------------------------------------
        # 9. Export to Excel
        # --------------------------------------------------

        with pd.ExcelWriter(
            output_path,
            engine="openpyxl"
        ) as writer:

            df.to_excel(
                writer,
                sheet_name="result1",
                index=False
            )

            ws = writer.sheets["result1"]

            # Freeze header row
            ws.freeze_panes = "A2"

            # --------------------------------------------------
            # 10. Format header row (BOLD)
            # --------------------------------------------------

            for cell in ws[1]:

                font = copy(cell.font)
                font.bold = True
                cell.font = font

            # --------------------------------------------------
            # 11. Format Total row (BOLD)
            # --------------------------------------------------

            for cell in ws[ws.max_row]:

                font = copy(cell.font)
                font.bold = True
                cell.font = font

            # --------------------------------------------------
            # 12. Adjust column widths
            # --------------------------------------------------

            widths = [
                18,  # ISIN
                25,  # Security Description
                48,  # Indicative Value
                48,  # Outstanding Position
                18   # Sec Holdings
            ]

            for i, width in enumerate(widths, start=1):

                column_letter = get_column_letter(i)

                ws.column_dimensions[
                    column_letter
                ].width = width

    except Exception as e:

        raise RuntimeError(
            f"Beautification failed: {e}"
        ) from e

    print("Excel transformation completed successfully!")
    print(f"Output file: {output_path}")

    return str(output_path)


# --------------------------------------------------
# Standalone execution
# --------------------------------------------------

if __name__ == "__main__":

    parser = argparse.ArgumentParser(
        description="Beautify the result1 sheet in an Excel file."
    )

    parser.add_argument(
        "input_file",
        help="Input Excel file path"
    )

    parser.add_argument(
        "-o",
        "--output",
        help="Output Excel file path (optional)"
    )

    args = parser.parse_args()

    try:
        beautify(
            args.input_file,
            args.output
        )

    except Exception as e:
        print(f"ERROR: {e}")
        sys.exit(1)
