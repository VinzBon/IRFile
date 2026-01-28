import pandas as pd
from pathlib import Path
from io import BytesIO
import streamlit as st

def read_csv_absorbance(file_obj) -> pd.Series:
    """
    Reads one CSV (same template):
    - skips first 2 rows
    - extracts 2nd column (absorbance)
    - handles comma decimals
    - converts to numeric
    Tries separators: comma, semicolon, tab.
    """
    def read_with_sep(f, sep: str) -> pd.DataFrame:
        f.seek(0)
        return pd.read_csv(f, skiprows=2, header=None, sep=sep)

    df = None
    for sep in [",", ";", "\t"]:
        try:
            temp = read_with_sep(file_obj, sep)
            if temp.shape[1] >= 2:
                df = temp
                break
        except Exception:
            pass

    if df is None:
        raise ValueError("Could not parse with separators ',', ';', or '\\t'.")

    absorbance = df.iloc[:, 1]
    absorbance = absorbance.astype(str).str.replace(",", ".", regex=False)
    absorbance = pd.to_numeric(absorbance, errors="coerce")

    if absorbance.isna().any():
        bad = int(absorbance.isna().sum())
        raise ValueError(f"{bad} absorbance values could not be parsed as numbers.")

    return absorbance

def read_csv_x_header(file_obj) -> list:
    """
    Reads the first column (after skipping first 2 rows) and returns it as a list
    to be used as Excel headers for the absorbance values.
    Assumes all CSV files share identical first-column values.
    """
    def read_with_sep(f, sep: str) -> pd.DataFrame:
        f.seek(0)
        return pd.read_csv(f, skiprows=2, header=None, sep=sep)

    df = None
    for sep in [",", ";", "\t"]:
        try:
            temp = read_with_sep(file_obj, sep)
            if temp.shape[1] >= 2:
                df = temp
                break
        except Exception:
            pass

    if df is None:
        raise ValueError("Could not parse with separators ',', ';', or '\\t'.")

    x = df.iloc[:, 0]
    return x.astype(str).tolist()

def dataframe_to_excel_bytes(df: pd.DataFrame) -> bytes:
    """Write DataFrame to an in-memory Excel file and return bytes."""
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    buffer.seek(0)
    return buffer.getvalue()

def normalize_xlsx_name(name: str) -> str:
    """Ensure a safe .xlsx filename."""
    name = (name or "").strip()
    if not name:
        return "output.xlsx"
    if not name.lower().endswith(".xlsx"):
        name += ".xlsx"
    return name

def main():
    st.title("CSV Absorbance Transposer")

    st.write(
        "Upload one or more CSV files (same template). "
        "The app will extract the 2nd column (absorbance), transpose it, "
        "and save all results into one Excel file (one row per CSV)."
    )

    output_name = st.text_input("Output Excel file name", value="transposed_absorbance.xlsx")
    output_name = normalize_xlsx_name(output_name)

    uploaded_files = st.file_uploader(
        "Select CSV files",
        type=["csv"],
        accept_multiple_files=True
    )

    process = st.button("Process")

    if process:
        if not uploaded_files:
            st.warning("No files selected.")
            return

        try:
            # Build headers from the first CSV's first column
            x_headers = read_csv_x_header(uploaded_files[0])

            rows = []
            expected_len = None

            for f in uploaded_files:
                absorbance = read_csv_absorbance(f)

                if expected_len is None:
                    expected_len = len(absorbance)
                    if len(x_headers) != expected_len:
                        raise ValueError(
                            f"Header length mismatch from first column: {len(x_headers)} values, "
                            f"but absorbance has {expected_len} points."
                        )
                elif len(absorbance) != expected_len:
                    raise ValueError(
                        f"Length mismatch: '{f.name}' has {len(absorbance)} points, expected {expected_len}."
                    )

                rows.append([Path(f.name).name] + absorbance.tolist())

            columns = ["File"] + x_headers
            out_df = pd.DataFrame(rows, columns=columns)

            st.success("Done! Preview below and download the Excel file.")

            # Preview (show first rows and a limited number of columns to keep UI responsive)
            max_cols = min(12, out_df.shape[1])
            st.dataframe(out_df.iloc[:, :max_cols])

            excel_bytes = dataframe_to_excel_bytes(out_df)

            st.download_button(
                label="Download Excel",
                data=excel_bytes,
                file_name=output_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        except Exception as e:
            st.error(f"Error: {e}")

if __name__ == "__main__":
    main()
