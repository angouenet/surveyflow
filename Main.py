import re
from io import BytesIO

import pandas as pd
import streamlit as st


# ---------- Helpers ----------
def cleanse_series(s: pd.Series, upper: bool = True) -> pd.Series:
    """Approximation of Alteryx Cleanse macro:
    - cast to string where non-null
    - strip
    - collapse internal whitespace
    - uppercase (optional)
    """
    if s is None:
        return s
    # Keep NaNs as NaNs
    out = s.copy()
    mask = out.notna()
    out.loc[mask] = out.loc[mask].astype(str).str.strip()
    out.loc[mask] = out.loc[mask].apply(lambda x: re.sub(r"\s+", " ", x))
    if upper:
        out.loc[mask] = out.loc[mask].str.upper()
    return out


def cleanse_df(df: pd.DataFrame, cols: list[str], upper: bool = True) -> pd.DataFrame:
    df = df.copy()
    for c in cols:
        if c in df.columns:
            df[c] = cleanse_series(df[c], upper=upper)
    return df


def build_qc_output(
    responses_df: pd.DataFrame,
    recon_df: pd.DataFrame,
    numerical_df: pd.DataFrame,
    drop_from_responses: list[str] | None = None,
    responses_key_fields: list[str] | None = None,
    recon_sheet_expected: bool = True,
) -> pd.DataFrame:
    # 1) Cleanse responses (pre-select)
    responses_df = cleanse_df(responses_df, cols=[], upper=False)  # no-op placeholder

    # 2) Select/drop columns like Alteryx Select (optional)
    if drop_from_responses:
        responses_df = responses_df.drop(columns=[c for c in drop_from_responses if c in responses_df.columns], errors="ignore")

    # 3) Transpose (Alteryx Transpose: key fields + Name/Value)
    if responses_key_fields is None:
        responses_key_fields = ["Respondent", "End time (GMT)", "Panel"]

    # Ensure key fields exist; keep only those that exist
    keys = [c for c in responses_key_fields if c in responses_df.columns]
    data_cols = [c for c in responses_df.columns if c not in keys]

    long_df = responses_df.melt(
        id_vars=keys,
        value_vars=data_cols,
        var_name="Name",
        value_name="Value",
    )

    # 4) Filter Value is not null
    long_df = long_df[long_df["Value"].notna()].copy()

    # 5) Cleanse after transpose (like your ToolID 7)
    long_df = cleanse_df(long_df, cols=["Respondent", "Panel", "Name"], upper=True)

    # 6) Recon select + cleanse (like ToolID 62 + 10)
    # (We won't drop recon columns unless you want to; we just cleanse important text cols)
    recon_cols_to_clean = [
        "Name",
        "Question number",
        "Section",
        "Question Text",
        "Answer Option",
        "Loop Variable",
        "Matrix Dimension",
        "Double Loop Var1",
        "Double Loop Var2",
    ]
    recon_df = cleanse_df(recon_df, cols=recon_cols_to_clean, upper=True)

    # 7) Join long responses to recon on Name (inner join like Alteryx Join output)
    joined = long_df.merge(recon_df, on="Name", how="inner", suffixes=("", "_recon"))

    # 8) Numerical recon join on Value
    # Clean numerical "Value" as well to improve matching (optional but usually helpful)
    if "Value" in numerical_df.columns:
        numerical_df = numerical_df.copy()
        numerical_df["Value"] = cleanse_series(numerical_df["Value"], upper=True)

    # Also cleanse joined Value to align with numerical mapping
    joined["Value"] = cleanse_series(joined["Value"], upper=True)

    # Equivalent to Alteryx Join + Left union:
    # In pandas, a LEFT join already keeps unmatched rows; that matches your union result.
    out = joined.merge(numerical_df, on="Value", how="left", suffixes=("", "_num"))

    return out


def to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Sheet1") -> bytes:
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return bio.getvalue()


# ---------- Streamlit UI ----------
st.set_page_config(page_title="Survey QC Builder", layout="wide")
st.title("Survey QC Builder (Alteryx → Python)")

st.write("Upload the three inputs and click **Run** to generate the QC output Excel.")

col1, col2, col3 = st.columns(3)
with col1:
    up_responses = st.file_uploader("1) Responses (.xlsx)", type=["xlsx"])
with col2:
    up_recon = st.file_uploader("2) Recon (.xlsm)", type=["xlsm", "xlsx"])
with col3:
    up_numerical = st.file_uploader("3) Numerical Recon (.xlsx)", type=["xlsx"])

with st.expander("Advanced options"):
    responses_sheet = st.text_input("Responses sheet name", value="Sheet1")
    recon_sheet = st.text_input("Recon sheet name", value="Master Recon")
    numerical_sheet = st.text_input("Numerical recon sheet name", value="Sheet1")

    key_fields = st.text_input("Key fields (comma-separated)", value="Respondent,End time (GMT),Panel")
    drop_cols = st.text_area(
        "Drop columns from Responses (one per line) — optional",
        value="Status\nTerm reason\nStart time (GMT)"
    )
    output_filename = st.text_input("Output filename", value="QC.xlsx")

run = st.button("Run", type="primary", disabled=not (up_responses and up_recon and up_numerical))

if run:
    try:
        # Read uploads
        responses_df = pd.read_excel(up_responses, sheet_name=responses_sheet)
        recon_df = pd.read_excel(up_recon, sheet_name=recon_sheet)
        numerical_df = pd.read_excel(up_numerical, sheet_name=numerical_sheet)

        drop_from_responses = [c.strip() for c in drop_cols.splitlines() if c.strip()]
        responses_key_fields = [c.strip() for c in key_fields.split(",") if c.strip()]

        out_df = build_qc_output(
            responses_df=responses_df,
            recon_df=recon_df,
            numerical_df=numerical_df,
            drop_from_responses=drop_from_responses,
            responses_key_fields=responses_key_fields,
        )

        st.success(f"Done. Output rows: {len(out_df):,} | cols: {len(out_df.columns):,}")

        st.dataframe(out_df.head(50), use_container_width=True)

        xbytes = to_excel_bytes(out_df, sheet_name="Sheet1")
        st.download_button(
            label="Download QC Excel",
            data=xbytes,
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        st.error(f"Run failed: {e}")
        st.exception(e)