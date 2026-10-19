import io
from pathlib import Path
import pandas as pd
import plotly.express as px
import streamlit as st



base = Path(__file__).resolve().parents[1]   


# names shown in the sidebar  ->  (folder path, also search sub-folders?)
DATASETS = {
    "Woreda added":         (base / "data" / "woreda_added", True),
    "Database cleaned":     (base / "data" / "cleaned" / "data_base_files_cleaned", True),
    "Sent files extracted": (base   / "data" / "extracted" / "sent_files", False),
    "Sent files cleaned":   (base / "data" / "sent_files_cleaned", False),
}

file_pattern = "*.xlsx"
header_row  = 0
folder_name = "Folder"   # name of the column built from the folder name
date = "MS"          # "D" day, "W" week, "MS" month, "QS" quarter, "YS" year

st.set_page_config(page_title="Farmers Data Report Dashboard", layout="wide")
st.title("Report Dashboard")


# loader -------------------------------------


@st.cache_data(show_spinner="Reading Excel files...")
def load_folder(root: str, recursive: bool) -> pd.DataFrame:
    root = Path(root)
    paths = root.glob(file_pattern)

    frames = []
    for path in paths:
        if path.name.startswith("~$"):   # skip Excel temp files
            continue
        try:
            sheets = pd.read_excel(path, sheet_name=None, header=header_row)
        except Exception as e:
            st.warning(f"Skipped {path.name}: {e}")
            continue
        for sheet_name, sheet_df in sheets.items():
            sheet_df = sheet_df.copy()
            sheet_df[folder_name] = path.parent.name
            sheet_df["source_file"] = path.name
            sheet_df["source_sheet"] = sheet_name
            frames.append(sheet_df)

    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True)



st.sidebar.header("Dataset")
choice = st.sidebar.radio("Choose data", list(DATASETS.keys()))
folder_path, recursive = DATASETS[choice]

st.sidebar.caption(f"Reading from: {folder_path}")
if st.sidebar.button("Reload files"):
    st.cache_data.clear()

if not folder_path.exists():
    st.error(f"Folder not found: {folder_path}")
    st.stop()


df = load_folder(str(folder_path), recursive)

if choice in ("Woreda added", "Database cleaned"):
    df = df.drop_duplicates(subset=["id"])



if df.empty:
    st.warning(f"No {file_pattern} files found in {folder_path}")
    st.stop()


# filter 
st.sidebar.header("Filters")

folders = sorted(df[folder_name].unique())
chosen_folders = st.sidebar.multiselect(folder_name, folders, default=folders)
df = df[df[folder_name].isin(chosen_folders)]


files = sorted(df["source_file"].unique())
chosen_files = st.sidebar.multiselect("File", files, default=files)
df = df[df["source_file"].isin(chosen_files)]



# kpi 

num_cols = df.select_dtypes("number").columns.tolist()

c1, c2, c3 = st.columns(3)
c1.metric("Rows", f"{len(df):,}")
c2.metric("Files", df["source_file"].nunique())
c3.metric(folder_name + "s", df[folder_name].nunique())


# ch 
st.subheader("Explore")

text_cols = [c for c in df.columns
             if df[c].dtype == "object"
             and c not in (folder_name, "source_file", "source_sheet")]
cat_cols = [folder_name, "source_file"] + text_cols

if num_cols:
    left, right = st.columns(2)
    with left:
        x = st.selectbox("Group by", cat_cols)
        y = st.selectbox("Value", num_cols)
        agg = st.selectbox("Aggregation", ["sum", "mean", "count", "max", "min"])
        grouped = df.groupby(x, as_index=False)[y].agg(agg)
        st.plotly_chart(px.bar(grouped, x=x, y=y, title=f"{agg} of {y} by {x}"),
                        use_container_width=True)
    with right:
        share = df.groupby(folder_name , as_index=False)[y].agg(agg)
        st.plotly_chart(px.pie(share, names=folder_name, values=y,
                               title=f"{y} share by {folder_name}"),
                        use_container_width=True)


else:
    st.info("No numeric columns in this dataset, so no charts. Check the table below.")



st.subheader("Data")
st.dataframe(df, use_container_width=True)

buf = io.BytesIO()
df.to_excel(buf, index=False)
st.download_button("Download this data as Excel", buf.getvalue(),
                   file_name="combined_report.xlsx")