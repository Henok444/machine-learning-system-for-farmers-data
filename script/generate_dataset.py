from pathlib import Path
import pandas as pd 
import openpyxl



base = Path(__file__).resolve().parents[1]  

path = base / "data" / "cleaned" / "data_base_files_cleaned"
output_path = base / "data" / "data_for_analaysis" 
name = "full_dataset.xlsx"
Path(output_path).mkdir(parents=True , exist_ok= True)

file_pattern = "*.xlsx"
header_row  = 0
folder_name = "Folder"   # name of the column built from the folder name
date = "MS" 


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



df = load_folder(path , True)

#df.to_excel(output_path / name  , index = False , engine = "openpyxl")
df.to_csv(output_path / "full_dataset.csv" , index= False)
