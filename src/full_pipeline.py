"""


In this script the full pipeline will be implemented and exeption will be handled 

======================================================================

what do the proccess looks like and what do all the script do 

-1-2 put the whole file inside input folder 

=============the script to run in order ==================

-3 sheet_extractor will run and put all sheet in there respective folder 
input - "input"
output - in extracted  and data_base_extracted


-4 master_clean.py     will clean the the input in extracted and put them to cleaned 
-5 dict_creator.py will run and create a dictionary and put it in dict in excel format 
-6 map.py will be excuted and create input from data_base_cleaned and cleaned folder then create the data base excel file with woreda in them 

input dict_creator.py
input cleaned and data_base_cleaned
output woreda_added




-7 finished now automate this in this folder 




"""
from pathlib import Path

#=====================================================================

from src.sheet_extracter import extract

print("&"*200)

print("sheet extractor started")
print("&"*200)


base = Path(__file__).resolve().parents[1]
c = base /"data"/"input"

list_of_files = [f for f in c.glob("*.xlsx") ]

total_no_sheet = 0 
for file in list_of_files:
    no_of_sheet = extract(file)
    total_no_sheet = total_no_sheet + no_of_sheet 
print("="*200)

print(f"Number of files found to be extracted is (({len(list_of_files)})) and extracted (({total_no_sheet})) ")
print("--------------extracting part is done")


print("&"*200)
print("sheet extractor ended")
print("&"*200)

#============================================================

from src.master_clean import clean_all

print("="*80)
print("master cleaner started")
print("="*80)

BASE = Path(__file__).resolve().parents[1]

path = BASE / "data" / "extracted"
output_path = BASE / "data" / "cleaned"

list_of_files = [file for file in path.glob("*.xlsx")]



for file in list_of_files:
    clean = clean_all()
    clean.clean_sheet(sheet_path = file , output_path=output_path)

print("="*80)
print(f"-----------------------------------------------------------cleaned---------------------------------------------- (({len(list_of_files)}))")
print("master cleaner ended")
print("="*80)

#==========================================================




from src.dict_creator import run

print("="*80)
print("dictionary creator started ")
print("="*80)


BASE =  Path(__file__).resolve().parents[1]
path = BASE / "data" / "cleaned"
output_path = BASE / "data" / "dict" / "dict.xlsx"

list_of_files = [f for f in path.glob("*.xlsx")]

dictionary = run(list_of_files)
dictionary.to_excel(output_path , index = False , engine= "openpyxl")
print(dictionary)

print("="*80)
print("dictionary creator ended ")
print("="*80)


#================================================================



from src.map import add_woreda

print("="*80)
print("woreda adder  started ")
print("="*80)


BASE = Path(__file__).resolve().parents[1]
path = BASE / "data" / "cleaned" 
database_path = BASE / "data" / "data_base_cleaned"
output_path = BASE / "data" / "woreda_added"

list_of_path = [file for file in path.glob("*.xlsx")]
list_of_path_for_database = [f for f in database_path.glob("*.xlsx")]


for path in list_of_path_for_database:
    print(f"adding woreda for (({path.name}))")
    new_path = output_path / Path(path).name
    add_woreda(path , new_path , dictionary)


print("="*80)
print("--------------------------woreda adder  ended -----------------")
print("="*80)

