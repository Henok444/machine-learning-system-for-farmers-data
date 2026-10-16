
import pandas as pd 
from pathlib import Path 
import re 






"""

goal ---> input many sheets and many files ----->> output many files with single sheet 

simpler version ----> input single file with multiple sheet ---->>  output manyfile with single shit 


psedo 

first it reads one file from path  and set sheet_name to none , 

for each sheet 
save the name of one sheet 

create a a single idependendent file containing only that one sheet and name it the sheet name and save to the output location 

count the no of sheet there a
report how many you created last

  -----------------done 

now lets expand to do multiple files at once

first create a list of file found in the folder 
then iterate the function accross the file 

---------------------done 

  

=====================================================
tools i needed and created when applying the concept 

first a tool that remove white space and insert _ between the strings its basicly name cleaner and also add .xlsx format at the end 
clean_name()

psdo 
first clean 

recive a string 
replace single doble and tripple spaces by _

then add 

append .xlsx to the last of the string 



"""
# helper tool // clean 

def clean_name(name : str )->str:

   # clean 
   new_name =  re.sub(r"\s+" , "_" , name )
   # add 
   new_name = f"{new_name}.xlsx"

   return new_name



# read one file from path 
def extract(file: Path )-> None :

   BASE  = Path(__file__).resolve().parents[1]
   input_path = file 
   multi_sheet = pd.read_excel(input_path , sheet_name= None , engine= "openpyxl")
   output_path = BASE /"data"/ "data_base_cleaned"


   # if there is only one sheet , we want it to have the name it has before
   print(len([file for file , df in multi_sheet.items()]))
   if len([file for file , df in multi_sheet.items()]) <= 1:
      path = output_path / Path(input_path).name
      sheet = pd.read_excel(input_path  , engine= "openpyxl")
      sheet.to_excel(path , engine = "openpyxl")
      print("breaked")
      return None

   

   # for each sheet save the name of the sheet 
   i = 0 
   names = []
   for sheet_name , df in multi_sheet.items(): # tested and passed 
      name = clean_name(sheet_name)

      path = output_path/name
      # creating a single file containing only that one sheet 
      df = pd.DataFrame(df)
      df.to_excel(path , index= False , engine= "openpyxl")

      i += 1 
      names.append(name)

   print(f"found {i}sheets from this file ,\n{names} \nand stored in {output_path} ,  ")


if __name__ == "__main__":


   base = Path(__file__).resolve().parents[1]
   c = base /"data"/"data_base_input"

   list_of_files = [f for f in c.glob("*.xlsx") ]

   for file in list_of_files:
      extract(file)
   print(list_of_files)
   print("done")