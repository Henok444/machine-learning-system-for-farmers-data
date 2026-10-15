
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
   output_path = BASE /"data"/ "extracted" / "sent_files"
   output_path_2 = BASE / "data" / "extracted" /  "data_base_files"
   output_path_3 = BASE / "data" / "unrelated_files"
   print("=" *200)
   print(f"extracting the file :---(({file})) ")

   # if there is only one sheet , we want it to have the name it has before
   no_of_sheet = len([f for f, df in multi_sheet.items()])
   print(f"found ({no_of_sheet}) sheet in this file --------")



   if  no_of_sheet <= 1:
      
      sheet = pd.read_excel(input_path  , engine= "openpyxl")
   
      column_name_1 = sheet.columns[0]
      column_name_2 = sheet.columns[1]
      

      if column_name_1 == "id" or  column_name_2 == "created_by":
         output_path = output_path_2
         print(f"the file is data_base_file so it is stored to data_base_extracted folder---------")
      if column_name_1 == "Farmer UID" or  column_name_2 == "Decision":
         output_path = output_path_3
         print(f"the file is unrelated so it is stored to unrelated_files  folder---------")

      main_name = Path(file).stem


      final_name = main_name + "_sheet_name_" + "only_one_sheet"

      name = clean_name(final_name)


      path = output_path / Path(input_path).name

      sheet.to_excel(path , engine = "openpyxl")
      print("the file has only one sheet so it is stored to data_base_cleaned directly and the extract function is breaked here--------")
      print(f"------Created {no_of_sheet} independent sheet from this file")
      return no_of_sheet
   

   # for each sheet save the name of the sheet 
   i = 0 
   names = []
   for sheet_name , df in multi_sheet.items(): # tested and passed 

      # new added 0ct 8 1152
      main_name = Path(file).stem


      final_name = main_name + "_sheet_name_" + sheet_name 

      name = clean_name(final_name)
      # new added 0ct 8 1152

      # creating a single file containing only that one sheet 
      if df.empty:

         print(f"---the excel file (({file})) and sheet name  {sheet_name} is empty so it is jumped and ignored -----")
         no_of_sheet -= 1 
         continue
      df = pd.DataFrame(df)



      """
      here detecting which folder it belongs will be detected 

      psdo 

      

      """
     # print(df)

      column_name_1 = df.columns[0]
      column_name_2 = df.columns[1]

      if column_name_1 == "id" or  column_name_2 == "created_by":
         output_path = output_path_2
      if column_name_1 == "Farmer UID" or  column_name_2 == "Decision":
         output_path = output_path_3
         

      path = output_path/name

      
      # new added 0ct 8 1152_2

      # new added 0ct 8 1152_2
      

      df.to_excel(path , index= False , engine= "openpyxl")

      i += 1 
      names.append(name)
   print(f"------Created {no_of_sheet} independent sheet from this file")
   return no_of_sheet



if __name__ == "__main__":


   base = Path(__file__).resolve().parents[1]
   c = base /"data"/"input"

   list_of_files = [f for f in c.glob("*.xlsx") ]
   total_no_sheet = 0 
   for file in list_of_files:
      no_of_sheet = extract(file)
      total_no_sheet = total_no_sheet + no_of_sheet 
   print(f"Number of files found to be extracted is (({len(list_of_files)})) and extracted (({total_no_sheet})) ")
   print("--------------extracting part is done")