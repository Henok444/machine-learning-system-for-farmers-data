"""


mapping 


idea we want all the data base data to have woreda in them 



so we map them from the dictionarry 

psdo 

created simply for one file from the data base  , one the dictionary 


load the file ------------done 
create the dictionary -----------done 
locate phone from the database -------------done 
create a list 
add a woreda column to the dataframe ----------------
for the first row phone value  match the respective phone from the dictionary 
extract the woreda for that phone no  ,
fill the woreda column in that same row with the woreda found 
do this for all 
export the excel file 


"""

from pathlib import Path
import pandas as pd 
import openpyxl

#=========================load the file ==================

BASE = Path(__file__).resolve().parents[1]
path = BASE / "data" / "cleaned"  / "sent_files_cleaned"
database_path = BASE / "data" /"cleaned" / "data_base_files_cleaned"
output_path = BASE / "data" / "woreda_added"
dic_path = BASE / "data" / "dict" / "dict.xlsx"


list_of_path = [file for file in path.glob("*.xlsx")]
list_of_path_for_database = [f for f in database_path.glob("*.xlsx")]




#==========================end======================

#================create the dictionary========================
BASE =  Path(__file__).resolve().parents[1]








#======================end=============================
#====================locate phone ========================
def add_woreda(path: Path , new_output_path: Path  )->None:

    data = pd.read_excel(path , engine= "openpyxl") # main excel file to be exported after woreda is added 
    dictionary = pd.read_excel(dic_path , engine= "openpyxl")


    

    dictionary["phone"] = dictionary["phone"].astype(str)
    dictionary["phone"] = dictionary["phone"].astype(str).str.replace(r"\.0" , "" , regex = True )

    
    #print(data.columns)


    list = data.columns.to_list()
   # print(data.columns)
    #print(dictionary)
    if "Woreda" in list or "woreda" in list:
        print("heyyyy")
        data.to_excel(new_output_path , index = False , engine= "openpyxl")
        return None


    if "Phone Number" in list:
        data = data.rename(columns= {"Phone Number" ,  "phone_number"} , inplace= True)
    print(data.columns)

    data["phone_number"] = data["phone_number"].astype(str)

   # print(data["phone_number"])


    new_data = data.merge(
        dictionary , 
        left_on= "phone_number" , 
        right_on = "phone" , 
        how = "left"
        )
    new_data['woreda'] = new_data['woreda'].fillna("no phone match in the dictionary")
    new_data.to_excel(new_output_path , index = False , engine= "openpyxl")
    #print(new_data)

if __name__ == "__main__":
    for path in list_of_path_for_database:

        new_path = output_path / Path(path).name
        add_woreda(path , new_path)




#=========================end=================================

#==========================add woreda column to data ==============













