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
from src.dict_creator import run
from pathlib import Path
import pandas as pd 
import openpyxl

#=========================load the file ==================

BASE = Path(__file__).resolve().parents[1]
path = BASE / "data" / "cleaned"  / "sent_files_cleaned"
database_path = BASE / "cleaned" / "data_base_files_cleaned"
output_path = BASE / "data" / "woreda_added"


list_of_path = [file for file in path.glob("*.xlsx")]
list_of_path_for_database = [f for f in database_path.glob("*.xlsx")]




#==========================end======================

#================create the dictionary========================
BASE =  Path(__file__).resolve().parents[1]
path_for_dic = BASE / "data" / "cleaned"



list_of_files = [f for f in path.glob("*.xlsx")]

dictionary = run(list_of_files)

dictionary = pd.DataFrame()




#======================end=============================
#====================locate phone ========================
def add_woreda(path: Path , new_output_path: Path , dictionary  )->None:

    data = pd.read_excel(path , engine= "openpyxl") # main excel file to be exported after woreda is added 


    
    try:
        data["phone_number"] = data["phone_number"].astype(str)
        new_data = data.merge(
            dictionary , 
            left_on= "phone_number" , 
            right_on = "phone" , 
            how = "left"
        )
        try : 
            new_data = new_data.drop(columns= ["Phone Number", "Woreda"])
        except:
            pass

    
    except Exception:
        data["Phone Number"] = data["Phone Number"].astype(str)    
        new_data = data.merge(
            dictionary , 
            left_on= "Phone Number" , 
            right_on = "phone" , 
            how = "left"
        )
        try:
            new_data = new_data.drop(columns= ["phone", "woreda"])
        except:
            pass
    
    
    new_data.to_excel(new_output_path , index = False , engine= "openpyxl")


if __name__ == "__main__":
    for path in list_of_path_for_database:

        new_path = output_path / Path(path).name
        add_woreda(path , new_path , dictionary)





#=========================end=================================

#==========================add woreda column to data ==============













