"""

the idea 

=====================================================================================================
there is two folder one contain all the database file that is one sheet per file 
the other have a files that is sent by agents 


what we want is all the database file to have woreda in them , 

====================================================================================================
 
proposed solution 

so what i proposed is first lets create a dictionary with phone and woreda 
then for every file map there woreda using phone no from the dict 

=========================================================================
what i have already 

1) I have a function that extract every sheet from a file that contain multiple sheets 
2) I have a function that if you give it two files that can correctly map them and add woreda and document it well 

===================================================================

so first create the dict 

a function that takes many file and output one massive dictionary 

so for each file that is in database folder 

we take one and give it to the add_woreda function with dictionary created as one massive xlsx file 
then the function recive two files and output one 
and do this for all files 

c
then for every checked file 

    create a woreda column 
    extract there phone no for each row 
    find a phone no that match it in the dictionary 
    add the value of that key in that row of woreda column 
    do this for each row 
    do this for every file 

"""

# create the dictionary 

"""
for every file extract phone and woreda 

find the file ----done 
from header locate woreda ----------done 
from header locate phone_now ----------done 
create a dictionary that put phone and woreda as key and value repectively 

do it first for one  then expand 

"""
#==============================================find phone and woreda =====================

from pathlib import Path 
import pandas as pd 
import re




BASE =  Path(__file__).resolve().parents[1]
path = BASE / "data" / "cleaned"
output_path = BASE / "data" / "dict" / "dict.xlsx"

list_of_files = [f for f in path.glob("*.xlsx")]




def find_column_no(path):

    df = pd.read_excel(path,  nrows= 0 , engine= "openpyxl")

    list_of_phone_names = ["Phone Number" , "phone" , "Contact Number" , "phone" , "Phone" , "PHONE No",]
    list_of_woreda_names = ["woreda" , "Woreda" , "WOREDA" ]

    df = df.columns.tolist()
   

    woreda_column_n = 0 
    phone_column_n = 0 

    i = 0 

    for name in df:

        if name in  list_of_phone_names : 
            phone_column_n = i 
        if name  in list_of_woreda_names:
            woreda_column_n = i 
            woreda_name = name 
        i+=1 

    return phone_column_n ,woreda_column_n , woreda_name

# =================================================

def create_dict(path )-> dict :
    """
    locate one row create one key and value go to next 
    """
    phone , woreda , woreda_name = find_column_no(path)
    df = pd.read_excel(path, usecols= [phone , woreda], engine= "openpyxl")

    
   
    return df , woreda_name




def run(files):
    
    
    
    i = 0 
    j = 0 
    
    for file in files:
        print("@"*150)
        if i == j:
            df , woreda_name = create_dict(file)
            print("heyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyy")
            print([item for item in df[woreda_name] if isinstance(item , int)])
            print(any([item for item in df[woreda_name]if isinstance(item , int)]))
            print([item for item in df[woreda_name] if (isinstance(item , float) and not pd.isna(item))]) 
            print(any([item for item in df[woreda_name]  if (isinstance(item , float) and not pd.isna(item))])) 
            print("heyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyy")

            if any([item for item in df[woreda_name] if  isinstance(item , int)]) or any([item for item in df[woreda_name]  if (isinstance(item , float) and not pd.isna(item))]) :
                print(f"🚨🚨🚨🚨🚨the file ((({file})))  contain no in the woreda column must be reviewed and cleaned by hand so it is droped out of the dictionary")
                i+=1
                j+=1
                print(df)
                print("@"*150)
                continue
            i+=1
            print("@"*150)
            continue

        
        next_df , woreda_name = create_dict(file)
        print("heyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyy")
        print([item for item in next_df[woreda_name] if isinstance(item , int)])
        print(any([item for item in next_df[woreda_name] if isinstance(item , int)]))
        print([item for item in next_df[woreda_name]  if (isinstance(item , float) and not pd.isna(item))])
        print(any([item for item in next_df[woreda_name]  if (isinstance(item , float) and not pd.isna(item))]))
        print("heyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyy")
        if any([item for item in next_df[woreda_name] if isinstance(item , int)]) or any([item for item in next_df[woreda_name] if (isinstance(item , float) and not pd.isna(item))]) :
            print(f"🚨🚨🚨🚨🚨the file ((({file})))  contain no in the woreda column must be reviewed and cleaned by hand so it is droped out of the dictionary")
            i+=1
            j+=1
            print(next_df)
            continue
        df = pd.concat([df, next_df] , ignore_index= True)
        i +=1

    df["phone"] = df["phone"].astype(str).str.replace(r"\.0" , "" , regex = True ) 
    print("@"*150)
    return df

if __name__ == "__main__":

    df = run(list_of_files)
    df.to_excel(output_path , index = False , engine= "openpyxl")
    print(df)

