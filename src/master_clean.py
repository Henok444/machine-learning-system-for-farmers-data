import pandas as pd 
import openpyxl
from pathlib import Path
import re 



"""
The goal of this script is to clean the excel file sent from the the farmer 

This logic only apply after the sheet_extractor has finished running , 
then after that this script input each file (extracted) and clean there woreda and phone column


"""

"""psdo code


first do it only for one file and simple then expand 

first  find the path of the file --------done
locate the phone and woreda column 

first clean the name of the column to proper constant name 

then clean the each phone no make them constant 

fill the empty woreda or phone no with nan 

export them to new destination 



"""
#===========================================find column name of woreda and column no and clean them ===========================
BASE = Path(__file__).resolve().parents[1]

path = BASE / "data" / "cleaned" / "sent_files_precleaned"
output_path_2 = BASE / "data" / "no_woreda_files"


class clean_all:
    def __init__(self , header_row = 0 ):

        self.header = header_row
    def name_clean(self , list: list )->list:
        i = 0 
        for name in list:
            if not isinstance(name ,str ):
                continue
            cleaned = re.sub(r"\s+" , "" , name).lower()
            list[i] = cleaned
            i+=1
        return list



    def find_column_no(self , path):

        df = pd.read_excel(path,  nrows= 0 , engine= "openpyxl")
        df_2 = pd.read_excel(path , nrows = 0 , header= 1 , engine = "openpyxl")
        df_2 = df_2.columns.to_list()
        df = pd.DataFrame(df)



        list_of_phone_names = ["PhoneNumber" , "phone" , "contactnumber" , "phonenumber" , "Phone" , "phoneno", "phone_number" , "farmerphonenumber"]
        list_of_woreda_names = ["woreda" , "Woreda" , "WOREDA" ]
        

        df = df.columns.tolist()

    

        woreda_column_n = 0 
        phone_column_n = 0 
        #if (df[1] == "Unnamed: 1"  or df[2] == "Unnamed: 2" or df[3] == "Unnamed: 3") and isinstance(df_2[0] , str):
        # Check if "Unnamed" is inside EVERY column name string
        i = 1
        print("ssssssssssssssssssssssssssssssssssssssssssssssssssss")
        print(all("Unnamed" in str(col) for col in df) , (any("Unnamed" in str(col) for col in df) , isinstance(df_2[0] , str)  , df_2[0] != "Unnamed: 0"))
        while  all("Unnamed" in str(col) for col in df) or (any("Unnamed" in str(col) for col in df) and isinstance(df_2[0] , str) and df_2[0] != "Unnamed: 0" ) or any("Farmer" in str(col) for col in df_2) or any("Name" in str(col) for col in df_2) :
            print("flag of the emptyness for visual =>>" , "&" * 100)
            print("🚨 Every single column name is an Unnamed column!")
            print("info about the second row data type" , df_2[0] , type(df_2[0]))
            print(f"The columns are ==> (({df}))")

            df_2 = pd.read_excel(path , nrows = 0 , header= i+1 , engine = "openpyxl")
            df_2 = df_2.columns.to_list()

            df = pd.read_excel(path,  nrows= 0 , header= i, engine= "openpyxl")
            self.header = i 
            df = df.columns.tolist()
            print( "--" * 100)
            print(f"So using the second row as a header ==> (({df}))")
            i += 1
            print("flag of the emptyness for visual =>>" , "&" * 100)
        
        self.df = df 
        self.df_2 = df_2  

        df = self.name_clean(df)

        i = 0    
        j = 0
        k = 0   

        for name in df:
            print(name , "phone")
            if "phone"  in str(name): 
                print("yes")
                
                phone_column_n = i 
                j+=1
            if "woreda" in str(name):
                print("yes")
                woreda_column_n = i 
                k+=1
            i+=1 
        
        if j == 0 or k ==0:#  to check if there is no phone or woreda in the header
            return None , None , "error"
        return phone_column_n ,woreda_column_n , "ok"
    #=======================================================end==========================================

    #==================================================clean each phone no ===============================
    def clean_text(self , value):
        if value is None:
            return None
        try:
            if pd.isna(value):
                return None
        except (TypeError, ValueError):
            pass
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        text = re.sub(r"\s+", "", str(value)).strip()
        return text or None

    def normalize_phone(self , value):
        text = self.clean_text(value)
        if not text:
            return None
        if re.fullmatch(r"\d+(\.\d+)?[eE]\+?\d+", text):
            text = str(int(float(text)))
        text = re.sub(r"\.0+$", "", text)
        digits = re.sub(r"\D", "", text)
        if not digits:
            return None
        if digits.startswith("00251") and len(digits) == 14:
            digits = digits[5:]
        elif digits.startswith("251") and len(digits) == 12:
            digits = digits[3:]
        elif digits.startswith("0") and len(digits) == 10:
            digits = digits[1:]
        if len(digits) == 9 and digits[0] in "79":
            return "0" + digits
        return digits
    #====================================================  end     =================================================

    #============================================clean woreda and normalize the text =====================


    def normalize_name(self , lists: list )->list:
        i = 0 
        for name in lists:
            if not isinstance(name , str):
                lists[i] = None 
                continue
            cleaned = re.sub(r"\s+" , "_" , name).lower()
            lists[i] = cleaned
            i+=1
        return lists

    #===================================================end======================================
    def clean_sheet(self , sheet_path : Path , output_path  : Path )-> None :

        print(f"----------cleaning the sheet (({sheet_path.name}))")

        # phone and woreda column name are cleaned ================================================
        
        phone , woreda , status= self.find_column_no(sheet_path)# column no found and cleaned 
        if status == "error":
            sheet = pd.read_excel(sheet_path ,header=  self.header ,  engine= "openpyxl")

            final_path = output_path_2 / Path(sheet_path).name

            sheet.to_excel(final_path , index= False ,engine= "openpyxl")
            return None

        print(phone , woreda)



        # now extract list phonen no--------------------------------------------------------------


        phone_sheet= pd.read_excel(sheet_path , usecols= [phone] , header= self.header , engine= "openpyxl")




        phone_list = phone_sheet.iloc[: , 0].tolist()
        
        # clean every phone no and create a list of phone no ========
        i = 0 
        for phones in phone_list:


            phone_list[i] = self.normalize_phone(phones)
            i+=1
        # created ================= phone_list is the list of phone no 

        # now extract list of woreda----------------------------------------------------------------------
        woreda_sheet = pd.read_excel(sheet_path ,  usecols=[woreda] , header= self.header , engine= "openpyxl")



        woreda_list = woreda_sheet.iloc[:, 0 ].to_list()


        #clean the woreda string names 
        woreda_list = self.normalize_name(woreda_list)
        #=================================================end cleaning each name ==============================

        #=======================================rename the column name and overwrite the column with the new list 

        sheet = pd.read_excel(sheet_path ,header=  self.header ,  engine= "openpyxl")

        print()

        new_woreda_name = "woreda"
        new_phone_no_name = "phone"

        headers = list(sheet.columns)
        headers[woreda] = new_woreda_name
        headers[phone] = new_phone_no_name

        sheet.columns = headers 
        sheet[new_woreda_name] = woreda_list
        sheet[new_phone_no_name] = phone_list

        final_path = output_path / Path(sheet_path).name

        sheet.to_excel(final_path , index= False ,engine= "openpyxl")
        print("done")


        

        """
        creating the cleanin 


        """

if __name__ == "__main__":
    BASE = Path(__file__).resolve().parents[1]

    path = BASE / "data" / "cleaned" / "sent_files_precleaned"
    output_path = BASE / "data" / "cleaned" / "sent_files_cleaned"

    list_of_files = [file for file in path.glob("*.xlsx")]
    
    print(f"found (({len(list_of_files)})) files to be cleaned  and cleaned ")

    for file in list_of_files:
        clean = clean_all()
        clean.clean_sheet(file , output_path)
    print(len(list_of_files))



