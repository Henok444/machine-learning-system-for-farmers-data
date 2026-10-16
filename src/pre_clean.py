from pathlib import Path

import pandas as pd
import re


BASE  = Path(__file__).resolve().parents[1]

path = BASE  / "data" / "extracted" / "sent_files"



files = [f for f in path.glob("*.xlsx")]

out_put_path = BASE / "data" / "cleaned" / "sent_files_precleaned"
out_put_path_2 = BASE / "data" / "files_must_be_fixed"


paths =files[0]

class findit:
    def __init__(self , header = 0 , starting_col = 0 ):
        self.header = header
        self.column = starting_col
        pass

    def clean_text(self , list):
        ret_list = []
        for value  in list:

            text = re.sub(r"\s+", "", str(value)).strip().lower()
            ret_list.append(text)
        return ret_list
    
    def find_header(self , path):
        """  first find the the column header"""
        

        df_1 = pd.read_excel(path  , nrows= 0 , header= 0 , engine= "openpyxl")
        df_2 = pd.read_excel(path  , nrows= 0 , header= 1 , engine= "openpyxl")
        df_3 = pd.read_excel(path , nrows= 0 , header= 2 , engine= "openpyxl")



        df_1 = df_1.columns.to_list()
        df_2 = df_2.columns.to_list()
        df_3 = df_3.columns.to_list()

        df_1 = self.clean_text([str(l) for l in df_1])
        df_2 = self.clean_text([str(l) for l in df_2])
        df_3 = self.clean_text([str(l) for l in df_3])

        

        headers = [df_1 , df_2 , df_3]
        print(headers)

        list_of_names = ["PhoneNumber" , 'Sex' , "mobileno" , "sex" , "name" , "Name" , "FarmersName", "phone" , "contactnumber" , "phonenumber" , "Phone" , "phoneno", "phone_number" , "FarmerPhoneNumber" , "woreda" , "Woreda" , "WOREDA"  ]
        
        
        i = 0 

        found_df_1_headers = []
        found_df_2_headers = []
        found_df_3_headers = []


        for one_row in headers:
            if i == 0 :
                print(one_row)
                for item in one_row:  
                    print(item)   
                   
                    
                    if item in list_of_names:
                        found_df_1_headers.append(item)
                            
            if i == 1 :
                for string in one_row:
                    
                    if string in list_of_names:
                        found_df_2_headers.append(string)
                            
            if i == 2 :
                for string in one_row: 
                    if string in list_of_names:
                        found_df_3_headers.append(string)
                            
            i+=1
            
        header_dict = { "a" : found_df_1_headers  , "b": found_df_2_headers , "c" : found_df_3_headers}

        print(header_dict)

        largest_list = max(header_dict , key = lambda k: len(header_dict[k]) )

        if largest_list == "a":
            self.header = 0 
        if largest_list == "b":
            self.header = 1 
        if largest_list == "c":
            self.header = 2 

        
        
        if len(found_df_1_headers) == len(found_df_2_headers)  == len(found_df_3_headers) :
            return "error"
        
        return "go"
    def find_col(self, path):
        
        df_1 = pd.read_excel(path  , nrows= 5 , usecols = [0] , engine= "openpyxl")
        df_2 = pd.read_excel(path  , nrows= 5 , usecols = [1] , engine= "openpyxl")
        df_3 = pd.read_excel(path , nrows= 5 , usecols = [2], engine= "openpyxl")



        df_1 = df_1.iloc[1,0]
        df_2 = df_2.iloc[1,0]
        df_3 = df_3.iloc[1,0]
        headers = [df_1 , df_2 , df_3]



        print("hey" , df_1 , df_2 , df_3 , "hey")
        print(type(df_1))
        print(headers)

        i = 0 

        for header in headers:
            if isinstance(header , str):
                self.column = i 
                return self.column -1
            i+=1

    def clean(self , path , output_path):

        
        status = self.find_header(path)
        col = self.find_col(path)
        
        if status == "error":
            df = pd.read_excel(path , engine= "openpyxl")

            final_path = out_put_path_2 / Path(path).name

            df.to_excel(final_path ,index=False , engine= "openpyxl") 
                    
            return None
        # start checking the no of colimn called phone and wored if more than one send it to fix folder
        df_check = pd.read_excel(path , header= self.header , nrows=0 , engine= "openpyxl")
        df_check = df_check.columns.to_list()
        s=0
        k=0
        for l in df_check:
        
            if "phone" in str(l):
                s+=1 
            if "woreda" in str(l):
                k+=1
        if k>1 or s>1:
            df = pd.read_excel(path , engine= "openpyxl")

            final_path = out_put_path_2 / Path(path).name

            df.to_excel(final_path ,index=False , engine= "openpyxl") 
                    
            return None   
        # end checking the no of colimn called phone and wored if more than one send it to fix folder       
        
        df = pd.read_excel(path , header= self.header , engine= "openpyxl")
        df = df.iloc[:, col:]

        final_path = output_path / Path(path).name


        df.to_excel(final_path ,index=False , engine= "openpyxl") 


if __name__ == "__main__":
    

    path = BASE  / "data" / "extracted" / "sent_files"



    files = [f for f in path.glob("*.xlsx")]

    out_put_path = BASE / "data" / "cleaned" / "sent_files_precleaned"

    paths =files[0]




    clean = findit()
    

    for file in files:
        find = findit()

        find.clean(file , out_put_path)







