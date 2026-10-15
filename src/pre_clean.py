from pathlib import Path

import pandas as pd


BASE  = Path(__file__).resolve().parents[1]

path = BASE  / "data" / "extracted" / "sent_files"



files = [f for f in path.glob("*.xlsx")]

out_put_path = BASE / "data" / "cleaned" / "sent_files_precleaned"

paths =files[0]

class findit:
    def __init__(self , header = 0 , starting_col = 0 ):
        self.header = header
        self.column = starting_col
        pass
    def find_header(self , path):
        """  first find the the column header"""
        

        df_1 = pd.read_excel(path  , nrows= 0 , header= 0 , engine= "openpyxl")
        df_2 = pd.read_excel(path  , nrows= 0 , header= 1 , engine= "openpyxl")
        df_3 = pd.read_excel(path , nrows= 0 , header= 2 , engine= "openpyxl")



        df_1 = df_1.columns.to_list()
        df_2 = df_2.columns.to_list()
        df_3 = df_3.columns.to_list()

        df_1 = [str(l) for l in df_1]
        df_2 = [str(l) for l in df_2]
        df_3 = [str(l) for l in df_3]

        

        headers = [df_1 , df_2 , df_3]
        print(headers)

        list_of_names = ["PhoneNumber" , 'Sex' , "sex" , "name" , "Name" , "Farmers Name", "phone" , "contactnumber" , "phonenumber" , "Phone" , "phoneno", "phone_number" , "Farmer Phone Number" , "woreda" , "Woreda" , "WOREDA"  ]
        
        
        i = 0 

        found_df_1_headers = []
        found_df_2_headers = []
        found_df_3_headers = []


        for one_row in headers:
            if i == 0 :
                for string in one_row:           
                    
                    if string in list_of_names:
                        found_df_1_headers.append(string)
                            
            if i == 1 :
                for string in one_row:
                    
                    if string in list_of_names:
                        found_df_2_headers.append(string)
                            
            if i == 2 :
                    
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
            self.header == 2 
        return pd.read_excel(path , header= self.header , engine= "openpyxl")
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
        self.find_header(path)
        col = self.find_col(path)
        df = pd.read_excel(path , header= self.header , engine= "openpyxl")
        df = df.iloc[:, col:]

        final_path = output_path / Path(path).name

        df.to_excel(final_path ,index=False , engine= "openpyxl") 

    
for file in files:
    find = findit()

    find.clean(file , out_put_path)






