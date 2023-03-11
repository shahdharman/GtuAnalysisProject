import pandas as pd
import openpyxl as op
import string
import global_variables as gb
import os
class OverallResult:
    def __init__(self):
        #load gtu file
        self.gtu_workbook_path="gtufiles"+"/"+"SEM "+gb.e1.get()+"/"+"SEM"+gb.e1.get()+"_gtufile.xls"
        #load gtu file workbook
        self.gtu_workbook_df=pd.read_excel(self.gtu_workbook_path)

    #subfolder names "Overall Result"
    def sub_folders(self):
        self.path1=os.path.join("outputfiles"+"/"+"SEM "+gb.e1.get(),"OVERALL")
        try:
            os.mkdir(self.path1)
            print("subfolder created")
        except:
            print("Error!creating subfolder!!")
            
    #insert data into created csv outputfile. 
    def insert_into_csv(self,data):
        self.sub_folders()
        
        #output file path
        self.output_file="outputfiles"+"/"+"SEM "+gb.e1.get()+"/"+"OVERALL"+"/"+"OverallResult.txt"
        data.to_csv(self.output_file, header=['Overall','Count'], index=None, sep='\t',mode='w')#The values() function is used to get a Numpy representation of the DataFrame
        print("txt file saved!!")
        
    def result_overall(self):
        self.df_result=self.gtu_workbook_df['RESULT']
        self.pass_counter=0
        self.fail_counter=0
        self.total=0
        for self.res in self.df_result:
            self.total+=1
            if self.res=="PASS":
                self.pass_counter+=1
            else:
                self.fail_counter+=1
        self.result=(self.pass_counter/self.total)*100
        print("Total students",self.total)
        print("Passed students:",self.pass_counter)
        print("Failed students:",self.fail_counter)
        print("Result:",self.result)
        self.data={"TOTAL":self.total,"PASS":self.pass_counter,"FAIL":self.fail_counter,"Result(%)":self.result}
        self.data=pd.DataFrame(list(self.data.items()))
        self.insert_into_csv(self.data)#passing overall-result to outputfile
          
