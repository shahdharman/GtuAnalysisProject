import pandas as pd
import global_variables as gb
import openpyxl as op
import os
class DivWiseOverall:
    def __init__(self):
        
        self.gtu_workbook_path="inputfiles"+"/"+"SEM "+gb.e1.get()+"/"+"inputfile_SEM"+gb.e1.get()+".xlsx"
        #load input file workbook
        self.gtu_workbook=op.load_workbook(self.gtu_workbook_path)
        # list of sheet_names (which tells total divisions)
        gb.div_list=self.gtu_workbook.sheetnames
        #list of no. of dataframes in inputfile.
        #creating df of each sheet & appending df to gb.df_list
        for self.ws in gb.div_list:
            self.df=pd.read_excel(self.gtu_workbook_path,sheet_name=self.ws)
            gb.df_list.append(self.df)
            #dataframe of input file workbook    
        #print("divisions:",gb.div_list)
        self.sub_folders()
    #subfolder names "Overall Result"
    def sub_folders(self):
        self.path1=os.path.join("outputfiles"+"/"+"SEM "+gb.e1.get(),'Division Wise')
        self.path2=os.path.join("outputfiles"+"/"+"SEM "+gb.e1.get()+"/"+"Division Wise","Overall")
        try:
            os.mkdir(self.path1)
            print("subfolder created")
        except:
            print("path1!!Error!creating subfolder!!")
        try:
            os.mkdir(self.path2)
            print("subfolder created")
        except:
            print("path2!!Error!creating subfolder!!")
            
    #insert data into created csv outputfile. 
    def insert_into_csv(self,data,ind):
        #output file path
        self.output_file="outputfiles"+"/"+"SEM "+gb.e1.get()+"/"+"Division Wise"+'/'+"Overall"+"/"+gb.div_list[ind]+"_Overall.txt"
        ind+=1
        print(self.output_file)
        data.to_csv(self.output_file, header=['Overall','Count'], index=None, sep='\t',mode='w')#The values() function is used to get a Numpy representation of the DataFrame
        print("txt file saved!!")
        
    #Div-wise overall PASS/FAIL Result
    def div_result_overall(self):
        self.index=0
        #counting pass/fail in  each sheet 
        for self.each in gb.df_list:
            self.pass_count=self.each.loc[self.each['RESULT']=="PASS"]['RESULT'].count()#counts pass
            print("Passed students :",self.pass_count)
            self.fail_count=self.each.loc[self.each.RESULT=="FAIL"]['RESULT'].count()
            print("Failed students in :",self.fail_count)
            self.total=self.each['RESULT'].count()
            self.result=(self.pass_count/self.total)*100
            print("Total",self.total)
            #create combined dataframe of results.
            self.data={"TOTAL":self.total,"PASS":self.pass_count,"FAIL":self.fail_count,"Result(%)":self.result}
            self.data=pd.DataFrame(list(self.data.items()))
            self.insert_into_csv(self.data,self.index)#passing div-wise overall-result to outputfile
            self.index+=1
