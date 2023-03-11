import pandas as pd 
import openpyxl as op
import string
import global_variables as gb
import csv
import os
class SubjectWiseResults:
    def __init__(self,subject_data):
        global subject
        subject=subject_data
        #load gtu file
        self.gtu_workbook_path="gtufiles"+"/"+"SEM "+gb.e1.get()+"/"+"SEM"+gb.e1.get()+"_gtufile.xls"
        #load gtu file workbook
        self.gtu_workbook_df=pd.read_excel(self.gtu_workbook_path)
        global output_df 
        output_df=pd.DataFrame()
        output_df[0]=['PASS','FAIL','TOTAL']
    #insert data into created csv outputfile. 
    def insert_into_csv(self,data):
        #output file path
        self.output_file="outputfiles"+"/"+"SEM "+gb.e1.get()+"/"+"OVERALL"+"/"+subject[0]+".txt"
        data.to_csv(self.output_file,header=['','Theor.','Prac.','Overall'],index=False,sep='\t',mode='a')#The values() function is used to get a Numpy representation of the DataFrame
        print("txt file saved!!")
    def fetch_results(self):
        #for theory
        self.theory_total=self.gtu_workbook_df[subject[2]].count()
        self.theory_fail=self.gtu_workbook_df.loc[(self.gtu_workbook_df[subject[2]]=='Y - - -')|(self.gtu_workbook_df[subject[2]]=='Y Y - -')|(self.gtu_workbook_df[subject[2]]=='Y - Y -')|(self.gtu_workbook_df[subject[2]]=='Y - - Y')|(self.gtu_workbook_df[subject[2]]=='Y Y Y -')|(self.gtu_workbook_df[subject[2]]=='Y - Y Y')|(self.gtu_workbook_df[subject[2]]=='Y Y - Y')|(self.gtu_workbook_df[subject[2]]=='Y Y Y Y')][subject[2]].count()
        self.theory_pass=self.theory_total-self.theory_fail
        output_df["THEORY"]=[self.theory_pass,self.theory_fail,self.theory_total]

        #for practical
        self.practical_total=self.gtu_workbook_df[subject[2]].count()
        self.practical_fail=self.gtu_workbook_df.loc[(self.gtu_workbook_df[subject[2]]=='- - - Y')|(self.gtu_workbook_df[subject[2]]=='- Y - Y')|(self.gtu_workbook_df[subject[2]]=='- - Y Y')|(self.gtu_workbook_df[subject[2]]=='Y - - Y')|(self.gtu_workbook_df[subject[2]]=='- Y Y Y')|(self.gtu_workbook_df[subject[2]]=='Y - Y Y')|(self.gtu_workbook_df[subject[2]]=='Y Y - Y')|(self.gtu_workbook_df[subject[2]]=='Y Y Y Y')][subject[2]].count()
        self.practical_pass=self.practical_total-self.practical_fail
        output_df["PRACTICAL"]=[self.practical_pass,self.practical_fail,self.practical_total]

        #for Overall
        self.overall_total=self.gtu_workbook_df[subject[1]].count()#overall total
        self.overall_fail=self.gtu_workbook_df.loc[self.gtu_workbook_df[subject[1]]=='FF'][subject[1]].count()#theory and practical total
        self.overall_pass=self.overall_total-self.overall_fail
        output_df["OVERALL"]=[self.overall_pass,self.overall_fail,self.overall_total]
        self.insert_into_csv(output_df)
        
    '''
    def fon_result_overall(self):
        print("FON Overall!!")
        fon_overall=self.df["SUB3GR"]
        pass_counter=0
        fail_counter=0
        total=0
        for res in fon_overall:
            total+=1
            if res=="FF":
                fail_counter+=1
            else:
                pass_counter+=1
        result=(pass_counter/total)*100
        print("Total students",total)
        print("Inserting FON total students Overall in excel!!")
        self.insert_into_excel(total)   
        print("Passed students:",pass_counter)
        print("Inserting FON pass students Overall in excel!!")
        self.insert_into_excel(pass_counter)
        print("Failed students:",fail_counter)
        print("Inserting FON Fail students Overall in excel!!")
        self.insert_into_excel(fail_counter)
        print("Result:",result)
        print("Inserting FON RESULT Overall in excel!!")
        self.insert_into_excel(result)
    def fon_result_theory(self):
        #print(df["SUB1B"])
        print("FON Theory!!")
        pass_counter=0
        fail_counter=0
        total=0
        fon_theory=self.df["subject[2]"]
        for res in fon_theory:
            total+=1
            if (res=='Y - - -')|(res=='Y Y - -')|(res=='Y - Y -')|(res=='Y - - Y')|(res=='Y Y Y -')|(res=='Y - Y Y')|(res=='Y Y - Y')|(res=='Y Y Y Y'):
                 fail_counter+=1
            else:
                pass_counter+=1
        result=(pass_counter/total)*100
        print("Total students",total)
        print("Inserting FON Total Students in Thoery in excel!!")
        self.insert_into_excel(total)   
        print("Passed students:",pass_counter)
        print("Inserting FON Pass Students in Thoery in excel!!")
        self.insert_into_excel(pass_counter)   
        print("Failed students:",fail_counter)
        print("Inserting FON Fail Students in Thoery in excel!!")
        self.insert_into_excel(fail_counter)   
        print("Result:",result)
        print("Inserting FON RESULT in Thoery in excel!!")
        self.insert_into_excel(result)   
    
    def fon_result_practical(self):
        #print(df["SUB1B"])
        print("FON Practical!!")
        pass_counter=0
        fail_counter=0
        total=0
        fon_practical=self.df["subject[2]"]
        for res in fon_practical:
            total+=1
            if (res=='- - - Y')|(res=='- Y - Y')|(res=='- - Y Y')|(res=='Y - - Y')|(res=='- Y Y Y')|(res=='Y - Y Y')|(res=='Y Y - Y')|(res=='Y Y Y Y'):
                 fail_counter+=1
            else:
                pass_counter+=1
        result=(pass_counter/total)*100
        print("Total students",total)
        print("Inserting FON Total Students in Practical in excel!!")
        self.insert_into_excel(total)  
        print("Passed students:",pass_counter)
        print("Inserting FON Pass Students in Practical in excel!!")
        self.insert_into_excel(pass_counter)  
        print("Failed students:",fail_counter)
        print("Inserting FON Fail Students in Practical in excel!!")
        self.insert_into_excel(fail_counter)  
        print("Result:",result)
        print("Inserting FON RESULT in Practical in excel!!")
        self.insert_into_excel(result)
        '''
