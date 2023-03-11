import pandas as pd
import openpyxl as op
import string
import global_variables as gb
import csv
import os
class DivWiseSubjectResults:
    def __init__(self,subject_data):
        global subject
        subject=subject_data
        #load input file workbook
        self.gtu_workbook_path="inputfiles"+"/"+"SEM "+gb.e1.get()+"/"+"inputfile_SEM"+gb.e1.get()+".xlsx"
        self.gtu_workbook=op.load_workbook(self.gtu_workbook_path)
        #gb.df_list and gb.div_list are global varaible and already initialized in Div_wise_result file,so no need to define again as they  are defined in global_variables.
        self.sub_folders()
        #output df
        global output_df 
        output_df=pd.DataFrame()
        output_df[0]=['PASS','FAIL','TOTAL']
    def sub_folders(self):
        self.path1=os.path.join("outputfiles"+"/"+"SEM "+gb.e1.get()+"/"+"Division Wise",subject[0])
        try:
            os.mkdir(self.path1)
            print("subfolder created")
        except:
            print("Error!creating subfolder!!")


    def insert_into_csv(self,data,ind):
        counter=0
        #output file path
        self.output_file="outputfiles"+"/"+"SEM "+gb.e1.get()+"/"+"Division Wise"+"/"+subject[0]+"/"+gb.div_list[ind]+"_"+subject[0]+".txt"
        data.to_csv(self.output_file,header=['','Theor.','Prac.','Overall'],index=False,sep='\t',mode='a')#The values() function is used to get a Numpy representation of the DataFrame
        print("txt file saved!!")
        ind+=1
    def fetch_results(self):
        self.index=0
        for self.df in gb.df_list:
            #for theory
            self.theory_total=self.df[subject[2]].count()
            self.theory_fail=self.df.loc[(self.df[subject[2]]=='Y - - -')|(self.df[subject[2]]=='Y Y - -')|(self.df[subject[2]]=='Y - Y -')|(self.df[subject[2]]=='Y - - Y')|(self.df[subject[2]]=='Y Y Y -')|(self.df[subject[2]]=='Y - Y Y')|(self.df[subject[2]]=='Y Y - Y')|(self.df[subject[2]]=='Y Y Y Y')][subject[2]].count()
            self.theory_pass=self.theory_total-self.theory_fail
            output_df["THEORY"]=[self.theory_pass,self.theory_fail,self.theory_total]

            #for practical
            self.practical_total=self.df[subject[2]].count()
            self.practical_fail=self.df.loc[(self.df[subject[2]]=='- - - Y')|(self.df[subject[2]]=='- Y - Y')|(self.df[subject[2]]=='- - Y Y')|(self.df[subject[2]]=='Y - - Y')|(self.df[subject[2]]=='- Y Y Y')|(self.df[subject[2]]=='Y - Y Y')|(self.df[subject[2]]=='Y Y - Y')|(self.df[subject[2]]=='Y Y Y Y')][subject[2]].count()
            self.practical_pass=self.practical_total-self.practical_fail
            output_df["PRACTICAL"]=[self.practical_pass,self.practical_fail,self.practical_total]

            #for Overall
            self.overall_total=self.df[subject[1]].count()#overall total
            self.overall_fail=self.df.loc[self.df[subject[1]]=='FF'][subject[1]].count()#theory and practical total
            self.overall_pass=self.overall_total-self.overall_fail
            output_df["OVERALL"]=[self.overall_pass,self.overall_fail,self.overall_total]
            self.insert_into_csv(output_df,self.index)
            self.index+=1
    '''def div_result_fon_total(self):
            self.total_overall=self.df['SUB3GR'].count()#overall total
            self.total_practical=self.total_theory=self.df['SUB3B'].count()#theory and practical total
            
    def div_result_fon_overall(self):
        self.index=0
        for self.df in gb.df_list:
            self.fail_count=self.df.loc[self.df["SUB3GR"]=="FF"]["SUB3GR"].count()
            self.total=self.df['SUB3GR'].count()
            self.pass_count=self.total-self.fail_count
            print("Passed students in FON overall:",self.pass_count)
            print("Failed students in FON overall:",self.fail_count)
            print("Total students in FON overall:",self.total)
            self.data={"TOTAL":self.total,"PASS":self.pass_count,"FAIL":self.fail_count}
            self.data=pd.DataFrame(list(self.data.items()))
            self.insert_into_csv(self.data,self.index,"Overall")#passing div-wise overall-result to outputfile
            self.index+=1
    def div_result_fon_theory(self):
        self.index=0
        for self.df in gb.df_list:
            self.fail_count=self.df.loc[(self.df[subject[2]]=='Y - - -')|(self.df[subject[2]]=='Y Y - -')|(self.df[subject[2]]=='Y - Y -')|(self.df[subject[2]]=='Y - - Y')|(self.df[subject[2]]=='Y Y Y -')|(self.df[subject[2]]=='Y - Y Y')|(self.df[subject[2]]=='Y Y - Y')|(self.df[subject[2]]=='Y Y Y Y')]['EnrollmentNo'].count()
            self.total=self.df['SUB3B'].count()
            self.pass_count=self.total-self.fail_count
            print("Passed students in FON theory:",self.pass_count)
            print("Failed students in FON theory:",self.fail_count)
            self.data={"TOTAL":self.total,"PASS":self.pass_count,"FAIL":self.fail_count}
            self.data=pd.DataFrame(list(self.data.items()))
            self.insert_into_csv(self.data,self.index)#passing div-wise overall-result to outputfile
            self.index+=1
    def div_result_fon_practical(self):
        self.index=0
        for self.df in gb.df_list:
            #print("Total students in division",div,":",count)
            self.fail_count=self.df.loc[(self.df[subject[2]]=='- - - Y')|(self.df[subject[2]]=='- Y - Y')|(self.df[subject[2]]=='- - Y Y')|(self.df[subject[2]]=='Y - - Y')|(self.df[subject[2]]=='- Y Y Y')|(self.df[subject[2]]=='Y - Y Y')|(self.df[subject[2]]=='Y Y - Y')|(self.df[subject[2]]=='Y Y Y Y')]['EnrollmentNo'].count()
            self.total=self.df['SUB3B'].count()
            self.pass_count=self.total-self.fail_count
            print("Passed students in FON theory:",self.pass_count)
            print("Failed students in FON theory:",self.fail_count)
            self.data={"TOTAL":self.total,"PASS":self.pass_count,"FAIL":self.fail_count}
            self.data=pd.DataFrame(list(self.data.items()))
            self.insert_into_csv(self.data,self.index)#passing div-wise overall-result to outputfile
            self.index+=1
    
    '''
