import pandas as pd 
import openpyxl as op
import global_variables as gb
import string
import os
import csv
class DivWiseCurrentBacklogs:
    def __init__(self):
        #load input file workbook
        self.gtu_workbook_path="inputfiles"+"/"+"SEM "+gb.e1.get()+"/"+"inputfile_SEM"+gb.e1.get()+".xlsx"
        self.gtu_workbook=op.load_workbook(self.gtu_workbook_path)
        #gb.df_list and gb.div_list are global varaible and already initialized in Div_wise_result file,so no need to define again as they  are defined in global_variables.
        self.sub_folders()
        self.creating_csvs()
        #print(gb.div_list)
        #print(gb.df_list)
    def creating_csvs(self):
        for self.each in range(len(gb.div_list)):
            with open("outputfiles"+"/"+"SEM "+gb.e1.get()+'/'+"Division Wise"+"/"+"CurrentBacklogs"+"/"+gb.div_list[self.each]+"_CurrentBacklogs.txt", 'w', newline='') as self.csvfile:
                # creating a csv writer object  
                self.csvwriter = csv.writer(self.csvfile,delimiter=" ")
                self.header=['CrrBacks','Count']
                self.csvwriter.writerow(self.header)#writes as first row
        print("Csv files created")

    #subfolder names "CurrentBacklogs"
    def sub_folders(self):
        self.path1=os.path.join("outputfiles"+"/"+"SEM "+gb.e1.get()+'/'+"Division Wise","CurrentBacklogs")
        try:
            os.mkdir(self.path1)
            print("subfolder created")
        except:
            print("Error!creating subfolder!!")

    #insert data into created csv outputfile.
    def insert_into_csv(self,data,index):
        #output file path
        self.output_file="outputfiles"+"/"+"SEM "+gb.e1.get()+'/'+"Division Wise"+"/"+"CurrentBacklogs"+"/"+gb.div_list[index]+"_CurrentBacklogs.txt"
        print(self.output_file)
        data.to_csv(self.output_file,header=None, index=None, sep='\t',mode='a')#The values() function is used to get a Numpy representation of the DataFrame
        print("txt file saved!!")
        
    def current_backlogs_0(self):
        self.index=0
        self.my_dict={}
        #counting 0 backlogs in each sheet
        for self.df in gb.df_list:
            counter=0
            curback=self.df['CURBACKL']
            for backlog in curback:
                if backlog==0:
                    counter+=1
            #appending to dictionary
            self.my_dict['0']=counter
            self.df=pd.DataFrame(list(self.my_dict.items()))
            print("a:",self.index)
            self.insert_into_csv(self.df,self.index)
            self.index+=1
    def current_backlogs_1(self):
        self.index=0
        self.my_dict={}
        #counting 1 backlogs in each sheet
        for self.df in gb.df_list:
            counter=0
            curback=self.df['CURBACKL']
            for backlog in curback:
                if backlog==1:
                    counter+=1
            self.my_dict['1']=counter
            self.df=pd.DataFrame(list(self.my_dict.items()))
            self.insert_into_csv(self.df,self.index)
            self.index+=1
    def current_backlogs_2(self):
        self.index=0
        self.my_dict={}
        #print(df["SUB1B"])
        #counting 2 backlogs in each sheet
        for self.df in gb.df_list:
            counter=0
            curback=self.df['CURBACKL']
            for backlog in curback:
                if backlog==2:
                    counter+=1
            self.my_dict['2']=counter
            self.df=pd.DataFrame(list(self.my_dict.items()))
            self.insert_into_csv(self.df,self.index)
            self.index+=1
    def current_backlogs_3(self):
        self.index=0
        self.my_dict={}
        #print(df["SUB1B"])
        #counting 3 backlogs in each sheet
        for self.df in gb.df_list:
            counter=0
            curback=self.df['CURBACKL']
            for backlog in curback:
                if backlog==3:
                    counter+=1
            self.my_dict['3']=counter
            self.df=pd.DataFrame(list(self.my_dict.items()))
            self.insert_into_csv(self.df,self.index)
            self.index+=1
    def current_backlogs_4(self):
        self.index=0
        self.my_dict={}
        #print(df["SUB1B"])
        #counting 4 backlogs in each sheet
        for self.df in gb.df_list:
            counter=0
            curback=self.df['CURBACKL']
            for backlog in curback:
                if backlog==4:
                    counter+=1
            self.my_dict['4']=counter
            #Atlast inserting to csv file.
            self.df=pd.DataFrame(list(self.my_dict.items()))
            self.insert_into_csv(self.df,self.index)
            self.index+=1
'''obj=currentBacklogs()
back_0=obj.current_backlogs_0()
back_1=obj.current_backlogs_1()
back_2=obj.current_backlogs_2()
back_3=obj.current_backlogs_3()
back_4=obj.current_backlogs_4()

print("No. of Students with 0 Backlogs:",back_0)
print("No. of Students with 1 Backlogs:",back_1)
print("No. of Students with 2 Backlogs:",back_2)
print("No. of Students with 3 Backlogs:",back_3)
print("No. of Students with 4 Backlogs:",back_4)

'''
