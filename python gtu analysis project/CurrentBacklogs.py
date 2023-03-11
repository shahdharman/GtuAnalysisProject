import pandas as pd 
import openpyxl as op
import string
import global_variables as gb
import os
import csv
class CurrentBacklogs:
    def __init__(self):
        #load gtu file
        self.gtu_workbook_path="gtufiles"+"/"+"SEM "+gb.e1.get()+"/"+"SEM"+gb.e1.get()+"_gtufile.xls"
        #load gtu file workbook
        self.gtu_workbook_df=pd.read_excel(self.gtu_workbook_path)
        self.creating_csvs()
    
    #No need for extra folder as it will be a single file inside "OVERALL" folder

    def creating_csvs(self):
        #creating csv file and initialize it with header
        with open("outputfiles"+"/"+"SEM "+gb.e1.get()+"/"+"OVERALL"+"/"+"CurrentBacklogs.txt", 'w', newline='') as self.csvfile:
            # creating a csv writer object  
            self.csvwriter = csv.writer(self.csvfile,delimiter=" ")
            self.header=['CrrBacks','Count']
            self.csvwriter.writerow(self.header)#writes as first row
            print("Csv files created")

    #insert data into created csv outputfile. 
    def insert_into_csv(self,data):
        #output file path
        self.output_file="outputfiles"+"/"+"SEM "+gb.e1.get()+"/"+"OVERALL"+"/"+"CurrentBacklogs.txt"
        data.to_csv(self.output_file, header=None, index=None, sep='\t',mode='a')#The values() function is used to get a Numpy representation of the DataFrame
        print("txt file saved!!")
        
    def current_backlogs_0(self):
        #print(df["SUB1B"])
        self.my_dict={}
        counter=0
        curback=self.gtu_workbook_df['CURBACKL']
        for backlog in curback:
            if backlog==0:
                counter+=1
        self.my_dict['0']=counter
        self.data=pd.DataFrame(list(self.my_dict.items()))
        self.insert_into_csv(self.data)   
    def current_backlogs_1(self):
        self.my_dict={}
        #print(df["SUB1B"])
        counter=0
        curback=self.gtu_workbook_df['CURBACKL']
        for backlog in curback:
            if backlog==1:
                counter+=1
        self.my_dict['1']=counter
        self.data=pd.DataFrame(list(self.my_dict.items()))
        self.insert_into_csv(self.data)
    def current_backlogs_2(self):
        #print(df["SUB1B"])
        self.my_dict={}
        counter=0
        curback=self.gtu_workbook_df['CURBACKL']
        for backlog in curback:
            if backlog==2:
                counter+=1
        self.my_dict['2']=counter
        self.data=pd.DataFrame(list(self.my_dict.items()))
        self.insert_into_csv(self.data)
    def current_backlogs_3(self):
        #print(df["SUB1B"])
        self.my_dict={}
        counter=0
        curback=self.gtu_workbook_df['CURBACKL']
        for backlog in curback:
            if backlog==3:
                counter+=1
        self.my_dict['3']=counter
        self.data=pd.DataFrame(list(self.my_dict.items()))
        self.insert_into_csv(self.data)
    def current_backlogs_4(self):
        #print(df["SUB1B"])
        self.my_dict={}
        counter=0
        curback=self.gtu_workbook_df['CURBACKL']
        for backlog in curback:
            if backlog==4:
                counter+=1
        self.my_dict['4']=counter
        self.data=pd.DataFrame(list(self.my_dict.items()))
        self.insert_into_csv(self.data)
