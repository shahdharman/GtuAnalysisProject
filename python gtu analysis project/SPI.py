import pandas as pd 
import openpyxl as op
import string
import global_variables as gb
import os
import csv
class SPI:
    def __init__(self):
        #load gtu file
        self.gtu_workbook_path="gtufiles"+"/"+"SEM "+gb.e1.get()+"/"+"SEM"+gb.e1.get()+"_gtufile.xls"
        #load gtu file workbook
        self.gtu_workbook_df=pd.read_excel(self.gtu_workbook_path)
        self.creating_csvs()
    
    #No need for extra folder as it will be a single file inside "OVERALL" folder

    def creating_csvs(self):
        #creating csv file and initialize it with header
        with open("outputfiles"+"/"+"SEM "+gb.e1.get()+"/"+"OVERALL"+"/"+"SPI.txt", 'w', newline='') as self.csvfile:
            # creating a csv writer object  
            self.csvwriter = csv.writer(self.csvfile,delimiter="\t")
            self.header=['SPI','Count']
            self.csvwriter.writerow(self.header)#writes as first row
            print("Csv files created")

    #insert data into created csv outputfile. 
    def insert_into_csv(self,data):
        #output file path
        self.output_file="outputfiles"+"/"+"SEM "+gb.e1.get()+"/"+"OVERALL"+"/"+"SPI.txt"
        data.to_csv(self.output_file, header=None, index=None, sep='\t',mode='a')#The values() function is used to get a Numpy representation of the DataFrame
        print("txt file saved!!")
        
    def spi_greaterthan8(self):
        #print(df["SUB1B"])
        self.my_dict={}
        counter=0
        spi_sheet=self.gtu_workbook_df['SPI']
        for spi in spi_sheet:
            if spi>=8:
                counter+=1
        self.my_dict['>8']=counter
        self.data=pd.DataFrame(list(self.my_dict.items()))
        self.insert_into_csv(self.data)
        
    def spi_greaterthan7(self):
        #print(df["SUB1B"])
        self.my_dict={}
        counter=0
        spi_sheet=self.gtu_workbook_df['SPI']
        for spi in spi_sheet:
            if spi>=7:
                counter+=1
        self.my_dict['>7']=counter
        self.data=pd.DataFrame(list(self.my_dict.items()))
        self.insert_into_csv(self.data)
        
    def spi_greaterthan6(self):
        #print(df["SUB1B"])
        self.my_dict={}
        counter=0
        spi_sheet=self.gtu_workbook_df['SPI']
        for spi in spi_sheet:
            if spi>=6:
                counter+=1
        self.my_dict['>6']=counter
        self.data=pd.DataFrame(list(self.my_dict.items()))
        self.insert_into_csv(self.data)
    def spi_greaterthan5(self):
        #print(df["SUB1B"])
        self.my_dict={}
        counter=0
        spi_sheet=self.gtu_workbook_df['SPI']
        for spi in spi_sheet:
            if spi>=5:
                counter+=1
        self.my_dict['>5']=counter
        self.data=pd.DataFrame(list(self.my_dict.items()))
        self.insert_into_csv(self.data)
         
    def spi_greaterthan4(self):
        #print(df["SUB1B"])
        self.my_dict={}
        counter=0
        spi_sheet=self.gtu_workbook_df['SPI']
        for spi in spi_sheet:
            if spi>=4:
                counter+=1
        self.my_dict['>4']=counter
        self.data=pd.DataFrame(list(self.my_dict.items()))
        self.insert_into_csv(self.data)
