import pandas as pd 
import openpyxl as op
import string
import global_variables as gb
import os
import csv
class DivWiseSPI:
    def __init__(self):
        #load input file workbook
        self.gtu_workbook_path="inputfiles"+"/"+"SEM "+gb.e1.get()+"/"+"inputfile_SEM"+gb.e1.get()+".xlsx"
        self.gtu_workbook=op.load_workbook(self.gtu_workbook_path)
        #gb.df_list and gb.div_list are global varaible and already initialized in Div_wise_result file,so no need to define again as they  are defined in global_variables.
        self.sub_folders()
        self.creating_csvs()
        #print(gb.div_list)
        #print(gb.df_list)
    def sub_folders(self):
        self.path1=os.path.join("outputfiles"+"/"+"SEM "+gb.e1.get()+'/'+"Division Wise","SPI")
        try:
            os.mkdir(self.path1)
            print("subfolder created")
        except:
            print("Error!creating subfolder!!")

    def creating_csvs(self):
        for self.each in range(len(gb.div_list)):
            with open("outputfiles"+"/"+"SEM "+gb.e1.get()+'/'+"Division Wise"+"/"+"SPI"+"/"+gb.div_list[self.each]+"_SPI.txt", 'w', newline='') as self.csvfile:
                # creating a csv writer object  
                self.csvwriter = csv.writer(self.csvfile,delimiter="\t")
                self.header=['Greater than ','Count']
                self.csvwriter.writerow(self.header)#writes as first row
        print("Csv files created")
        
    def insert_into_csv(self,data,index):
        #output file path
        self.output_file="outputfiles"+"/"+"SEM "+gb.e1.get()+'/'+"Division Wise"+"/"+"SPI"+"/"+gb.div_list[index]+"_SPI.txt"
        print(self.output_file)
        data.to_csv(self.output_file,header=None, index=None, sep='\t',mode='a')#The values() function is used to get a Numpy representation of the DataFrame
        print("txt file saved!!")

    def spi_greaterthan8(self):
        self.index=0
        self.my_dict={}
        #print(df["SUB1B"])
        for self.df in gb.df_list:
            counter=0
            spi_sheet=self.df['SPI']
            for spi in spi_sheet:
                if spi>=8:
                    counter+=1
            self.my_dict['>8']=counter
            self.df=pd.DataFrame(list(self.my_dict.items()))
            self.insert_into_csv(self.df,self.index)
            self.index+=1
    def spi_greaterthan7(self):
        self.index=0
        self.my_dict={}
        #print(df["SUB1B"])
        for self.df in gb.df_list:
            counter=0
            spi_sheet=self.df['SPI']
            for spi in spi_sheet:
                if spi>=7:
                    counter+=1
            self.my_dict['>7']=counter
            self.df=pd.DataFrame(list(self.my_dict.items()))
            self.insert_into_csv(self.df,self.index)
            self.index+=1
        
    def spi_greaterthan6(self):
        #print(df["SUB1B"])
        self.index=0
        self.my_dict={}
        #print(df["SUB1B"])
        for self.df in gb.df_list:
            counter=0
            spi_sheet=self.df['SPI']
            for spi in spi_sheet:
                if spi>=6:
                    counter+=1
            self.my_dict['>6']=counter
            self.df=pd.DataFrame(list(self.my_dict.items()))
            self.insert_into_csv(self.df,self.index)
            self.index+=1
    def spi_greaterthan5(self):
        #print(df["SUB1B"])
        self.index=0
        self.my_dict={}
        #print(df["SUB1B"])
        for self.df in gb.df_list:
            counter=0
            spi_sheet=self.df['SPI']
            for spi in spi_sheet:
                if spi>=5:
                    counter+=1
            self.my_dict['>5']=counter
            self.df=pd.DataFrame(list(self.my_dict.items()))
            self.insert_into_csv(self.df,self.index)
            self.index+=1
         
    def spi_greaterthan4(self):
        #print(df["SUB1B"])
        self.index=0
        self.my_dict={}
        #print(df["SUB1B"])
        for self.df in gb.df_list:
            counter=0
            spi_sheet=self.df['SPI']
            for spi in spi_sheet:
                if spi>=4:
                    counter+=1
            self.my_dict['>4']=counter
            self.df=pd.DataFrame(list(self.my_dict.items()))
            self.insert_into_csv(self.df,self.index)
            self.index+=1
'''             
obj=SPI()
spi_4=obj.spi_greaterthan4()
spi_5=obj.spi_greaterthan5()
spi_6=obj.spi_greaterthan6()
spi_7=obj.spi_greaterthan7()
spi_8=obj.spi_greaterthan8()
'''
