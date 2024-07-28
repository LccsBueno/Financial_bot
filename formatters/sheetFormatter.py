import xlsxwriter
import subprocess
import datetime

from . import dataFormatter as dt

class SheetFormatter: 
    
    def __init__(
            self,
            generatedSheetPath,
            qtd_columns,
             **kwargs):
            
        self.excel_path = r"C:/Program Files/Microsoft Office/root/Office16/EXCEL.EXE"
        self.generatedSheetPath = generatedSheetPath
        self.qtd_columns = qtd_columns
        
        self.integerColumns = kwargs.get("integerColumn")
        self.currencyColumns = kwargs.get("currencyColumns")
        self.decimalColumns = kwargs.get("decimalColumns")
        self.percentageColumns = kwargs.get("percentageColumns")
        
        self.workbook = xlsxwriter.Workbook(self.generatedSheetPath)

        self.worksheet = self.workbook.add_worksheet("Ativos") 
        
    def dataHora(self):
        
        bold = self.workbook.add_format()
        bold.set_bold()
        bold.set_font_size(13)
        
        now = datetime.datetime.now()  
        
        self.worksheet.write(1, 1, "Data", bold)
        self.worksheet.write(2, 1, str(now.day)+"/"+str(now.month)+"/"+str(now.year))
        
        self.worksheet.write(1, 2, "Hora", bold)
        self.worksheet.write(2, 2, str(now.hour)+":"+str(now.minute))
    
    def sheetGenerator(self, data):
        
        currency = self.workbook.add_format({'num_format': '\\R$ #,##0.00;[Red]-R$ #,##0.00;"-"'})        
        decimal = self.workbook.add_format({'num_format': '#,##0.00;[Red]-#,##0.00;"-"'})        
        percentage = self.workbook.add_format({'num_format': '0.00,##%;[Red]-0.00,##%;"-"'})
        integer = self.workbook.add_format({'num_format': '0'})
        
        qtd_column = 0
        qtd_row = 1
        
        if self.integerColumns == None: 
            self.integerColumns = [0]
            
        if self.currencyColumns == None: 
            self.currencyColumns = [0]
            
        if self.decimalColumns == None: 
            self.decimalColumns = [0]
            
        if self.percentageColumns == None: 
            self.percentageColumns = [0]

        self.dataHora()
        
        for table_row in data:
            
            for table_data in table_row:
                
                if qtd_column == self.qtd_columns:
                    qtd_column = 0
                    qtd_row+=1
                    
                if qtd_row == 1:
                    
                    bold = self.workbook.add_format()
                    bold.set_bold()
                    bold.set_font_size(13)
                    
                    self.worksheet.autofilter('B5:'+chr(ord('@')+self.qtd_columns)+'5')
                    
                    self.worksheet.write(qtd_row+3, qtd_column+1, table_data, bold)

                
                elif qtd_column in self.currencyColumns and qtd_row >= 2:

                    self.worksheet.write(qtd_row+3, qtd_column+1, table_data, currency)    
                        
                
                elif qtd_column in self.decimalColumns and qtd_row >=2:

                    self.worksheet.write(qtd_row+3, qtd_column+1, table_data, decimal)   
                    
                
                elif qtd_column in self.percentageColumns and qtd_row >= 2:

                    self.worksheet.write(qtd_row+3, qtd_column+1, table_data, percentage)    

                elif (qtd_column in self.integerColumns and qtd_row >= 2) and not type(self.integerColumns) == "NoneType": 

                    self.worksheet.write(qtd_row+3, qtd_column+1, table_data, integer)    
                    
                else:
                    self.worksheet.write(qtd_row+3, qtd_column+1, str(table_data))    
                    # worksheet.write(qtd_row, qtd_column, cell_format)

                qtd_column+=1


        try:
            self.workbook.close()
        except Exception as e:
            print("Permissão negada, não pode fechar o workbook: ",e)
        
        
        subprocess.run([self.excel_path, self.generatedSheetPath])
        