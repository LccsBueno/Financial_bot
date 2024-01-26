import xlsxwriter 

class DataFormatter:
    currency = workbook.add_format({'num_format': '\\R$ #,##0.00'})        
    decimal = workbook.add_format({'num_format': '#,##0.00'})        
    percentage = workbook.add_format({'num_format': '0.00,##%'})
    bold = workbook.add_format()
    bold.set_bold()
    bold.set_font_size(13)

    def formatToFloat(data):
        number = data.replace(".", "")
        number = round(float(number.replace(",", ".")), 2)
        return number
    
    def formatToPercentage(data):
        number = data.replace("%", "")
        number = number.replace(".", "")
        number = round(float(number.replace(",", ".")), 2) / 100
        return number
    
    def formatToInteger(data):
        number = int(data.replace(".", ""))
        return number
    
    def formatToString(data):
        string = str(data)
        return string
    