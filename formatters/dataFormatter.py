import xlsxwriter 

class DataFormatter:

    def formatDataToFloat(data):
        number = data.replace(".", "")
        number = round(float(number.replace(",", ".")), 2)
        return number
    
    def formatDataToPercentage(data):
        number = data.replace("%", "")
        number = number.replace(".", "")
        number = round(float(number.replace(",", ".")), 2) / 100
        return number
    
    def formatDataToInteger(data):
        number = int(data.replace(".", ""))
        return number
    
    def formatDataToString(data):
        string = str(data)
        return string
    