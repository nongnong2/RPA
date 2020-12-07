from openpyxl import Workbook
wb = Workbook() #create new Workbook
ws = wb.active #활성화된 sheet 가져오기 
ws.title = "NadoSheet" #sheet 이름 설정 
wb.save("sample.xlsx")
wb.close()