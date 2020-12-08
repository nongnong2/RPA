from openpyxl import load_workbook
wb = load_workbook("sample4.xlsx")
ws = wb.active

# ws.insert_rows(8) #8번째 줄이 비워짐
# ws.insert_rows(8,5) #8번째 줄에 5줄이 추가됨

# ws.insert_cols(2) #2번쨰 열이 비워짐
# ws.insert_cols(2,3) #2번째 열로부터 3열 비워짐 

wb.save("sample_insert.xlsx")