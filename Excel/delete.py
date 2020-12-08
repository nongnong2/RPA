from openpyxl import load_workbook

wb = load_workbook("sample4.xlsx")
ws = wb.active

# ws.delete_rows(8) #8번째줄 데이터 삭제
# ws.delete_rows(8, 3) #8번쨰 줄에 있는 줄부터 총 3줄 삭제 

# ws.delete_cols(2)
# ws.delete_cols(2, 2)

wb.save("sample_delete.xlsx")