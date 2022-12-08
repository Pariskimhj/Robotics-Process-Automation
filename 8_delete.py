from openpyxl import load_workbook
wb = load_workbook('sample.xlsx')
ws = wb.active

# �� ����

ws.delete_rows(8) # 8��° �ٿ� �ִ� 7�� �л� ������ ����
ws.delete_rows(8, 3) # 8��° �ٺ��� �� 3�� ����

wb.save('sample_delete_row.xlsx')

# �� ����

ws.delete_cols(2) # 2��° �� (B) ����
ws.delete_cols(2, 2) # 2��° ���κ��� �� 2�� �� ����

wb.save("sample_delete_col.xlsx")