import openpyxl

wb = openpyxl.load_workbook('notas_original.xlsx')
sheet = wb.active

for row in range(1, 40):
    for col in range(1, 15):
        cell_value = sheet.cell(row=row, column=col).value
        if cell_value and isinstance(cell_value, str):
            if "inicia" in cell_value.lower() or "culmina" in cell_value.lower() or "trimestre" in cell_value.lower():
                print(f"FOUND DATES Row {row}, Col {col}: {cell_value.encode('utf-8', 'ignore').decode('utf-8')}")
