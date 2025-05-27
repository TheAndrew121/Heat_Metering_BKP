import openpyxl
from openpyxl.styles import Font

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Температурный график"

headers = [
    "Температура наружного воздуха, °C",
    "Температура воды в подающем трубопроводе, °C",
    "Температура воды в обратном трубопроводе, °C"
]
ws.append(headers)

data = [
    [-36, 80, 60],
    [-35, 79, 59],
    [-34, 78, 59],
    [-33, 77, 58],
    [-32, 76, 58],
    [-31, 75, 57],
    [-30, 74, 57],
    [-29, 73, 56],
    [-28, 73, 55],
    [-27, 72, 55],
    [-26, 71, 54],
    [-25, 70, 54],
    [-24, 69, 53],
    [-23, 68, 53],
    [-22, 67, 52],
    [-21, 66, 51],
    [-20, 65, 51],
    [-19, 64, 50],
    [-18, 63, 49],
    [-17, 62, 49],
    [-16, 61, 48],
    [-15, 60, 48],
    [-14, 59, 47],
    [-13, 58, 46],
    [-12, 57, 46],
    [-11, 56, 45],
    [-10, 55, 44],
    [-9, 54, 44],
    [-8, 53, 43],
    [-7, 51, 42],
    [-6, 50, 41],
    [-5, 49, 41],
    [-4, 48, 40],
    [-3, 47, 39],
    [-2, 46, 39],
    [-1, 45, 38],
    [0, 44, 37],
    [+1, 43, 36],
    [+2, 41, 35],
    [+3, 40, 35],
    [+4, 39, 34],
    [+5, 38, 33],
    [+6, 37, 32],
    [+7, 35, 31],
    [+8, 34, 30]
]

for row in data:
    ws.append(row)

for cell in ws[1]:
    cell.font = Font(bold=True)

for col in ws.columns:
    max_length = 0
    column = col[0].column_letter
    for cell in col:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(str(cell.value))
        except:
            pass
    adjusted_width = (max_length + 2) * 1.2
    ws.column_dimensions[column].width = adjusted_width

wb.save("Температурный график.xlsx")
