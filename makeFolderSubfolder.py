import os
import openpyxl
import time
import shutil

wb = openpyxl.load_workbook('pastas.xlsx')

for sheet in wb.worksheets:
    path = sheet['A2'].value
    for row in sheet.iter_cols(min_row=2, min_col=2, max_col=2):
        for cell in row:
            os.makedirs(os.path.join(path, str(cell.value)))
            shutil.copy2('C:/Users/Admin/Google Drive/Arquivos/FluxoObraPadrao4.0.xlsm', os.path.join(path, str(cell.value)))
            os.rename(os.path.join(path, str(cell.value), 'FluxoObraPadrao4.0.xlsm'), (os.path.join(path, str(cell.value), str(cell.value) + '.xlsm')))
            # wb1 = openpyxl.load_workbook(os.path.join(path, str(cell.value), str(cell.value) + '.xlsm'), read_only=False, keep_vba=True)
            # sheet2 = wb1['Ficha_Solicitação_Material']
            # sheet2['H5'] = str(cell.value)
            # wb1.save(str(cell.value + '.xlsm'))
            time.sleep(3)
