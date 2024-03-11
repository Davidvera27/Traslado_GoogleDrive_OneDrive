import os
import win32com.client

# Ruta de acceso al archivo Excel
excel_path = r"C:\Users\DAVID\Desktop\DAVID\LITIGIO VIRTUAL\Excel\CONTROLADOR BBVA.xls.xlsm"
# Ruta de acceso a la carpeta con los archivos Word
word_folder = r"C:\Users\DAVID\Desktop\DAVID\LITIGIO VIRTUAL\WORDs"

# Crear una instancia de Excel
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True

# Abrir el libro de Excel
wb = excel.Workbooks.Open(excel_path)
# Seleccionar la hoja que se requiera para trabajar
sheet = wb.Worksheets("Controlador")

# Iterar sobre las filas para procesar los archivos correspondientes
for row in range(2, sheet.Cells(sheet.Rows.Count, 4).End(-4162).Row + 1):
    cliente_name = sheet.Cells(row, 4).Value.strip()
    caso_fng = str(sheet.Cells(row, 13).Value).strip() if sheet.Cells(row, 13).Value is not None else ""
    archivo_ver = sheet.Cells(row, 15).Value.strip() if sheet.Cells(row, 15).Value is not None else ""
    if cliente_name:
        if caso_fng == "FNG":
            # En caso de ser FNG, buscar el archivo asociado en la columna "N"
            word_file_name_fng = sheet.Cells(row, 14).Value.strip() if sheet.Cells(row, 14).Value is not None else ""
            if word_file_name_fng:
                word_file_path_fng = os.path.join(word_folder, f"{word_file_name_fng}.docx")
                if os.path.exists(word_file_path_fng):
                    sheet.Cells(row, 15).Value = word_file_path_fng
                else:
                    print(f"Archivo FNG para '{cliente_name}' no encontrado.")
        elif archivo_ver == "Ver":
            # En caso de tener "Ver", buscar el archivo correspondiente en la columna "N"
            word_file_name_ver = sheet.Cells(row, 14).Value.strip() if sheet.Cells(row, 14).Value is not None else ""
            if word_file_name_ver:
                word_file_path_ver = os.path.join(word_folder, f"{word_file_name_ver}.docx")
                if os.path.exists(word_file_path_ver):
                    sheet.Cells(row, 15).Value = word_file_path_ver
                else:
                    print(f"Archivo 'Ver' para '{cliente_name}' no encontrado.")
        else:
            # Para otros casos, buscar el archivo directamente con el nombre del cliente
            word_file_path = os.path.join(word_folder, f"{cliente_name}.docx")
            if os.path.exists(word_file_path):
                sheet.Cells(row, 15).Value = word_file_path
            else:
                print(f"Archivo para '{cliente_name}' no encontrado.")

# Guardar y cerrar el libro de Excel
wb.Save()
wb.Close()
excel.Quit()
