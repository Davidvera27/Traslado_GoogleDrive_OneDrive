# Traslado_GoogleDrive_OneDrive
 

Copy code
import os
import win32com.client
import os: Esta línea importa el módulo os, que proporciona funciones para interactuar con el sistema operativo.

import win32com.client: Esta línea importa el módulo win32com.client, que permite interactuar con aplicaciones COM (Component Object Model) como Microsoft Excel.

Copy code
excel_path = r"C:\Users\DAVID\Desktop\DAVID\LITIGIO VIRTUAL\Excel\CONTROLADOR BBVA.xls.xlsm"
word_folder = r"C:\Users\DAVID\Desktop\DAVID\LITIGIO VIRTUAL\WORDs"
Se definen dos variables: excel_path y word_folder.
excel_path contiene la ruta de acceso al archivo de Excel con el que se trabajará.
word_folder contiene la ruta de acceso a la carpeta que contiene los archivos de Word con los que se va a interactuar.

Copy code
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True
Se crea una instancia de Excel y se asigna a la variable excel.
win32com.client.Dispatch("Excel.Application") crea una instancia de la aplicación Excel.
excel.Visible = True hace que la instancia de Excel sea visible en la pantalla.

Copy code
wb = excel.Workbooks.Open(excel_path)
sheet = wb.Worksheets("Controlador")
Se abre el libro de Excel especificado por excel_path y se asigna a la variable wb.
Se selecciona la hoja de cálculo llamada "Controlador" y se asigna a la variable sheet.

Copy code
for row in range(2, sheet.Cells(sheet.Rows.Count, 4).End(-4162).Row + 1):
Se inicia un bucle for que iterará sobre las filas de la hoja de cálculo.
range(2, sheet.Cells(sheet.Rows.Count, 4).End(-4162).Row + 1) genera un rango de números desde la fila 2 hasta la última fila de la hoja de cálculo en la columna 4.

Copy code
cliente_name = sheet.Cells(row, 4).Value.strip()
caso_fng = str(sheet.Cells(row, 13).Value).strip() if sheet.Cells(row, 13).Value is not None else ""
archivo_ver = sheet.Cells(row, 15).Value.strip() if sheet.Cells(row, 15).Value is not None else ""
Se extraen los valores de las celdas en las columnas específicas para cada fila y se eliminan los espacios en blanco alrededor de ellos.
El código restante dentro del bucle for realiza diferentes operaciones dependiendo de los valores extraídos en las variables cliente_name, caso_fng y archivo_ver. Básicamente, busca archivos de Word correspondientes y actualiza las celdas en la hoja de cálculo con las rutas de acceso a esos archivos.


Copy code
wb.Save()
wb.Close()
excel.Quit()
Finalmente, se guardan los cambios en el libro de Excel, se cierra el libro y se cierra la instancia de Excel.