'Cargar librerías
LoadFunctionLibrary "..\..\CommonLibrary\Login.qfl"
LoadFunctionLibrary "..\..\CommonLibrary\Menu.qfl"
LoadFunctionLibrary "..\..\CommonLibrary\Utility.qfl"
LoadFunctionLibrary "..\..\CommonLibrary\oracle.qfl"
LoadFunctionLibrary "..\..\CommonLibrary\paperless.qfl"

' Cargar el archivo Excel
Dim excelName

excelName = Environment("TestDir") & "\..\..\DataTable\DATA_IMPRESION_PAPERLESS_DIG.xlsx"

startExcel(excelName)

Call get_data_from_paperless_DB("DIG", 10)

endExcel(excelName)

' Obtener el número de filas
DataTable.Import excelName
Dim num_polizas
num_polizas = DataTable.GetSheet("Global").GetRowCount

'Cierra, abre y realiza login sobre la aplicación
Call close_and_launch_app()
Call login()
	
Call paperless(num_polizas)

DataTable.Export excelName
'Cerrar la aplicación para empezar siempre de cero
Call close_app()

