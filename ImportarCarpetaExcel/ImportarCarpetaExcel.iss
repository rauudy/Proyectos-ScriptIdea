Sub Main
	Call bulk_import_excel
End Sub

Function bulk_import_excel
	' Define variables
	Dim folderPath As String
	Dim fileName As String
	Dim fileList() As String
	Dim fileCount As Integer
	Dim i As Integer
	Dim sheetNames As String
	Dim excelApp As Object
	Dim wb As Object
	Dim ws As Object
	
	' Set the folder path
	folderPath = "C:\prueba\" ' <------------------------------------------------------------------------------------------------------ Reemplaza con la ruta de tu carpeta, debe terminar con \
	
	' Ensure the path ends with a backslash
	If Right(folderPath, 1) <> "\" Then
		folderPath = folderPath & "\"
	End If
	
	' Initialize variables
	fileCount = 0
	
	' Obtener el primer archivo de la carpeta
	fileName = Dir(folderPath & "*.xlsx")
	
	' Loop through all the files in the folder
	Do While fileName <> ""
		' Increment the file count
		fileCount = fileCount + 1
		
		' Resize the array to accommodate the new file
		ReDim Preserve fileList(1 To fileCount)
		
		' Add the file name to the array
		fileList(fileCount) = fileName
		
		' Get the next file in the folder
		fileName = Dir
	Loop
	
	' Create a new instance of Excel
	Set excelApp = CreateObject("Excel.Application")
	
	' Process each file to get sheet names
	For i = 1 To UBound(fileList)
		' Open the workbook
		Set wb = excelApp.Workbooks.Open(folderPath & fileList(i))
		
		' Initialize the sheetNames string
		sheetNames = ""
		
		'Import Excel
		Set task = Client.GetImportTask("ImportExcel")
		dbName = folderPath & fileList(i)
		task.FileToImport = dbName
		
		' Get the names of the sheets
		For Each ws In wb.Sheets
			task.SheetToImport = ws.name
		Next ws
		
		task.OutputFilePrefix = "ex" & i '<---------------------------------------------------------------------------------------------------- Change to output file
		task.FirstRowIsFieldName = "TRUE"
		task.EmptyNumericFieldAsZero = "TRUE"
		task.PerformTask
		dbName = task.OutputFilePath(fileList(i))
		Set task = Nothing
		
		wb.Close SaveChanges:=False
	
	Next i
	
	' Quit Excel application
	excelApp.Quit
	
	' Release the object variables
	Set ws = Nothing
	Set wb = Nothing
	Set excelApp = Nothing
	
	Client.RefreshFileExplorer
End Function

