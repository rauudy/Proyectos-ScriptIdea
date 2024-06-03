Sub Main
	escribirExcel
End Sub

Function escribirExcel
	Set excel = CreateObject("Excel.Application")
	Set libro = excel.workbooks.add
	Set hoja1 = libro.sheets(1)
	
	filename = Client.CommonDialogs.FileExplorer()
	Set db = Client.OpenDatabase(filename)
	Set tabla = db.TableDef
	contC = tabla.Count
	Set registros = db.RecordSet
	contR = registros.Count
	Set registro = registros.ActiveRecord
	Dim siono As Integer
	siono = 1
	
	msg = "Desea agregar manualmente los campos"
	manual = MsgBox(msg, MB_YesNo, "Campos a Excel")
	If manual = IDYES Then
		For i = 1 To contC
			Set campo = tabla.GetFieldAt(i)
			agregar = MsgBox("Desea agregar este campo: " & campo.name, MB_YesNo, "AGG")
			If agregar = IDYES Then
				hoja1.cells(2,1+siono).value = campo.name
				For j = 1 To contR
					registros.GetAt(j)
					If campo.IsNumeric Then
						hoja1.cells(2+j,1+siono).value = registro.GetNumValueAt(i)
					ElseIf campo.IsCharacter Then
						hoja1.cells(2+j,1+siono).value = registro.GetCharValueAt(i)
					ElseIf campo.IsDate Then
						hoja1.cells(2+j,1+siono).value = registro.GetDateValueAt(i)
					End If
				Next j
				siono = siono + 1
			Else
				siono = siono
			End If
		Next i
		
		'SI
		Set inicioE = hoja1.cells(2,2)
		Set finE = hoja1.cells(2,siono)
		Set rangoE = hoja1.range(inicioE,finE)
		rangoE.interior.color = RGB(114, 240, 56)
		rangoE.borders.color = RGB(0,0,0)
		rangoE.font.bold = True
		rangoE.horizontalalignment = -4108
		
		Set inicioT = hoja1.cells(3,2)
		Set finT = hoja1.cells(2+contR,siono)
		Set rangoT = hoja1.range(inicioT,finT)
		rangoT.interior.color = RGB(110, 237, 255)
		rangoT.borders.color = RGB(0,0,0)
		rangoT.horizontalAlignment = -4108
	Else
		For i = 1 To contC
			Set campo = tabla.GetFieldAt(i)
			hoja1.Cells(2,1+i).value = campo.name
			For j = 1 To contR
				registros.GetAt(j)
				If campo.IsNumeric Then
					hoja1.cells(2+j,1+i).value = registro.GetNumValueAt(i)
				ElseIf campo.IsCharacter Then
					hoja1.cells(2+j,1+i).value = registro.GetCharValueAt(i)
				ElseIf campo.IsDate Then
					hoja1.cells(2+j,1+i).value = registro.GetDateValueAt(i)
				End If
			Next j
		Next i
		
		'NO
		Set inicioE = hoja1.cells(2,2)
		Set finE = hoja1.cells(2,1+contC)
		Set rangoE = hoja1.range(inicioE,finE)
		rangoE.interior.color = RGB(114, 240, 56)
		rangoE.borders.color = RGB(0,0,0)
		rangoE.font.bold = True
		rangoE.horizontalAlignment = -4108
		
		Set inicioT = hoja1.cells(3,2)
		Set finT = hoja1.cells(2+contR,1+contC)
		Set rangoT = hoja1.range(inicioT,finT)
		rangoT.interior.color = RGB(110, 237, 255)
		rangoT.borders.color = RGB(0,0,0)
		rangoT.horizontalAlignment = -4108
	End If
	
	Set logo = hoja1.shapes.addpicture("D:\DRR3\Descargas\DFK.png",False,True,0,0,50,30)
	hoja1.cells(2,2).select
	
	hoja1.columns.entirecolumn.autofit
	excel.visible = true

End Function
