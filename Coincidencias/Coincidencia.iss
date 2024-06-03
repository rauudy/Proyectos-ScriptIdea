Dim camposNumericos() AS string
Dim todosCampos() AS string

Begin Dialog Coincidencia 195,48,248,136,"Coincidencia", .NuevoDiálogo
  Text 15,13,40,7, "Archivo:", .Text1
  TextBox 15,22,115,10, .TextBox1
  PushButton 140,17,90,16, "Buscar Base", .PushButton1
  DropListBox 15,59,115,10, camposNumericos(), .DropListBox1
  DropListBox 15,94,115,10, todosCampos(), .DropListBox2
  Text 15,50,60,7, "Campo a coincidir:", .Text2
  Text 141,49,60,7, "Valor a coincidir:", .Text3
  TextBox 140,58,90,10, .TextBox2
  Text 16,84,80,7, "Campo con valor de retorno:", .Text4
  OKButton 140,91,40,14, "Aceptar", .OKButton1
  CancelButton 190,91,40,14, "Cancelar", .CancelButton1
End Dialog
Sub Main
	Coincidir
	flag = true
	Do While flag
		msg = MsgBox("¿Volver a ejecutar?", MB_YESNO, "Salir")			
		If msg = IDYES Then
			Coincidir
		Else
			MsgBox "ADIOS"
			flag = false
		End If
	Loop
End Sub

Function Coincidir 
	Dim ventana As Coincidencia
	bandera = true
	btn = Dialog(ventana)
	
	Do 
		If btn = 1 Then
				filename = Client.CommonDialogs.FileExplorer()
		ElseIf btn = -1 Then
			MsgBox "Seleccione una base de datos antes"
		Else
			MsgBox "Saliendo"
		End If
	Loop Until filename <> ""
	
	Set db = Client.OpenDatabase(filename)
	Set tabla = db.TableDef
	Set registros = db.RecordSet
	Set registro = registros.ActiveRecord
	contRegistros = registros.Count
	numCampos = tabla.Count
	ReDim camposNumericos(1)
	ReDim todosCampos(1)
	contador = 0
	conn = 0

	For i = 1 To numCampos
		Set campos = tabla.GetFieldAt(i)
		todosCampos(conn) = campos.name
		conn = conn + 1
		ReDim preserve todosCampos(conn)
		
		If campos.IsNumeric Then
			camposNumericos(contador) = campos.name
			contador = contador + 1
			ReDim preserve camposNumericos(contador)
		End If	
	Next i
	
	contador = contador - 1
	conn = conn - 1
	ventana.TextBox1 = filename
	btn = Dialog(ventana)
	selectCampoNum = ventana.DropListBox1
	selectCampoResul = ventana.DropListBox2
	coincidir = ventana.TextBox2
	
	If btn = -1 Then
		For i = 1 To contRegistros
			registros.GetAt(i)
			If registro.GetNumValue(camposNumericos(selectCampoNum)) = coincidir Then
				registros.GetAt(i)
				MsgBox registro.GetCharValue(todosCampos(selectCampoResul)) 
			End If
		Next i
	ElseIf btn = 0 Then
		MsgBox "Saliendo"
	End If
End Function

