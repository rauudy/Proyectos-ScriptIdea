Begin Dialog consultaSQL 318,14,125,221,"Consulas SQL", .NuevoDiálogo
  Text 10,38,84,7, "Base a consultar:", .Text1
  Text 10,68,70,7, "Nombre base resultante:", .Text2
  Text 10,98,40,7, "Columnas:", .Text3
  Text 10,128,40,7, "Condiciones:", .Text4
  OKButton 15,158,85,14, "Aceptar", .OKButton1
  CancelButton 15,178,85,14, "Cancelar", .CancelButton1
  Text 10,10,90,7, "Nombre de la conexion ODBC:", .Text5
  TextBox 10,20,95,10, .TextBox1
  TextBox 10,48,95,10, .TextBox2
  TextBox 10,78,95,10, .TextBox3
  TextBox 10,108,95,10, .TextBox4
  TextBox 10,138,95,10, .TextBox5
End Dialog
' Script sencillo para hacer Bases de Datos como consultas de SQL. 

'1. Crear una conexion ODBC anclada con un proyecto.
'2. Pasar el nombre de la conexion
'3. La base de datos de referencia tiene que ser que esté dentro del proyecto anclado del ODBC.
'3. No colocar la extension de las bases.
'4. Aceptar o salir.

Sub Main
	'\ Llamar a la funcion de la consulta
	Call consulta
	'\ Refrescar el explorardor de archivos
	Client.RefreshFileExplorer
End Sub

Function consulta
	'\ Llamar al cuadro de dialogo
	Dim win As consultaSQL
	boton = Dialog(win)
	
	'\ Extraer datos del cuadro de dialogo
	conexion = win.TextBox1
	conODBC = ";DSN=" & conexion & ";"
	dbReferencia = win.TextBox2
	dbRes = win.TextBox3
	dbName = dbRes & ".IMD"
	columna = win.TextBox4
	condicion = win.TextBox5
	
	'\ Opciones de los botones
	If boton = -1 Then
		'\ Importar BD desde el ODBC con instrucciones SQL
		Client.ImportODBCFile "" & Chr(34) & dbref & Chr(34) & "", dbName, FALSE, conODBC, "SELECT " &columna & " FROM " & Chr(34) & dbReferencia & Chr(34) & condicion
		Client.OpenDatabase (dbName)
	Else
		'\ Mensaje de salida
		mb = MsgBox("Saliendo de Consultas SQL", 64, "Consultas SQL")
	End If
End Function
