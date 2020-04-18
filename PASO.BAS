Sub A_Borrar ()
	 Dim i As Single
	 
	 If List1.ListCount < 0 Then Exit Sub
	 Ruta = Cual()
	 Select Case Right(Ruta, 3)
	Case "arj": If Not VerCompresor("Arj") Then RutinaError "Arj", "Borrar": Exit Sub Else Start = DIR_DE_ARJ + " l " + Ruta + " > " + PATH_LIST: Ext = "Arj"
	Case "lzh": If Not VerCompresor("Lha") Then RutinaError "Lha", "Borrar": Exit Sub Else Start = DIR_DE_LHA + " l " + Ruta + " > " + PATH_LIST: Ext = "Lha"
	Case "pak": If Not VerCompresor("Pak") Then RutinaError "Pak", "Borrar": Exit Sub Else Start = DIR_DE_PAK + " t " + Ruta + " > " + PATH_LIST: Ext = "Pak"
	Case "zip": If Not VerCompresor("Pkzip") Then RutinaError "Pkzip", "Borrar": Exit Sub Else Start = DIR_DE_PKZIP + " -vb " + Ruta + " > " + PATH_LIST: Ext = "Pkzip"
	Case "rar": If Not VerCompresor("Rar") Then RutinaError "Rar", "Borrar": Exit Sub Else Start = DIR_DE_RAR + " l " + Ruta + " > " + PATH_LIST: Ext = "Rar"
	Case Else
	ShowErr " borrar "
	Exit Sub
	 End Select
	
	 Colocar Ext
	 Acoplar Start, " 1 > " + PATH_ERROR, " 0 > " + PATH_ERROR
	 MsgBox Start
	 NumTareas = GetNumTasks()
	 
	 Id = Shell(PATH_BAT, 3)
	 Do While GetNumTasks() <> NumTareas
	Id = DoEvents()
	 Loop
	 ChDir app.Path
	 Open PATH_ERROR For Input As #1
	Cerr = LTrim(Input$(LOF(1), 1))
	 Close #1
	 Kill PATH_ERROR
	 COMPRESS_ACTIVE = usar.List(usar.ListIndex)
	 Select Case Cerr
	Case "1"
		 MsgBox "Error en el archivo " + LCase(Ruta), 48, "Ready"
	Case Else
		 ACTUAL = Ruta
		 fborrar.Show 1
	 End Select
End Sub

Sub A_Comprimir ()
	 If usar.Text = "Pkunzip" Then
	Msg = "Error de selección." + CRLF
	Msg = Msg + CRLF
	Msg = Msg + "Pkunzip es un descompresor de archivos" + CRLF
	Msg = Msg + "en formato zip." + CRLF
	Msg = Msg + "Seleccione en la lista Usar otro " + CRLF
	Msg = Msg + "compresor y vuelva a intentar." + CRLF
	MsgBox Msg, 48, "Ready"
	Exit Sub
	 End If
	 If List1.ListCount > 0 Then fcomprimir.Show 1
End Sub

Sub A_Convertir ()
	 If List1.ListCount < 0 Then Exit Sub
	 Ruta = Cual()
	 Select Case Right(Ruta, 3)
	Case "arj": If Not VerCompresor("Arj") Then RutinaError "Arj", "Convertir": Exit Sub
	Case "lzh": If Not VerCompresor("Lha") Then RutinaError "Lha", "Convertir": Exit Sub
	Case "pak": If Not VerCompresor("Pak") Then RutinaError "Pak", "Convertir": Exit Sub
	Case "rar": If Not VerCompresor("Rar") Then RutinaError "Rar", "Convertir": Exit Sub
	Case "zip": : If Not VerCompresor("Pkunzip") Then RutinaError "Pkunzip", "Convertir": Exit Sub
	Case Else
	ShowErr " convertir "
	Exit Sub
	 End Select
	 fconvertir.Show 1
End Sub

Sub A_CrearExe ()
	 If List1.ListCount < 0 Then Exit Sub
	 Ruta = Cual()
	 Select Case Right(Ruta, 3)
	Case "arj": If VerCompresor("Arj") Then Start = DIR_DE_ARJ + " y - je1 " + Ruta Else RutinaError "Arj", "Crear exe": Exit Sub
	Case "lzh": If VerCompresor("Lha") Then Start = DIR_DE_LHA + " s " + Ruta Else RutinaError "Lha", "Crear exe": Exit Sub
	Case "pak": If VerCompresor("Pak") Then Start = DIR_DE_PAK + " c /exe " + Ruta Else RutinaError "Pak", "Crear exe": Exit Sub
	Case "rar": If VerCompresor("Rar") Then Start = DIR_DE_RAR + " s " + Ruta Else RutinaError "Rar", "Crear exe": Exit Sub
	Case "zip": : If DIR_DE_ZIP2EXE <> "" Then Start = DIR_DE_ZIP2EXE + " " + Ruta Else RutinaError "Zip2exe", "Crear exe": Exit Sub
	Case Else
	ShowErr " crear exe "
	Exit Sub
	 End Select
	 MsgError = "Error al crear archivo EXE > " + PATH_ERROR
	 MsgOk = "Exito al crear archivo EXE >" + PATH_ERROR
	 Go Start, MsgError, MsgOk
	 Start = ""
End Sub

Sub A_Descomprimir ()
	 Dim i As Single
	 Dim Archivo As String
	 If List1.ListCount < 0 Then Exit Sub
	 Ruta = Cual()
	 Ext = LCase(Right(Ruta, 3))
	 Select Case Ext
	Case "arj": If Not VerCompresor("Arj") Then RutinaError "Arj", "Descomprimir": Exit Sub:  Else Start = DIR_DE_ARJ + " l " + Ruta + " > " + PATH_LIST: Ext = "Arj"
	Case "lzh": If Not VerCompresor("Lha") Then RutinaError "Lha", "Descomprimir": Exit Sub:  Else Start = DIR_DE_LHA + " l " + Ruta + " > " + PATH_LIST: Ext = "Lha"
	Case "pak": If Not VerCompresor("Pak") Then RutinaError "Pak", "Descomprimir": Exit Sub:  Else Start = DIR_DE_PAK + " t " + Ruta + " > " + PATH_LIST: Ext = "Pak"
	Case "zip": If Not VerCompresor("Pkunzip") Then RutinaError "Pkunzip", "Descomprimir": Exit Sub:  Else Start = DIR_DE_PKUNZIP + " " + Ruta + " -@" + PATH_LIST: Ext = "Pkunzip"
	Case "rar": If Not VerCompresor("Rar") Then RutinaError "Rar", "Descomprimir": Exit Sub:  Else Start = DIR_DE_RAR + " l " + Ruta + " > " + PATH_LIST: Ext = "Rar"
	Case Else
	ShowErr " descomprimir "
	Exit Sub
	 End Select
	 
	 Colocar Ext
	 Acoplar Start, " 1 > " + PATH_ERROR, " 0 > " + PATH_ERROR
	 NumTareas = GetNumTasks()
	 Id = Shell(PATH_BAT, 1)
	 Do While GetNumTasks() <> NumTareas
	Id = DoEvents()
	 Loop
	 ChDir app.Path
	 Open PATH_ERROR For Input As #1
	Cerr = LTrim(Input$(LOF(1), 1))
	 Close #1
	 Kill PATH_ERROR
	 Kill PATH_BAT
	 COMPRESS_ACTIVE = usar.Text
	 Select Case Cerr
	Case "1"
		 MsgBox "Error en el archivo " + LCase(Ruta), 48, "Ready"
	Case Else
		 FileJob = Ruta
		 fdescomprimir.Show 1
	 End Select
End Sub

Sub A_Editar ()
	 If List1.ListCount > 0 Then PATH_EDITAR = Cual(): feditar.Show 1
End Sub

Sub A_EjecucionDirecta ()
	 fejedir.Show 1
End Sub

Sub A_imprimir ()
	 Dim i As Single
	 If List1.ListCount < 0 Then Exit Sub
	 Ruta = Cual()
	 Ext = LCase(Right(Ruta, 3))
	 Select Case Ext
	Case "arj": If Not VerCompresor("Arj") Then RutinaError "Arj", "Imprimir": Exit Sub Else Start = DIR_DE_ARJ + " l " + Ruta + " > " + PATH_LIST: Ext = "Arj"
	Case "lzh": If Not VerCompresor("Lha") Then RutinaError "Lha", "Imprimir": Exit Sub Else Start = DIR_DE_LHA + " l " + Ruta + " > " + PATH_LIST: Ext = "Lha"
	Case "pak": If Not VerCompresor("Pak") Then RutinaError "Pak", "Imprimir": Exit Sub Else Start = DIR_DE_PAK + " t " + Ruta + " > " + PATH_LIST: Ext = "Pak"
	Case "zip": If Not VerCompresor("Pkunzip") Then RutinaError "Pkunzip", "Listar": Exit Sub Else Start = DIR_DE_PKUNZIP + " " + Ruta + " -@" + PATH_LIST: Ext = "Pkunzip"
	Case "rar": If Not VerCompresor("Rar") Then RutinaError "Rar", "Imprimir": Exit Sub Else Start = DIR_DE_RAR + " l " + Ruta + " > " + PATH_LIST: Ext = "Rar"
	 Case Else
	ShowErr " Imprimir "
	Exit Sub
	 End Select
	 
	 Colocar Ext
	 
	 Acoplar Start, " 1 > " + PATH_ERROR, " 0 > " + PATH_ERROR
	 NumTareas = GetNumTasks()
	 Id = Shell(PATH_BAT, 1)
	 Do While GetNumTasks() <> NumTareas
	Id = DoEvents()
	 Loop
	 ChDir app.Path
	 Open PATH_ERROR For Input As #1
	Cerr = LTrim(Input$(LOF(1), 1))
	 Close #1
	 Kill PATH_ERROR
	 COMPRESS_ACTIVE = usar.List(usar.ListIndex)
	 Select Case Cerr
	Case "1"
		 MsgBox "Error en el archivo " + LCase(Ruta), 48, "Ready"
	Case Else
		 FileJob = Ruta: fImprimir.Show 1
	 End Select
End Sub

Sub A_Información ()
	 If List1.ListCount > 0 Then finfor.Show 1
End Sub

Sub A_Listar ()
	 Dim i As Single
	 If List1.ListCount < 0 Then Exit Sub
	 Ruta = Cual()
	 Ext = LCase(Right(Ruta, 3))
	 Select Case Ext
	Case "arj": If Not VerCompresor("Arj") Then RutinaError "Arj", "Listar": Exit Sub Else Start = DIR_DE_ARJ + " l " + Ruta + " > " + PATH_LIST: Ext = "Arj"
	Case "lzh": If Not VerCompresor("Lha") Then RutinaError "Lha", "Listar": Exit Sub Else Start = DIR_DE_LHA + " l " + Ruta + " > " + PATH_LIST: Ext = "Lha"
	Case "pak": If Not VerCompresor("Pak") Then RutinaError "Pak", "Listar": Exit Sub Else Start = DIR_DE_PAK + " l " + Ruta + " > " + PATH_LIST: Ext = "Pak"
	Case "zip": If Not VerCompresor("Pkzip") Then RutinaError "Pkzip", "Listar": Exit Sub Else Start = DIR_DE_PKZIP + " -vb " + Ruta + " > " + PATH_LIST: Ext = "Pkzip"
	Case "rar": If Not VerCompresor("Rar") Then RutinaError "Rar", "Listar": Exit Sub Else Start = DIR_DE_RAR + " l " + Ruta + " > " + PATH_LIST: Ext = "Rar"
	 Case Else
	ShowErr " listar "
	Exit Sub
	 End Select
	 
	 Colocar Ext
	 
	 Acoplar Start, " 1 > " + PATH_ERROR, " 0 > " + PATH_ERROR
	 NumTareas = GetNumTasks()
	 
	 Id = Shell(PATH_BAT, 1)
	 Do While GetNumTasks() <> NumTareas
	Id = DoEvents()
	 Loop
	 ChDir app.Path
	 Open PATH_ERROR For Input As #1
	Cerr = LTrim(Input$(LOF(1), 1))
	 Close #1
	 Kill PATH_ERROR
	 COMPRESS_ACTIVE = usar.List(usar.ListIndex)
	 Select Case Cerr
	Case "1"
		 MsgBox "Error en el archivo " + LCase(Ruta), 48, "Ready"
	Case Else
		 FileJob = Ruta: flistar.Show 1
	 End Select
End Sub

Sub A_Modulo ()
	 fmodulo.Show 1
End Sub

Sub A_Reparar ()
	 If List1.ListCount < 0 Then Exit Sub
	 Ruta = Cual()
	 Select Case Right(Ruta, 3)
	Case "arj", "lzh", "pak"
	Msg = "Extensión no válida" + CRLF
	Msg = Msg + CRLF
	Msg = Msg + "Para reparar archivos la extensión" + CRLF
	Msg = Msg + "debe ser rar o zip." + CRLF
	MsgBox Msg, 48, "Ready"
	Exit Sub
	Case "zip": If DIR_DE_PKZIPFIX <> "" Then Start = DIR_DE_PKZIPFIX + Ruta Else RutinaError "Pkzipfix", "Reparar": Exit Sub
	Case "rar": If VerCompresor("Rar") Then Start = DIR_DE_RAR + " r " + Ruta Else RutinaError "Rar", "Reparar": Exit Sub
	 End Select
	 MsgError = "Error al reparar archivo > " + PATH_ERROR
	 MsgOk = "Exito al reparar archivo >" + PATH_ERROR
	 Go Start, MsgError, MsgOk
	 Start = ""
End Sub

Sub A_Verificar ()
	 If List1.ListCount = -1 Then Exit Sub
	 Ruta = Cual()
	 Select Case Right(Ruta, 3)
	Case "arj": If VerCompresor("Arj") Then Start = DIR_DE_ARJ + " t -jr " + Ruta: Ext = "Arj" Else RutinaError "Arj", "Verificar": Exit Sub
	Case "lzh": If VerCompresor("Lha") Then Start = DIR_DE_LHA + " t " + Ruta: Ext = "Lha" Else RutinaError "Lha", "Verificar": Exit Sub
	Case "pak": If VerCompresor("Pak") Then Start = DIR_DE_PAK + " t " + Ruta: Ext = "Pak" Else RutinaError "Pak", "Verificar": Exit Sub
	Case "rar": If VerCompresor("Rar") Then Start = DIR_DE_RAR + " t " + Ruta: Ext = "Rar" Else RutinaError "Rar", "Verificar": Exit Sub
	Case "zip": : If VerCompresor("Pkunzip") Then Start = DIR_DE_PKUNZIP + " -t " + Ruta: Ext = "Pkunzip" Else RutinaError "Pkunzip", "Verificar": Exit Sub
	Case Else
	ShowErr " verificar "
	Exit Sub
	 End Select
	 Colocar Ext
	 MsgError = "Error al verificar archivo > " + PATH_ERROR
	 MsgOk = "Exito al verificar archivo >" + PATH_ERROR
	 Go Start, MsgError, MsgOk
	 Start = ""
End Sub

Sub Colocar (Who As String)
	 Dim i As Integer
	 For i = 0 To usar.ListCount - 1
	If usar.List(i) = Who Then usar.ListIndex = i: Exit For
	 Next i
End Sub

Function Cual () As String
	 For i = 0 To List1.ListCount - 1
	If List1.Selected(i) Then
		 Ruta = Path(Dir1) + List1.List(i)
		 ACTUAL = Ruta
		 Cual = Ruta: Exit Function
	End If
	 Next i
End Function

Sub Dir1_Change ()
	 Dim i As Single
	 Dim nBytes As Long
	 Dim cPath As String
	 file1.Path = Dir1.Path
	 cPath = Path(Dir1)
	 ChDir Dir1.Path
	 List1.Clear
	 screen.MousePointer = 11
	 lbllisto.Caption = "Espere un momento .... "
	 nTotal = file1.ListCount - 1
	 For i = 0 To nTotal
	List1.AddItem file1.List(i)
	nBytes = nBytes + FileLen(cPath + file1.List(i))
	 Next i
	 screen.MousePointer = 0
	 lbllisto.Caption = "Listo"
	 lbltot.Caption = List1.ListCount
	 lblbytes.Caption = nBytes
	 List1.ListIndex = -1
	 If List1.ListCount > 0 Then List1.Selected(0) = -1
End Sub

Sub Dir1_KeyPress (KeyAscii As Integer)
	 If KeyAscii = 13 Then Dir1.Path = Dir1.List(Dir1.ListIndex): KeyAscii = 0
End Sub

Sub E_AgregarTodos ()
	 Dim i As Single
	 screen.MousePointer = 11
	 If MULTIPLE = 0 Then
	For i = 0 To List1.ListCount - 1: List1.Selected(i) = -1: Next i
	List1.ListIndex = -1
	 Else
	For i = 0 To list3.ListCount - 1: list3.Selected(i) = -1: Next i
	list3.ListIndex = -1
	 End If
	 screen.MousePointer = 0
End Sub

Sub E_QuitarTodos ()
	 Dim i As Single
	 screen.MousePointer = 11
	 If MULTIPLE = 0 Then
	For i = 0 To List1.ListCount - 1: List1.Selected(i) = 0: Next i
	List1.ListIndex = -1
	 Else
	For i = 0 To list3.ListCount - 1: list3.Selected(i) = 0: Next i
	list3.ListIndex = -1
	 End If
	 screen.MousePointer = 0
End Sub

Sub E_Seleccion ()
	 fselmul.Show 1
End Sub

Sub EliminarArchivosDeTrabajo ()
	 On Error Resume Next
	 Kill PATH_EXC
	 Kill PATH_ERROR
	 Kill PATH_LIST
	 Kill PATH_FILES
	 Kill PATH_BAT
	 Kill PATH_DIRS
End Sub

Sub Form_Activate ()
	 If MULTIPLE = -1 Then list3.ListIndex = -1: Exit Sub Else file1.Refresh : Dir1_Change
End Sub

Sub Form_Resize ()
	 fready.WindowState = 2
End Sub

Sub Form_Unload (Cancel As Integer)
	 GrabarInformacionAIniFile
	 EliminarArchivosDeTrabajo
	 End
End Sub

Sub O_ActualizarPosibles ()
	 If mOpciones_Opciones(OPCIONES_LISTA_DIR).Checked = 0 Then
	mOpciones_Opciones(OPCIONES_LISTA_DIR).Checked = -1
	flistadir.Show 1
	 Else
	mOpciones_Opciones(OPCIONES_LISTA_DIR).Checked = 0
	 End If
	 LISTA_DIR = Str(mOpciones_Opciones(OPCIONES_LISTA_DIR).Checked)
End Sub

Sub O_Configuracion ()
	 fsetup.Show 1
End Sub

Sub RutinaError (MsgShow As String, Accion As String)
	 Msg = MsgShow + " no esta instalado." + CRLF
	 Msg = Msg + CRLF
	 Msg = Msg + "En el menú opciones seleccione" + CRLF
	 Msg = Msg + """Configuración"",después escriba el path" + CRLF
	 Msg = Msg + "en el cuadro de texto " + MsgShow + " ó " + CRLF
	 Msg = Msg + "haga click en el botón ""Buscar"" para" + CRLF
	 Msg = Msg + "encontrar en su computador el compresor " + MsgShow + CRLF
	 Msg = Msg + "Esto es necesario para realizar con" + CRLF
	 Msg = Msg + "éxito la acción de " + Accion + "." + CRLF
	 MsgBox Msg, 48, "Win Ready"
End Sub

Sub Salir ()
	 On Error Resume Next
	 GrabarInformacionAIniFile
	 Ret = WinHelp(hWnd, PATH_HELP, HELP_QUIT, 0&)
	 EliminarArchivosDeTrabajo
	 End
End Sub

Sub ShowErr (MsgShow As String)
	 Msg = "Extensión no válida." + CRLF
	 Msg = Msg + CRLF
	 Msg = Msg + "Para" + MsgShow + "archivos la extensión" + CRLF
	 Msg = Msg + "debe ser arj, lzh, pak, rar o zip." + CRLF
	 MsgBox Msg, 48, "Win Ready"
End Sub

Sub SugeHelp (Msg As String)
	 lblsuge.Caption = Msg
End Sub

Sub U_Arjsort ()
	 Dim Copiar, Archivo, Destino As String
	 If List1.ListIndex = -1 Then Exit Sub
	 Ruta = Cual()
	 Select Case Right(Ruta, 3)
	Case "arj": If VerCompresor("Arj") Then farjsort.Show 1 Else RutinaError "Arj", "Ordenar": Exit Sub
	Case Else
		 Msg = "Extensión no válida" + CRLF
		 Msg = Msg + CRLF
		 Msg = Msg + "Para sortear archivos la extensión" + CRLF
		 Msg = Msg + "debe ser arj." + CRLF
		 MsgBox Msg, 48, "Win Ready"
		 Exit Sub
	 End Select
End Sub

Sub U_Rcvt ()
	 Dim Archivo, Destino, JobPath As String
	 If List1.ListIndex = -1 Then Exit Sub
	 Ruta = Cual()
	 MsgError = "Error en la conversión > " + PATH_ERROR
	 MsgOk = "Exito en la conversión >" + PATH_ERROR
	 Select Case Right(Ruta, 3)
	Case "arj": JobPath = DIR_DE_ARJ: If Not VerCompresor("Arj") Then RutinaError "Arj", "Convertir": Exit Sub
	Case "lzh": JobPath = DIR_DE_LHA: If Not VerCompresor("Lha") Then RutinaError "Lha", "Convertir": Exit Sub
	Case "rar": JobPath = DIR_DE_RAR: If Not VerCompresor("Rar") Then RutinaError "Rar", "Convertir": Exit Sub
	Case "zip": JobPath = DIR_DE_PKUNZIP: If Not VerCompresor("Pkunzip") Then RutinaError "Pkunzip", "Convertir": Exit Sub
	Case Else
		 Msg = "Extensión no válida" + CRLF
		 Msg = Msg + CRLF
		 Msg = Msg + "Para convertir archivos la extensión" + CRLF
		 Msg = Msg + "debe ser arj,lzh,rar o zip." + CRLF
		 MsgBox Msg, 48, "Win Ready"
		 Exit Sub
	 End Select
	 Origen = QuitSlash(DIR_DE_RCVT)
	 Archivo = GetFileName(JobPath)
	 Destino = Origen + Archivo
	 
	 Start = DIR_DE_RCVT + " " + Ruta
	 Open Destino For Binary As #1
	 If LOF(1) = 0 Then
	FileCopy JobPath, Destino
	Go Start, MsgError, MsgOk
	Close #1
	Kill Destino
	 Else
	Close #1
	Go Start, MsgError, MsgOk
	 End If
End Sub

Sub U_Rearj ()
	 Dim Archivo, Destino, JobPath As String
	 If List1.ListIndex = -1 Then Exit Sub
	 Ruta = Cual()
	 MsgError = "Error en la conversión > " + PATH_ERROR
	 MsgOk = "Exito en la conversión >" + PATH_ERROR
	 Select Case Right(Ruta, 3)
	Case "arj": JobPath = DIR_DE_ARJ: If Not VerCompresor("Arj") Then RutinaError "Arj", "Convertir": Exit Sub
	Case "lzh": JobPath = DIR_DE_LHA: If Not VerCompresor("Lha") Then RutinaError "Lha", "Convertir": Exit Sub
	Case "pak": JobPath = DIR_DE_PAK: If Not VerCompresor("Pak") Then RutinaError "Pak", "Convertir": Exit Sub
	Case "rar": JobPath = DIR_DE_RAR: If Not VerCompresor("Rar") Then RutinaError "Rar", "Convertir": Exit Sub
	Case "zip": JobPath = DIR_DE_PKUNZIP: If Not VerCompresor("Pkunzip") Then RutinaError "Pkunzip", "Convertir": Exit Sub
	Case Else
		 Msg = "Extensión no válida" + CRLF
		 Msg = Msg + CRLF
		 Msg = Msg + "Para convertir archivos la extensión" + CRLF
		 Msg = Msg + "debe ser arj,lzh,pak,rar o zip." + CRLF
		 MsgBox Msg, 48, "Win Ready"
		 Exit Sub
	 End Select
	 Origen = QuitSlash(DIR_DE_REARJ)
	 Archivo = GetFileName(JobPath)
	 Destino = Origen + Archivo
	 Start = DIR_DE_REARJ + " " + Ruta
	 Open Destino For Binary As #1
	 If LOF(1) = 0 Then
	Close #1
	FileCopy JobPath, Destino
	Go Start, MsgError, MsgOk
	Kill Destino
	 Else
	Close #1
	Go Start, MsgError, MsgOk
	 End If
End Sub

Function VerCompresor (Cual As String)
	 Dim i As Integer

	 For i = 0 To usar.ListCount - 1
	If usar.List(i) = Cual Then
		 VerCompresor = True
		 COMPRESS_ACTIVE = usar.List(usar.ListIndex)
		 Exit Function
	End If
	 Next i
	 VerCompresor = False
End Function
