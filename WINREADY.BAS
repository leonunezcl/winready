'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Modulo de definicion de variables
' Creado por : MetalBoy
' Fecha : 06/08/96 my birthday day
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'Funciones para las llamadas de windows
Declare Function GetPrivateProfileString Lib "Kernel" (ByVal AppName As String, ByVal KeyName As String, ByVal lpDefault As String, ByVal lpReturnString As String, ByVal nSize As Integer, ByVal IniFileName As String) As Integer
Declare Function GetWindowsDirectory Lib "Kernel" (ByVal p$, ByVal S%) As Integer
Declare Function WritePrivateProfileString Lib "Kernel" (ByVal AppName As String, ByVal KeyName As String, ByVal NewString As String, ByVal FILENAME As String) As Integer
Declare Function WinHelp Lib "User" (ByVal hWnd As Integer, ByVal lpHelpFile As String, ByVal wCommand As Integer, ByVal dwData As Any) As Integer
Declare Function SendMessage Lib "user" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, Lparam As Any) As Long

' get window word constants
 Const GWW_HWNDPARENT = (-8)

' get window constants
 Const GW_HWNDFIRST = 0
 Const GW_HWNDNEXT = 2

' Other constants for listboxes, dialog boxes and windows calls

Global Const WM_USER = &H400
Global Const LB_GETSELCOUNT = WM_USER + 17
Global Const LB_GETCOUNT = WM_USER + 12
Global Const LB_SELITEMRANGE = WM_USER + 28
Global Const WM_SETREDRAW = &HB

'Declaracion para las llamadas al sistema operativo
Global NumTareas, Id, Todos As Integer
Global Const ATTR_DIRECTORY = 16
Global UNIDADES() As String
Global TOT_UNIDADES As Integer

'Menu archivo
Global Const ARCHIVO_COMPRIMIR = 1
Global Const ARCHIVO_DESCOMPRIMIR = 2
Global Const ARCHIVO_CONVERTIR = 4
Global Const ARCHIVO_CREAREXE = 5
Global Const ARCHIVO_LISTAR = 6
Global Const ARCHIVO_MOVER = 7
Global Const ARCHIVO_REPARAR = 8
Global Const ARCHIVO_VERIFICAR = 9
Global Const ARCHIVO_BORRAR = 11
Global Const ARCHIVO_EDITAR = 12
Global Const ARCHIVO_IMPRIMIR = 13
Global Const ARCHIVO_INFORMACION = 14
Global Const ARCHIVO_EJECUCION = 16
Global Const ARCHIVO_MODULO = 17
Global Const ARCHIVO_SALIR = 19

'Menu edicion
Global Const EDICION_AGREGAR_TODOS = 1
Global Const EDICION_QUITAR_TODOS = 2
Global Const EDICION_SELECCION = 4

'Menu opciones
Global Const OPCIONES_CONFIGURACION = 1
Global Const OPCIONES_LISTA_DIR = 2

'Menu utilitarios
Global Const UTILITARIOS_ARJSORT = 1
Global Const UTILITARIOS_RCVT = 2
Global Const UTILITARIOS_REARJ = 3

'Menu ayuda
Global Const AYUDA_INDICE = 1
Global Const AYUDA_BUSCAR = 2
Global Const AYUDA_USAR = 3
Global Const AYUDA_COMO = 5
Global Const AYUDA_REQUERIMIENTOS = 6
Global Const AYUDA_ACERCA = 8

'Ayuda
Global Const HELP_CONTENTS = &H1    'Indice
Global Const HELP_QUIT = &H2        ' Terminar ayuda
Global Const HELP_HELPONHELP = &H4  ' Usando ayuda
Global Const HELP_KEY = &H101       ' Desplegar topico
Global Const HELP_PARTIALKEY = &H105 'Buscar ayuda

' Definicion de variables para los directorios de los programas
Global DIR_DE_ARJ As String
Global DIR_DE_ARJSORT As String
Global DIR_DE_LHA As String
Global DIR_DE_PAK As String
Global DIR_DE_PKZIP As String
Global DIR_DE_PKUNZIP As String
Global DIR_DE_PKZIPFIX As String
Global DIR_DE_RAR As String
Global DIR_DE_RCVT As String
Global DIR_DE_RCVT_CFG As String
Global DIR_DE_REARJ As String
Global DIR_DE_REARJ_CFG As String
Global DIR_DE_ZIP2EXE As String
Global LISTA_DIR As String

'Datos inifile
Global Const INI_FILENAME = "WINREADY.INI"
Global Const KEY_APP_NAME = "WinReady"
Global Const KEY_DIR_ARJ = "Arj"
Global Const KEY_DIR_ARJSORT = "Arjsort"
Global Const KEY_DIR_LHA = "Lha"
Global Const KEY_DIR_PAK = "Pak"
Global Const KEY_DIR_PKZIP = "Pkzip"
Global Const KEY_DIR_PKUNZIP = "Pkunzip"
Global Const KEY_DIR_PKZIPFIX = "Pkzipfix"
Global Const KEY_DIR_RAR = "Rar"
Global Const KEY_DIR_RCVT = "Rcvt"
Global Const KEY_DIR_RCVT_CFG = "Rcvt cfg"
Global Const KEY_DIR_REARJ = "Rearj"
Global Const KEY_DIR_REARJ_CFG = "Rearj cfg"
Global Const KEY_DIR_ZIP2EXE = "Zip2exe"
Global Const KEY_LIS_NAME = "Lista"
Global Const KEY_LISTA_DIR = "Lista Directorios"

' Plantillas
Global PREDEF_ARJ As String
Global PREDEF_LHA As String
Global PREDEF_PAK As String
Global PREDEF_PKZIP As String
Global PREDEF_RAR As String

'tablas de trabajo
Global ARJ_TAB(0 To 7) As String
Global LHA_TAB(0 To 5) As String
Global PAK_TAB(0 To 5) As String
Global PKZIP_TAB(0 To 7) As String
Global RAR_TAB(0 To 6) As String

'Paths
Global PATH_READY As String  'Path de ready
Global PATH_EXC As String    'Path exclusion
Global PATH_HELP As String   'Path archivo ayuda
Global PATH_ERROR As String  'Path archivo de errores
Global PATH_LIST As String   'Path listados
Global PATH_FILES As String  'Path archivos
Global PATH_BAT As String    'Path de ejecución
Global PATH_EDITAR As String 'Path edición
Global PATH_DIRS As String   'Path lista de directorios


'Datos Varios
Global Plantilla_bat(7) As String
Global cSuge(0 To 13) As String
Global cUsar(0 To 5) As String
Global CONTRASENA As String
Global COMENTARIO As String
Global CRLF, Msg As String
Global FileJob As String
Global COMPRESS_ACTIVE As String
Global ACTUAL As String
Global MsgError, MsgOk As String
Global Start As String
Global Origen As String
Global Conversor As String
Global Ruta As String
Global Ext As String
Global PASSWORD As Integer
Global MULTIPLE As Integer
Global SW_COMPRIMIR As Integer
Global F_COMPRIMIR As Integer
Global nSELCOMP As Integer
Global Paso As String

Sub Acoplar (ByVal FirstLine As String, MsgErr As String, MsgOk As String)
    Dim i As Integer
    Open PATH_BAT For Output As #1
	Print #1, "@echo off"
	Print #1, FirstLine
	Plantilla_bat(2) = Plantilla_bat(2) + MsgErr
	Plantilla_bat(5) = Plantilla_bat(5) + MsgOk
	For i = 1 To 7: Print #1, Plantilla_bat(i): Next i
    Close #1
    Plantilla_bat(2) = "echo "
    Plantilla_bat(5) = "echo "
End Sub

Sub Actualizar (Path)
    Dim Count, D(), i, DirName
    Dim PathFind As String
    screen.MousePointer = 11
    DirName = Dir(Path, ATTR_DIRECTORY)
    
    Do While DirName <> ""
	If DirName <> "." And DirName <> ".." Then
	    If GetAttr(Path + DirName) = ATTR_DIRECTORY Then
		If (Count Mod 10) = 0 Then
		    ReDim Preserve D(Count + 10)
		End If
		Count = Count + 1
		D(Count) = LCase(DirName)
	    End If
	End If
	DirName = Dir
    Loop
    
    For i = 1 To Count
	PathFind = Path & D(i)
	If Right(PathFind, 1) <> "\" Then PathFind = PathFind + "\"
	flistadir.List1.AddItem PathFind
	flistadir.List1.Refresh
	Actualizar Path & D(i) & "\"
    Next i
    screen.MousePointer = 0
End Sub

Sub Addlist (Top As Integer, Down As Integer, xlist1 As ListBox, xlist2 As MListBox)
    Dim i As Single
    
    Select Case COMPRESS_ACTIVE
	Case "Arj"
	For i = Top To xlist1.ListCount - Down
	    xlist2.AddItem LCase(Left(xlist1.List(i), 12))
	Next i
	Case "Lha"
	For i = Top To xlist1.ListCount - Down
	    xlist2.AddItem LCase(Mid(xlist1.List(i), 3, 12))
	Next i
	Case "Pak"
	For i = Top To xlist1.ListCount - Down
	    xlist2.AddItem LCase(Mid(xlist1.List(i), 9, 12))
	Next i
	Case "Pkzip", "Pkunzip"
	For i = Top To xlist1.ListCount - Down
	    xlist2.AddItem LCase(Mid(xlist1.List(i), 48, 12))
	Next i
	Case "Rar"
	For i = Top To xlist1.ListCount - Down
	    xlist2.AddItem LCase(Mid(xlist1.List(i), 2, 12))
	Next i
    End Select
    xlist1.Clear
    xlist2.ListIndex = -1
End Sub

Sub AddXlist (xlist1 As ListBox, xlist2 As MListBox, Key As Integer)
    Dim i As Single
    
    Select Case Key
	Case 0
	For i = 0 To xlist1.ListCount - 1: xlist2.AddItem xlist1.List(i): Next i
	Case 1
	For i = 0 To xlist1.ListCount - 1
	    If xlist1.Selected(i) Then xlist2.AddItem xlist1.List(i)
	Next i
    End Select
End Sub

Sub CentrarForma (NombreForma As Form)
    NombreForma.Left = (screen.Width - NombreForma.Width) / 2
    NombreForma.Top = (screen.Height - NombreForma.Height) / 2
End Sub

Sub ChequearSetup ()
    
    'Leer los directorios de los programas
    DIR_DE_ARJ = LeerSetupIniFile(KEY_APP_NAME, KEY_DIR_ARJ)
    DIR_DE_ARJSORT = LeerSetupIniFile(KEY_APP_NAME, KEY_DIR_ARJSORT)
    DIR_DE_LHA = LeerSetupIniFile(KEY_APP_NAME, KEY_DIR_LHA)
    DIR_DE_PAK = LeerSetupIniFile(KEY_APP_NAME, KEY_DIR_PAK)
    DIR_DE_PKZIP = LeerSetupIniFile(KEY_APP_NAME, KEY_DIR_PKZIP)
    DIR_DE_PKUNZIP = LeerSetupIniFile(KEY_APP_NAME, KEY_DIR_PKUNZIP)
    DIR_DE_PKZIPFIX = LeerSetupIniFile(KEY_APP_NAME, KEY_DIR_PKZIPFIX)
    DIR_DE_RAR = LeerSetupIniFile(KEY_APP_NAME, KEY_DIR_RAR)
    DIR_DE_RCVT = LeerSetupIniFile(KEY_APP_NAME, KEY_DIR_RCVT)
    DIR_DE_RCVT_CFG = LeerSetupIniFile(KEY_APP_NAME, KEY_DIR_RCVT_CFG)
    DIR_DE_REARJ = LeerSetupIniFile(KEY_APP_NAME, KEY_DIR_REARJ)
    DIR_DE_REARJ_CFG = LeerSetupIniFile(KEY_APP_NAME, KEY_DIR_REARJ_CFG)
    DIR_DE_ZIP2EXE = LeerSetupIniFile(KEY_APP_NAME, KEY_DIR_ZIP2EXE)
    LISTA_DIR = LeerSetupIniFile(KEY_LIS_NAME, KEY_LISTA_DIR)
   
    FillFiles
End Sub

Sub EncenderMenu ()
	winjob.mArchivo_opcion(ARCHIVO_COMPRIMIR).Enabled = IIf(SW_COMPRIMIR = -1, -1, 0)
	winjob.mArchivo_opcion(ARCHIVO_DESCOMPRIMIR).Enabled = IIf(SW_COMPRIMIR = -1 Or DIR_DE_PKUNZIP <> "", -1, 0)
	winjob.mArchivo_opcion(ARCHIVO_CREAREXE).Enabled = IIf(SW_COMPRIMIR = -1 Or DIR_DE_ZIP2EXE <> "", -1, 0)
	winjob.mArchivo_opcion(ARCHIVO_CONVERTIR).Enabled = IIf(DIR_DE_RCVT <> "" And DIR_DE_RCVT_CFG <> "" Or DIR_DE_REARJ <> "" And DIR_DE_REARJ_CFG <> "", -1, 0)
	winjob.mArchivo_opcion(ARCHIVO_LISTAR).Enabled = IIf(SW_COMPRIMIR = -1 Or DIR_DE_ZIP2EXE <> "", -1, 0)
	winjob.mArchivo_opcion(ARCHIVO_MOVER).Enabled = IIf(SW_COMPRIMIR = -1, -1, 0)
	winjob.mArchivo_opcion(ARCHIVO_REPARAR).Enabled = IIf(DIR_DE_PKZIPFIX <> "" Or DIR_DE_RAR <> "", -1, 0): SW_REPARAR = IIf(DIR_DE_PKZIPFIX <> "" Or DIR_DE_RAR <> "", -1, 0)
	winjob.mArchivo_opcion(ARCHIVO_VERIFICAR).Enabled = IIf(SW_COMPRIMIR = -1, -1, 0)
	winjob.mArchivo_opcion(ARCHIVO_BORRAR).Enabled = IIf(SW_COMPRIMIR = -1, -1, 0)
	winjob.mArchivo_opcion(ARCHIVO_IMPRIMIR).Enabled = IIf(SW_COMPRIMIR = -1, -1, 0)
	winjob.mArchivo_opcion(ARCHIVO_EDITAR).Enabled = -1
	winjob.mArchivo_opcion(ARCHIVO_INFORMACION).Enabled = -1
	winjob.mArchivo_opcion(ARCHIVO_EJECUCION).Enabled = -1
	'winjob.mArchivo_opcion(ARCHIVO_MODULO).Enabled = IIf(DIR_DE_RAR <> "", -1, 0)
    
	'Menú edicion
	winjob.mEdicion.Enabled = IIf(SW_COMPRIMIR = -1, -1, 0)
	winjob.mEdicion_opcion(EDICION_SELECCION).Enabled = IIf(SW_COMPRIMIR = -1, -1, 0)
	
	'Menú opciones
	winjob.mOpcion_opcion(OPCIONES_CONFIGURACION).Enabled = -1
	winjob.mOpcion_opcion(OPCIONES_LISTA_DIR).Checked = Val(LISTA_DIR)
    
	'Menú utilitarios
	winjob.mUtilitarios.Enabled = IIf(DIR_DE_ARJSORT <> "" Or DIR_DE_RCVT <> "" Or DIR_DE_REARJ <> "", -1, 0)
	winjob.mUtilitarios_opcion(UTILITARIOS_ARJSORT).Enabled = IIf(DIR_DE_ARJSORT <> "" And DIR_DE_ARJ <> "", -1, 0)
	winjob.mUtilitarios_opcion(UTILITARIOS_RCVT).Enabled = IIf(DIR_DE_RCVT <> "" And DIR_DE_RCVT_CFG <> "", -1, 0)
	winjob.mUtilitarios_opcion(UTILITARIOS_REARJ).Enabled = IIf(DIR_DE_REARJ <> "" And DIR_DE_REARJ_CFG <> "", -1, 0)
	
	'Menú ayuda
	winjob.mAyuda.Enabled = IIf(PATH_HELP <> "", -1, 0)
    
	'Pictures menues
	winjob.ButtonBar1.ButtonEnabled(0) = IIf(SW_COMPRIMIR = -1, -1, 0)
	winjob.ButtonBar1.ButtonEnabled(1) = IIf(SW_COMPRIMIR = -1 Or DIR_DE_PKUNZIP <> "", -1, 0)
	winjob.ButtonBar1.ButtonEnabled(2) = IIf(SW_COMPRIMIR = -1 Or DIR_DE_ZIP2EXE <> "", -1, 0)
	winjob.ButtonBar1.ButtonEnabled(3) = IIf(SW_COMPRIMIR = -1, -1, 0)
	winjob.ButtonBar1.ButtonEnabled(4) = IIf(SW_COMPRIMIR = -1, -1, 0)
	winjob.ButtonBar1.ButtonEnabled(5) = -1
	winjob.ButtonBar1.ButtonEnabled(6) = IIf(SW_COMPRIMIR = -1, -1, 0)
	
	'Pictures edicion
	winjob.ButtonBar1.ButtonEnabled(7) = IIf(SW_COMPRIMIR = -1, -1, 0)
	winjob.ButtonBar1.ButtonEnabled(8) = IIf(SW_COMPRIMIR = -1, -1, 0)
	winjob.ButtonBar1.ButtonEnabled(9) = IIf(SW_COMPRIMIR = -1, -1, 0)
	
	'Combo box
	winjob.Usar.Enabled = IIf(SW_COMPRIMIR = -1 Or DIR_DE_PKUNZIP <> "", -1, 0)
End Sub

Sub FillFile (ListToFill As MListBox)
    Dim i As Single
    Dim Ruta As String
    Ruta = Path(winjob.Dir1)
    Open PATH_FILES For Output As #1
	If MULTIPLE = 0 Then
	    For i = 0 To ListToFill.ListCount - 1: Print #1, Ruta + ListToFill.List(i): Next i
	Else
	    For i = 0 To ListToFill.ListCount - 1: Print #1, ListToFill.List(i): Next i
	End If
    Close #1
End Sub

Sub FillFiles ()
    
    'Verificar compresores
    If DIR_DE_ARJ <> "" Then winjob.Usar.AddItem cUsar(0)
    If DIR_DE_LHA <> "" Then winjob.Usar.AddItem cUsar(1)
    If DIR_DE_PAK <> "" Then winjob.Usar.AddItem cUsar(2)
    If DIR_DE_PKZIP <> "" Then winjob.Usar.AddItem cUsar(3)
    If DIR_DE_PKUNZIP <> "" Then winjob.Usar.AddItem cUsar(4)
    If DIR_DE_RAR <> "" Then winjob.Usar.AddItem cUsar(5)
    
    SW_COMPRIMIR = IIf(DIR_DE_ARJ <> "" Or DIR_DE_LHA <> "" Or DIR_DE_PAK <> "" Or DIR_DE_PKZIP <> "" Or DIR_DE_RAR <> "", -1, 0)
    
    EncenderMenu
    For i = 0 To 5
	If winjob.Usar.List(i) <> "" Then winjob.Usar.ListIndex = i: Exit Sub
    Next i
    Msg = "Mensaje de WinReady" + CRLF
    Msg = Msg + CRLF
    Msg = Msg + "No se ha detectado ningún compresor" + CRLF
    Msg = Msg + "configurado previamente." + CRLF
    Msg = Msg + "En el menu opciones, elija Configuración," + CRLF
    Msg = Msg + "después haga click en el botón Buscar." + CRLF
    Msg = Msg + "Esto es necesario para que WinReady trabaje" + CRLF
    Msg = Msg + "normalmente." + CRLF
    MsgBox Msg, 16, "WinReady"
End Sub

Function GetCount (ListBox As Control)
		   
  GetCount = SendMessage(ListBox.hWnd, LB_GETCOUNT, 0, 0)

End Function

Function GetCountF (Filelistbox As Control)
	
    GetCountF = SendMessage(Filelistbox.hWnd, LB_GETCOUNT, 0, 0)

End Function

Function GetFileName (ByVal Namefile As String)
    Dim i As Single
    For i = Len(Namefile) To 1 Step -1
	If Mid(Namefile, i, 1) = "\" Then
	    GetFileName = Mid(Namefile, i + 1, Len(Namefile) - i)
	    Exit Function
	End If
    Next i
    GetFileName = Namefile
End Function

Sub Go (ByVal CommandLine As String, ByVal MsgErr As String, ByVal MsgOk As String)
    On Error Resume Next
    Dim Cerr As String
    
    Acoplar CommandLine, MsgErr, MsgOk
    MsgBox CommandLine
    DoShell PATH_BAT, SW_NORMAL
    ChDir app.Path
    Open PATH_ERROR For Input As #1
	Cerr = LTrim(Input$(LOF(1), 1))
    Close #1
    Kill PATH_ERROR
    Select Case Left(Cerr, 5)
	Case "Error"
	    MsgBox "Ha ocurrido un error en la acción realizada", 16, "WinReady"
	Case Else
	    MsgBox "Exito en la acción realizada", 48, "WinReady"
    End Select
    Kill PATH_BAT
End Sub

Sub GrabarInformacionAIniFile ()
    'Grabar los directorios de los programas
    
    GrabarSetupIniFile KEY_APP_NAME, KEY_DIR_ARJ, DIR_DE_ARJ, INI_FILENAME
    GrabarSetupIniFile KEY_APP_NAME, KEY_DIR_ARJSORT, DIR_DE_ARJSORT, INI_FILENAME
    GrabarSetupIniFile KEY_APP_NAME, KEY_DIR_LHA, DIR_DE_LHA, INI_FILENAME
    GrabarSetupIniFile KEY_APP_NAME, KEY_DIR_PAK, DIR_DE_PAK, INI_FILENAME
    GrabarSetupIniFile KEY_APP_NAME, KEY_DIR_PKZIP, DIR_DE_PKZIP, INI_FILENAME
    GrabarSetupIniFile KEY_APP_NAME, KEY_DIR_PKUNZIP, DIR_DE_PKUNZIP, INI_FILENAME
    GrabarSetupIniFile KEY_APP_NAME, KEY_DIR_PKZIPFIX, DIR_DE_PKZIPFIX, INI_FILENAME
    GrabarSetupIniFile KEY_APP_NAME, KEY_DIR_RAR, DIR_DE_RAR, INI_FILENAME
    GrabarSetupIniFile KEY_APP_NAME, KEY_DIR_RCVT, DIR_DE_RCVT, INI_FILENAME
    GrabarSetupIniFile KEY_APP_NAME, KEY_DIR_RCVT_CFG, DIR_DE_RCVT_CFG, INI_FILENAME
    GrabarSetupIniFile KEY_APP_NAME, KEY_DIR_REARJ, DIR_DE_REARJ, INI_FILENAME
    GrabarSetupIniFile KEY_APP_NAME, KEY_DIR_REARJ_CFG, DIR_DE_REARJ_CFG, INI_FILENAME
    GrabarSetupIniFile KEY_APP_NAME, KEY_DIR_ZIP2EXE, DIR_DE_ZIP2EXE, INI_FILENAME
    GrabarSetupIniFile KEY_LIS_NAME, KEY_LISTA_DIR, LISTA_DIR, INI_FILENAME
End Sub

Sub GrabarSetupIniFile (NombreAp As String, PalabraClave As String, ValorClave As String, FileIni As String)
    Dim temp As Integer
    temp = WritePrivateProfileString(NombreAp, PalabraClave, ValorClave, FileIni)
End Sub

Function IsValidPath (DestPath$, ByVal DefaultDrive$) As Integer
    DestPath$ = RTrim$(LTrim$(DestPath$))
    If Right$(DefaultDrive$, 1) <> ":" Or Len(DefaultDrive$) <> 2 Then
	MsgBox "Error al digitar el path de la unidad especificada.""" + DefaultDrive$ + """.  Debe ser la letra de la unidad y "":"".  Por ejemplo, ""C:"", ""D:""...", 64, "WinReady"
	GoTo parseErr
    End If
	 
    If Left$(DestPath$, 1) = "\" Then DestPath$ = DefaultDrive + DestPath$
    On Error Resume Next
    tmp$ = Dir$(DestPath$)
    If Err <> 0 Then GoTo parseErr
    If (InStr(DestPath$, "*") <> 0) GoTo parseErr
    If (InStr(DestPath$, "?") <> 0) GoTo parseErr
    If (InStr(DestPath$, " ") <> 0) GoTo parseErr
	 
    If Mid$(DestPath$, 2, 1) <> Chr$(58) Then GoTo parseErr
    If Len(DestPath$) > 2 Then
	If Right$(Left$(DestPath$, 3), 1) <> "\" Then
    DestPath$ = Left$(DestPath$, 2) + "\" + Right$(DestPath$, Len(DestPath$) - 2)
	End If
    End If

    drive$ = Left$(DestPath$, 1)
    ChDrive (drive$)                                                        ' Try to change to the dest drive
    If Err <> 0 Then GoTo parseErr
	 
    If Right$(DestPath$, 1) <> "\" Then DestPath$ = DestPath$ + "\"
    If Len(DestPath$) = 3 Then
	If Right$(DestPath$, 2) = ":\" Then GoTo ParseOK
    End If
	 
    If InStr(DestPath$, "\\") <> 0 Then GoTo parseErr
	 legalChar$ = "!#$%&'()-0123456789@ABCDEFGHIJKLMNOPQRSTUVWXYZ^_`{}~.üäöÄÖÜß"
	 BackPos = 3
	 forePos = InStr(4, DestPath$, "\")
	 Do
	temp$ = Mid$(DestPath$, BackPos + 1, forePos - BackPos - 1)
	
	'----------------------------
	' Test for illegal characters
	'----------------------------
	For i = 1 To Len(temp$)
		 If InStr(legalChar$, UCase$(Mid$(temp$, i, 1))) = 0 Then GoTo parseErr
	Next i

	'-------------------------------------------
	' Check combinations of periods and lengths
	'-------------------------------------------
	periodPos = InStr(temp$, ".")
	length = Len(temp$)
	If periodPos = 0 Then
		 If length > 8 Then GoTo parseErr                         ' Base too long
	Else
		 If periodPos > 9 Then GoTo parseErr                      ' Base too long
		 If length > periodPos + 3 Then GoTo parseErr             ' Extension too long
		 If InStr(periodPos + 1, temp$, ".") <> 0 Then GoTo parseErr' Two periods not allowed
	End If

	BackPos = forePos
	forePos = InStr(BackPos + 1, DestPath$, "\")
	 Loop Until forePos = 0

ParseOK:
	 IsValidPath = True
	 Exit Function

parseErr:
	 IsValidPath = False
End Function

Sub LeerArchivo (xlist1 As ListBox)
    Open PATH_LIST For Input As #1
	Do While Not EOF(1)
	    Line Input #1, ReadLine
	    xlist1.AddItem ReadLine
	Loop
    Close #1
    Kill PATH_LIST
End Sub

Function LeerSetupIniFile (NombreAp As String, PalabraClave As String) As String
    Dim CadenaObtenida  As String * 128, temp As Integer
    temp = GetPrivateProfileString(NombreAp, PalabraClave, "", CadenaObtenida, Len(CadenaObtenida), INI_FILENAME)
    LeerSetupIniFile = Left(CadenaObtenida, temp)
End Function

Sub LLenarDirs (combobox As Control)
    Dim Dirleido As String
    Open PATH_DIRS For Input As #1
    Do While Not EOF(1)
	Line Input #1, Dirleido
	combobox.AddItem Dirleido
    Loop
    Close #1
End Sub

Sub LLenarPlantilla ()
    
    PATH_READY = LCase(app.Path)
    If Right(PATH_READY, 1) <> "\" Then PATH_READY = PATH_READY + "\" Else PATH_READY = PATH_READY
    
    PATH_EXC = PATH_READY + "files.exc"
    PATH_ERROR = PATH_READY + "error.err"
    PATH_LIST = PATH_READY + "archivos.lst"
    PATH_FILES = PATH_READY + "files.txt"
    PATH_BAT = PATH_READY + "wready.bat"
    PATH_HELP = PATH_READY + "winready.hlp"
    PATH_DIRS = PATH_READY + "dirs.txt"
    
    'Llenar tablas de trabajo
    ARJ_TAB(0) = ""
    ARJ_TAB(1) = ""
    ARJ_TAB(2) = ""
    ARJ_TAB(3) = ""
    ARJ_TAB(4) = ""
    ARJ_TAB(5) = ""
    ARJ_TAB(6) = ""
    ARJ_TAB(7) = ""
    
    'Varios
    CRLF = Chr(13) + Chr(10)
    
    'Plantillas prefefinidas
    PREDEF_ARJ = " -jm -js -va "
    PREDEF_LHA = " -z2 -zarj -zrar "
    PREDEF_PAK = " /sec /p "
    PREDEF_PKZIP = " -ex "
    PREDEF_RAR = " -s -m5 -ehs "
    
    'Plantilla bat
    Plantilla_bat(1) = "if not errorlevel 1 goto exito"
    Plantilla_bat(2) = "echo "
    Plantilla_bat(3) = "goto end"
    Plantilla_bat(4) = ":exito"
    Plantilla_bat(5) = "echo "
    Plantilla_bat(6) = "goto end"
    Plantilla_bat(7) = ":end"
    
    'Plantilla de compresores
    cUsar(0) = "Arj"
    cUsar(1) = "Lha"
    cUsar(2) = "Pak"
    cUsar(3) = "Pkzip"
    cUsar(4) = "Pkunzip"
    cUsar(5) = "Rar"
    
    'Llenar plantilla sugerencias
    cSuge(0) = "Para seleccionar archivos que no están en forma conjunta mantenga presionada la tecla CTRL y haga clic en el archivo. Esta situación es válida para todas las acciones."
    cSuge(1) = "Para seleccionar un grupo de archivos que están en forma conjunta mantenga presionado el botón izquierdo del mouse y realize la selección."
    cSuge(2) = "Para comprimir archivos a varios discos utilize Arj, Pkzip ó Rar."
    cSuge(3) = "Para comprimir archivos que se encuentran en diferentes directorios seleccione en el menú ""Edición"" la opción ""Selección múltiple""."
    cSuge(4) = "Para convertir un archivo comprimido a formato ARJ o RAR seleccione en el menú ""Archivo"" la opción ""Convertir""."
    cSuge(5) = "Para convertir un archivo comprimido a formato EXE seleccione en el menú ""Archivo"" la opción ""Crear EXE""."
    cSuge(6) = "Para descomprimir un archivo haga doble clic en el."
    cSuge(7) = "El convertir un archivo comprimido a un archivo de extensión ""EXE"", le permite descomprimir este archivo sin necesidad de usar el compresor de origen."
    cSuge(8) = "Si es un usuario avanzado puede usar el compresor de su gusto y seleccionar en el menú ""Archivo"" la opción ""Ejecución directa""."
    cSuge(9) = "Si comprime usando una contraseña anote la contraseña en un lugar seguro junto con el nombre del archivo al cual le fue asignada."
    cSuge(10) = "El comprimir un archivo agregando a este un comentario le permite crear una identificación del programa el cual esta comprimiendo."
End Sub

Sub Msdos ()
    Dim x As Integer
    x = Shell("command.com", 2)
End Sub

Function Path (cDir As DirListBox)
    Path = IIf(Right(cDir, 1) <> "\", cDir + "\", cDir)
End Function

Function QuitSlash (Ruta As String) As String
    Dim i As Integer
    For i = Len(Ruta) To 1 Step -1
	If Mid(Ruta, i, 1) = "\" Then
	    QuitSlash = Left(Ruta, i)
	    Exit Function
	End If
    Next i
End Function

Sub ResetearMenu (KeyIndex As Integer)
    Dim i As Integer
    winjob.Drive1.Enabled = IIf(KeyIndex = 1, 0, -1)
    winjob.txtsearch.Enabled = IIf(KeyIndex = 1, 0, -1)
    DoEvents
    Select Case KeyIndex
	Case 0: EncenderMenu
	Case 1
	    winjob.mArchivo_opcion(2).Enabled = 0
	    For i = 4 To 9: winjob.mArchivo_opcion(i).Enabled = 0: Next i
	    For i = 11 To 14: winjob.mArchivo_opcion(i).Enabled = 0: Next i
	    winjob.mArchivo_opcion(16).Enabled = 0
	    winjob.mEdicion_opcion(EDICION_SELECCION).Enabled = 0
	    winjob.mOpcion_opcion(OPCIONES_CONFIGURACION).Enabled = 0
	    winjob.mUtilitarios.Enabled = 0
	    For i = 1 To 6: winjob.ButtonBar1.ButtonEnabled(i) = 0: Next i
	    winjob.ButtonBar1.ButtonEnabled(9) = 0
    End Select
End Sub

Sub StatusMsg (ByVal Msg As String)
    winjob.StatusBar1.Message(0) = Msg
End Sub

Function Windir () As String
    Dim Winpath As String
    Winpath = String(145, Chr(0))
    Windir = Left(Winpath, GetWindowsDirectory(Winpath, Len(Winpath)))
End Function

