VERSION 5.00
Begin VB.MDIForm WinJob 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   ClientHeight    =   4020
   ClientLeft      =   420
   ClientTop       =   2175
   ClientWidth     =   9600
   HelpContextID   =   1
   Icon            =   "WINJOB.frx":0000
   LinkTopic       =   "MDIForm1"
   Tag             =   "WinReady"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Panel3D 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4575
      Index           =   1
      Left            =   0
      ScaleHeight     =   4545
      ScaleWidth      =   9570
      TabIndex        =   11
      Top             =   1710
      Width           =   9600
      Begin VB.ListBox List3 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4320
         Left            =   50
         MultiSelect     =   2  'Extended
         TabIndex        =   7
         Top             =   120
         Visible         =   0   'False
         Width           =   9500
      End
      Begin VB.FileListBox File1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1080
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Columns         =   5
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4320
         Left            =   2280
         MultiSelect     =   2  'Extended
         TabIndex        =   3
         Top             =   120
         Width           =   7215
      End
      Begin VB.DirListBox Dir1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4305
         Left            =   50
         TabIndex        =   2
         Top             =   120
         Width           =   2130
      End
   End
   Begin VB.PictureBox Panel3D 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   765
      Index           =   0
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   9570
      TabIndex        =   9
      Top             =   945
      Width           =   9600
      Begin VB.PictureBox Panel3D 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   675
         Index           =   4
         Left            =   10
         ScaleHeight     =   645
         ScaleWidth      =   9210
         TabIndex        =   13
         Top             =   48
         Width           =   9240
         Begin VB.PictureBox Picture 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   555
            Index           =   0
            Left            =   60
            ScaleHeight     =   525
            ScaleWidth      =   9090
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   60
            Width           =   9120
            Begin VB.Timer Timer1 
               Enabled         =   0   'False
               Interval        =   500
               Left            =   2640
               Top             =   120
            End
            Begin VB.PictureBox Picture 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   1
               Left            =   120
               Picture         =   "WINJOB.frx":030A
               ScaleHeight     =   480
               ScaleWidth      =   480
               TabIndex        =   6
               TabStop         =   0   'False
               Top             =   30
               Width           =   480
            End
            Begin VB.Label lblsuge 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               Caption         =   "Sugerencias"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   720
               TabIndex        =   10
               Top             =   48
               Width           =   8055
               WordWrap        =   -1  'True
            End
         End
      End
      Begin VB.PictureBox Spin1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   9260
         ScaleHeight     =   585
         ScaleWidth      =   225
         TabIndex        =   16
         Top             =   75
         Width           =   255
      End
   End
   Begin VB.PictureBox ButtonBar2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   9570
      TabIndex        =   17
      Top             =   435
      Width           =   9600
      Begin VB.ComboBox txtsearch 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6600
         TabIndex        =   0
         Top             =   90
         Width           =   1935
      End
      Begin VB.ComboBox Usar 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   90
         Width           =   1695
      End
      Begin VB.DriveListBox Drive1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   480
         TabIndex        =   4
         Top             =   90
         Width           =   2175
      End
      Begin VB.Label Label 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Búsqueda rápida"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   5040
         TabIndex        =   1
         Top             =   150
         Width           =   1440
      End
      Begin VB.Label Label 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Usar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   2790
         TabIndex        =   8
         Top             =   150
         Width           =   405
      End
   End
   Begin VB.PictureBox StatusBar1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   9570
      TabIndex        =   18
      Top             =   3675
      Width           =   9600
   End
   Begin VB.PictureBox ButtonBar1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   0
      ScaleHeight     =   405
      ScaleWidth      =   9570
      TabIndex        =   19
      Top             =   0
      Width           =   9600
      Begin VB.PictureBox Restaurar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   7920
         ScaleHeight     =   165
         ScaleWidth      =   1545
         TabIndex        =   15
         Top             =   120
         Visible         =   0   'False
         Width           =   1575
      End
   End
   Begin VB.Menu mArchivo 
      Caption         =   "&Archivo"
      HelpContextID   =   1000
      Tag             =   "Opciones de compresión de archivos"
      Begin VB.Menu mArchivo_opcion 
         Caption         =   "&Comprimir"
         HelpContextID   =   1010
         Index           =   1
         Shortcut        =   ^C
         Tag             =   "Comprime archivos hacia la unidad especificada"
      End
      Begin VB.Menu mArchivo_opcion 
         Caption         =   "&Descomprimir"
         HelpContextID   =   1020
         Index           =   2
         Shortcut        =   ^D
         Tag             =   "Descomprime el archivo seleccionado"
      End
      Begin VB.Menu mArchivo_opcion 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mArchivo_opcion 
         Caption         =   "Conver&tir"
         HelpContextID   =   1030
         Index           =   4
         Shortcut        =   ^T
         Tag             =   "Convierte el archivo seleccionado a formato RAR ó ARJ"
      End
      Begin VB.Menu mArchivo_opcion 
         Caption         =   "Crear e&xe"
         HelpContextID   =   1040
         Index           =   5
         Shortcut        =   ^X
         Tag             =   "Crea un archivo ejecutable a partir de uno comprimido"
      End
      Begin VB.Menu mArchivo_opcion 
         Caption         =   "&Listar"
         HelpContextID   =   1050
         Index           =   6
         Shortcut        =   ^L
         Tag             =   "Muestra los archivos que están en el archivo comprimido"
      End
      Begin VB.Menu mArchivo_opcion 
         Caption         =   "&Mover"
         HelpContextID   =   1060
         Index           =   7
         Shortcut        =   ^M
         Tag             =   "Comprime archivos y luego los borra desde el lugar donde se efectuó la compresión"
      End
      Begin VB.Menu mArchivo_opcion 
         Caption         =   "R&eparar"
         HelpContextID   =   1070
         Index           =   8
         Shortcut        =   ^E
         Tag             =   "Repara archivos que han sido dañados. Sólo ZIP y RAR"
      End
      Begin VB.Menu mArchivo_opcion 
         Caption         =   "&Verificar"
         HelpContextID   =   1080
         Index           =   9
         Shortcut        =   ^V
         Tag             =   "Verifica la integridad del archivo comprimido"
      End
      Begin VB.Menu mArchivo_opcion 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mArchivo_opcion 
         Caption         =   "Bo&rrar"
         HelpContextID   =   1090
         Index           =   11
         Shortcut        =   {DEL}
         Tag             =   "Borra archivos del archivo comprimido"
      End
      Begin VB.Menu mArchivo_opcion 
         Caption         =   "Editar"
         HelpContextID   =   1100
         Index           =   12
         Shortcut        =   {F2}
         Tag             =   "Edita comentarios o otros archivos"
      End
      Begin VB.Menu mArchivo_opcion 
         Caption         =   "Im&primir"
         HelpContextID   =   1110
         Index           =   13
         Shortcut        =   ^P
         Tag             =   "Imprime archivos del archivo comprimido"
      End
      Begin VB.Menu mArchivo_opcion 
         Caption         =   "&Información"
         HelpContextID   =   1120
         Index           =   14
         Shortcut        =   ^I
         Tag             =   "Muestra una lista con los detalles del archivo"
      End
      Begin VB.Menu mArchivo_opcion 
         Caption         =   "E&jecución directa"
         HelpContextID   =   1130
         Index           =   16
         Shortcut        =   ^J
         Tag             =   "Ejecuta el programa seleccionado"
      End
      Begin VB.Menu mArchivo_opcion 
         Caption         =   "-"
         Index           =   18
      End
      Begin VB.Menu mArchivo_opcion 
         Caption         =   "&Salir"
         Index           =   19
         Tag             =   "Sale de WinReady"
      End
   End
   Begin VB.Menu mEdicion 
      Caption         =   "&Edición"
      HelpContextID   =   2000
      Tag             =   "Selección de archivos"
      Begin VB.Menu mEdicion_opcion 
         Caption         =   "&Agregar todos"
         HelpContextID   =   2000
         Index           =   1
         Tag             =   "Agrega todos los archivos de la acción a realizar"
      End
      Begin VB.Menu mEdicion_opcion 
         Caption         =   "&Quitar todos"
         HelpContextID   =   2000
         Index           =   2
         Tag             =   "Quita todos los archivos de la acción a realizar"
      End
      Begin VB.Menu mEdicion_opcion 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mEdicion_opcion 
         Caption         =   "&Selección múltiple"
         HelpContextID   =   2010
         Index           =   4
         Tag             =   "Selecciona archivos de diferentes directorios"
      End
   End
   Begin VB.Menu mOpcion 
      Caption         =   "&Opción"
      HelpContextID   =   3000
      Tag             =   "Opciones de WinReady"
      Begin VB.Menu mOpcion_opcion 
         Caption         =   "&Configuración"
         HelpContextID   =   3010
         Index           =   1
         Tag             =   "Configura los programas que necesita para trabajar"
      End
      Begin VB.Menu mOpcion_opcion 
         Caption         =   "&Lista de directorios"
         HelpContextID   =   3020
         Index           =   2
         Tag             =   "Busca todos los directorios y los graba para luego ser usados en la acción a realizar"
      End
   End
   Begin VB.Menu mUtilitarios 
      Caption         =   "&Utilitarios"
      HelpContextID   =   4000
      Tag             =   "Utilitarios que ofrecen los programas"
      Begin VB.Menu mUtilitarios_opcion 
         Caption         =   "&Arjsort"
         HelpContextID   =   4010
         Index           =   1
         Shortcut        =   ^{F1}
         Tag             =   "Ordena archivos que están en el archivo comprimido. Se requiere ARJ"
      End
      Begin VB.Menu mUtilitarios_opcion 
         Caption         =   "&Rcvt"
         HelpContextID   =   1030
         Index           =   2
         Shortcut        =   ^{F2}
         Tag             =   "Convierte archivos comprimidos a archivos de formato RAR"
      End
      Begin VB.Menu mUtilitarios_opcion 
         Caption         =   "R&earj"
         HelpContextID   =   1030
         Index           =   3
         Shortcut        =   ^{F3}
         Tag             =   "Convierte archivos comprimidos a archivos de formato ARJ. "
      End
   End
   Begin VB.Menu mAyuda 
      Caption         =   "A&yuda"
      Tag             =   "Ayuda de WinReady"
      Begin VB.Menu mAyuda_opcion 
         Caption         =   "&Indice"
         HelpContextID   =   5010
         Index           =   1
         Shortcut        =   {F1}
         Tag             =   "Indice de materias de WinReady"
      End
      Begin VB.Menu mAyuda_opcion 
         Caption         =   "&Buscar ayuda"
         HelpContextID   =   5020
         Index           =   2
         Tag             =   "Busca ayuda por algún determinado tema"
      End
      Begin VB.Menu mAyuda_opcion 
         Caption         =   "&Usar ayuda"
         HelpContextID   =   5030
         Index           =   3
         Tag             =   "Enseña cómo usar la ayuda"
      End
      Begin VB.Menu mAyuda_opcion 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mAyuda_opcion 
         Caption         =   "&Cómo  registrarse"
         HelpContextID   =   5040
         Index           =   5
         Tag             =   "Cómo registrar su copia de WinReady"
      End
      Begin VB.Menu mAyuda_opcion 
         Caption         =   "&Requerimientos"
         HelpContextID   =   5050
         Index           =   6
         Tag             =   "Programas que se necesitan para trabajar con WinReady"
      End
      Begin VB.Menu mAyuda_opcion 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mAyuda_opcion 
         Caption         =   "Acerca &de ...."
         Index           =   8
         Tag             =   "Presenta la información de Copyright y la dirección del software"
      End
   End
End
Attribute VB_Name = "WinJob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Dim nIndex As Integer
Dim IndexHelp As Integer
Dim SW_CHANGE

Private Sub A_Borrar()
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
     StatusMsg "WinReady irá al DOS por un momento. Espere por favor. "
     DoShell PATH_BAT, SW_NORMAL
     ChDir App.Path
     Open PATH_ERROR For Input As #1
    Cerr = LTrim(Input$(LOF(1), 1))
     Close #1
     Kill PATH_ERROR
     COMPRESS_ACTIVE = Usar.List(Usar.ListIndex)
     StatusMsg "Listo"
     Select Case Cerr
    Case "1"
         MsgBox "Error en el archivo " + LCase(Ruta), 48, "WinReady"
    Case Else
         ACTUAL = Ruta
         fborrar.Show 1
     End Select
End Sub

Private Sub A_Comprimir()
     If Usar.Text = "Pkunzip" Then
        Msg = "Error de selección." + CRLF
        Msg = Msg + CRLF
        Msg = Msg + "Pkunzip es un descompresor de archivos" + CRLF
        Msg = Msg + "en formato zip." + CRLF
        Msg = Msg + "Seleccione en la lista ""Usar"" otro " + CRLF
        Msg = Msg + "compresor y vuelva a intentar." + CRLF
        MsgBox Msg, 48, "WinReady"
        Exit Sub
     End If
     If List1.ListCount > 0 Then fcomprimir.Show 1
End Sub

Private Sub A_Convertir()
     If List1.ListCount < 0 Then Exit Sub
     Ruta = Cual()
     Select Case Right(Ruta, 3)
    Case "arj": If Not VerCompresor("Arj") Then RutinaError "Arj", "Convertir": Exit Sub
    Case "lzh": If Not VerCompresor("Lha") Then RutinaError "Lha", "Convertir": Exit Sub
    Case "pak": If Not VerCompresor("Pak") Then RutinaError "Pak", "Convertir": Exit Sub
    Case "rar": If Not VerCompresor("Rar") Then RutinaError "Rar", "Convertir": Exit Sub
    Case "zip":: If Not VerCompresor("Pkunzip") Then RutinaError "Pkunzip", "Convertir": Exit Sub
    Case Else
    ShowErr " convertir "
    Exit Sub
     End Select
     fconvertir.Show 1
End Sub

Private Sub A_CrearExe()
     If List1.ListCount < 0 Then Exit Sub
     Ruta = Cual()
     Select Case Right(Ruta, 3)
    Case "arj": If VerCompresor("Arj") Then Start = DIR_DE_ARJ + " y - je1 " + Ruta Else RutinaError "Arj", "Crear exe": Exit Sub
    Case "lzh": If VerCompresor("Lha") Then Start = DIR_DE_LHA + " s " + Ruta Else RutinaError "Lha", "Crear exe": Exit Sub
    Case "pak": If VerCompresor("Pak") Then Start = DIR_DE_PAK + " c /exe " + Ruta Else RutinaError "Pak", "Crear exe": Exit Sub
    Case "rar": If VerCompresor("Rar") Then Start = DIR_DE_RAR + " s " + Ruta Else RutinaError "Rar", "Crear exe": Exit Sub
    Case "zip":: If DIR_DE_ZIP2EXE <> "" Then Start = DIR_DE_ZIP2EXE + " " + Ruta Else RutinaError "Zip2exe", "Crear exe": Exit Sub
    Case Else
    ShowErr " crear exe "
    Exit Sub
     End Select
     MsgError = "Error al crear archivo EXE > " + PATH_ERROR
     MsgOk = "Exito al crear archivo EXE >" + PATH_ERROR
     StatusMsg "WinReady procederá a crear el archivo ejecutable " & Cual()
     Go Start, MsgError, MsgOk
     StatusMsg "Listo"
     Start = ""
End Sub

Private Sub A_Descomprimir()
     Dim i As Single
     Dim Archivo As String
     If List1.ListCount < 0 Then Exit Sub
     Ruta = Cual()
     Ext = LCase(Right(Ruta, 3))
     timer1.Enabled = 0
     Select Case Ext
    Case "arj": If Not VerCompresor("Arj") Then RutinaError "Arj", "Descomprimir": Exit Sub Else Start = DIR_DE_ARJ + " l " + Ruta + " > " + PATH_LIST: Ext = "Arj"
    Case "lzh": If Not VerCompresor("Lha") Then RutinaError "Lha", "Descomprimir": Exit Sub Else Start = DIR_DE_LHA + " l " + Ruta + " > " + PATH_LIST: Ext = "Lha"
    Case "pak": If Not VerCompresor("Pak") Then RutinaError "Pak", "Descomprimir": Exit Sub Else Start = DIR_DE_PAK + " t " + Ruta + " > " + PATH_LIST: Ext = "Pak"
    Case "zip": If Not VerCompresor("Pkunzip") Then RutinaError "Pkunzip", "Descomprimir": Exit Sub Else Start = DIR_DE_PKUNZIP + " " + Ruta + " -@" + PATH_LIST: Ext = "Pkunzip"
    Case "rar": If Not VerCompresor("Rar") Then RutinaError "Rar", "Descomprimir": Exit Sub Else Start = DIR_DE_RAR + " l " + Ruta + " > " + PATH_LIST: Ext = "Rar"
    Case Else
        timer1.Enabled = -1
        ' NumTareas = GetNumTasks()
        ' Do While GetNumTasks() <> NumTareas
        '    Id = DoEvents()
        'Loop
        If nSELCOMP = -1 Then Start = "": Ext = "": Exit Sub
     End Select
     StatusMsg "WinReady irá al DOS por un momento. Espere por favor. "
     Colocar Ext
     Acoplar Start, " 1 > " + PATH_ERROR, " 0 > " + PATH_ERROR
     DoShell PATH_BAT, SW_NORMAL
     ChDir App.Path
     Open PATH_ERROR For Input As #1
        Cerr = Val(LTrim(Input$(LOF(1), 1)))
     Close #1
     Kill PATH_ERROR
     Kill PATH_BAT
     COMPRESS_ACTIVE = Usar.Text
     StatusMsg "Listo"
     Select Case Cerr
        Case 1
            Msg = "El archivo " + LCase(Ruta) + " NO parece" + CRLF
            Msg = Msg + "ser un archivo en formato comprimido." + CRLF
            Msg = Msg + "o puede estar corrupto." + CRLF
            Msg = Msg + CRLF
            Msg = Msg + "Asegurese que el archivo que intenta" + CRLF
            Msg = Msg + "descomprimir sea un archivo comprimido." + CRLF
            MsgBox Msg, 48, "WinReady"
            Exit Sub
        Case 0
            FileJob = Ruta
            fdescomprimir.Show 1
     End Select
End Sub

Private Sub A_Editar()
     If List1.ListCount > 0 Then PATH_EDITAR = Cual(): feditar.Show 1
End Sub

Private Sub A_EjecucionDirecta()
     fejedir.Show 1
End Sub

Private Sub A_imprimir()
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
     StatusMsg "WinReady irá al DOS por un momento. Espere por favor. "
     DoShell PATH_BAT, SW_NORMAL
     ChDir App.Path
     Open PATH_ERROR For Input As #1
    Cerr = LTrim(Input$(LOF(1), 1))
     Close #1
     Kill PATH_ERROR
     COMPRESS_ACTIVE = Usar.List(Usar.ListIndex)
     StatusMsg "Listo"
     Select Case Cerr
    Case "1"
         MsgBox "Error en el archivo " + LCase(Ruta), 48, "WinReady"
    Case Else
         FileJob = Ruta: fimprimir.Show 1
     End Select
End Sub

Private Sub A_Información()
     If List1.ListCount > 0 Then finfor.Show 1
End Sub

Private Sub A_Listar()
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
     StatusMsg "WinReady irá al DOS por un momento. Espere por favor. "
     DoShell PATH_BAT, SW_NORMAL
     ChDir App.Path
     Open PATH_ERROR For Input As #1
    Cerr = LTrim(Input$(LOF(1), 1))
     Close #1
     Kill PATH_ERROR
     COMPRESS_ACTIVE = Usar.List(Usar.ListIndex)
     StatusMsg "Listo"
     
     Select Case Cerr
    Case "1"
         MsgBox "Error en el archivo " + LCase(Ruta), 48, "WinReady"
    Case Else
         FileJob = Ruta: flistar.Show 1
     End Select
End Sub

Private Sub A_Reparar()
     If List1.ListCount < 0 Then Exit Sub
     Ruta = Cual()
     Select Case Right(Ruta, 3)
    Case "arj", "lzh", "pak"
    Msg = "Extensión no válida." + CRLF
    Msg = Msg + CRLF
    Msg = Msg + "Para reparar archivos la extensión" + CRLF
    Msg = Msg + "debe ser rar o zip." + CRLF
    MsgBox Msg, 48, "WinReady"
    Exit Sub
    Case "zip": If DIR_DE_PKZIPFIX <> "" Then Start = DIR_DE_PKZIPFIX + Ruta Else RutinaError "Pkzipfix", "Reparar": Exit Sub
    Case "rar": If VerCompresor("Rar") Then Start = DIR_DE_RAR + " r " + Ruta Else RutinaError "Rar", "Reparar": Exit Sub
     End Select
     MsgError = "Error al reparar archivo > " + PATH_ERROR
     MsgOk = "Exito al reparar archivo >" + PATH_ERROR
     StatusMsg "WinReady reparara el siguiente archivo : " & Cual() & ". Espere un momento por favor. "
     Go Start, MsgError, MsgOk
     StatusMsg "Listo"
     
     Start = ""
End Sub

Private Sub A_Verificar()
     If List1.ListCount = -1 Then Exit Sub
     Ruta = Cual()
     Select Case Right(Ruta, 3)
    Case "arj": If VerCompresor("Arj") Then Start = DIR_DE_ARJ + " t -jr " + Ruta: Ext = "Arj" Else RutinaError "Arj", "Verificar": Exit Sub
    Case "lzh": If VerCompresor("Lha") Then Start = DIR_DE_LHA + " t " + Ruta: Ext = "Lha" Else RutinaError "Lha", "Verificar": Exit Sub
    Case "pak": If VerCompresor("Pak") Then Start = DIR_DE_PAK + " t " + Ruta: Ext = "Pak" Else RutinaError "Pak", "Verificar": Exit Sub
    Case "rar": If VerCompresor("Rar") Then Start = DIR_DE_RAR + " t " + Ruta: Ext = "Rar" Else RutinaError "Rar", "Verificar": Exit Sub
    Case "zip":: If VerCompresor("Pkunzip") Then Start = DIR_DE_PKUNZIP + " -t " + Ruta: Ext = "Pkunzip" Else RutinaError "Pkunzip", "Verificar": Exit Sub
    Case Else
    ShowErr " verificar "
    Exit Sub
     End Select
     Colocar Ext
     MsgError = "Error al verificar archivo > " + PATH_ERROR
     MsgOk = "Exito al verificar archivo >" + PATH_ERROR
     StatusMsg "WinReady verificará el siguiente archivo : " & Cual() & ". Espere un momento por favor. "
     Go Start, MsgError, MsgOk
     StatusMsg "Listo"
     
     Start = ""
End Sub

Private Sub Apagar()
    If List1.ListCount = 0 Then
        WinJob.Caption = "WinReady -[No hay archivos seleccionados]"
        For i = 0 To 9: ButtonBar1.ButtonEnabled(i) = 0: Next i
        For i = 1 To 2: WinJob.mArchivo_opcion(i).Enabled = 0: Next i
        For i = 4 To 9: WinJob.mArchivo_opcion(i).Enabled = 0: Next i
        For i = 11 To 14: WinJob.mArchivo_opcion(i).Enabled = 0: Next i
        WinJob.mArchivo_opcion(16).Enabled = 0
        mEdicion.Enabled = 0
        mUtilitarios.Enabled = 0
    End If
End Sub

Private Sub ButtonBar1_Click(button As Integer, Group As Integer, State As Integer)
    On Error Resume Next
    Select Case button
        Case 0: F_COMPRIMIR = -1: A_Comprimir
        Case 1: A_Descomprimir
        Case 2: A_CrearExe
        Case 3: A_Verificar
        Case 4: A_Borrar
        Case 5: A_Editar
        Case 6: A_imprimir
        Case 7: E_AgregarTodos
        Case 8: E_QuitarTodos
        Case 9: E_Seleccion
        Case 10: Ret = WinHelp(hWnd, PATH_HELP, HELP_CONTENTS, 1&)
    End Select
End Sub

Private Sub ButtonBar2_Click(button As Integer, Group As Integer, State As Integer)
    Select Case button
        Case 0
        File1.Refresh
        Drive1_Change
        Dir1_Change
        Case 1
        txtsearch.Text = "*.*"
        File1.Pattern = "*.*"
        File1.Refresh
        Drive1_Change
        Dir1_Change
    End Select
End Sub

Private Sub Colocar(Who As String)
     Dim i As Integer
     For i = 0 To Usar.ListCount - 1
    If Usar.List(i) = Who Then Usar.ListIndex = i: Exit For
     Next i
End Sub

Private Function Cual() As String
     For i = 0 To List1.ListCount - 1
    If List1.Selected(i) Then
         Ruta = Path(Dir1) + List1.List(i)
         ACTUAL = Ruta
         Cual = Ruta: Exit Function
    End If
     Next i
End Function

Private Sub Dir1_Change()
     On Error GoTo ManejoErrores
     Dim i As Single
     Dim nTotal As Long
     Dim cPath As String
     File1.Path = Dir1.Path
     cPath = Path(Dir1)
     ChDir Dir1.Path
     List1.Enabled = -1
     List1.Clear
     Screen.MousePointer = 11
     StatusBar1.Message(0) = "Espere un momento .... "
     nTotal = GetCountF(File1) - 1
     For i = 0 To nTotal: List1.AddItem File1.List(i): Next i
     Screen.MousePointer = 0
     StatusBar1.Message(0) = "Listo"
     List1.ListIndex = -1
     If List1.ListCount > 0 Then WinJob.Caption = "WinReady -[" + UCase(Path(Dir1) + File1.Pattern) + "]": List1.Selected(0) = -1: EncenderMenu Else Apagar
     ButtonBar2.ButtonVisible(1) = IIf(File1.Pattern = "*.*", 0, -1)
     ButtonBar2.ButtonEnabled(1) = IIf(File1.Pattern = "*.*", 0, -1)
Exit Sub
ManejoErrores:
    Dir1.Path = App.Path
    Exit Sub
End Sub

Private Sub Dir1_KeyPress(keyascii As Integer)
     If keyascii = 13 Then Dir1.Path = Dir1.List(Dir1.ListIndex): keyascii = 0
End Sub

Private Sub Drive1_Change()
    On Error GoTo ErrorUnidad
    Dir1.Path = drive1.Drive
    Exit Sub
ErrorUnidad:
    MsgBox "La unidad " + drive1.Drive + " no está preparada", 48, "WinReady"
    drive1.Drive = App.Path
    Exit Sub
End Sub

Private Sub E_AgregarTodos()
    Dim ListHwnd As Integer
    Dim Lparam&
    Dim Void     As Long

    ListHwnd = IIf(MULTIPLE = 0, List1.hWnd, List3.hWnd)
    Lparam& = IIf(MULTIPLE = 0, 0 + (GetCount(List1) * (2 ^ 16)), 0 + (GetCount(List3) * (2 ^ 16)))

    Void = SendMessage(ListHwnd, LB_SELITEMRANGE, 1, ByVal Lparam&)
    If MULTIPLE = 0 Then List1.ListIndex = -1 Else List3.ListIndex = -1
End Sub

Private Sub E_QuitarTodos()
     Dim i As Single
     Screen.MousePointer = 11
     If MULTIPLE = 0 Then
    For i = 0 To List1.ListCount - 1: List1.Selected(i) = 0: Next i
    List1.ListIndex = -1
     Else
    For i = 0 To List3.ListCount - 1: List3.Selected(i) = 0: Next i
    List3.ListIndex = -1
     End If
     Screen.MousePointer = 0
End Sub

Private Sub E_Seleccion()
     fselmul.Show 1
End Sub

Private Sub EliminarArchivosDeTrabajo()
     On Error Resume Next
     Kill PATH_EXC
     Kill PATH_ERROR
     Kill PATH_LIST
     Kill PATH_FILES
     Kill PATH_BAT
     Kill PATH_DIRS
End Sub

Private Sub List1_DblClick()
    A_Descomprimir
End Sub

Private Sub List1_MouseDown(button As Integer, Shift As Integer, X As Single, Y As Single)
    List1.MousePointer = 1
End Sub

Private Sub List1_MouseUp(button As Integer, Shift As Integer, X As Single, Y As Single)
    List1.MousePointer = 0
End Sub

Private Sub ListaDirectorios(Path)
    Dim Count, D(), i, DirName
    DirName = Dir(Path, ATTR_DIRECTORY)
    
    Screen.MousePointer = 11
    Do While DirName <> ""
        If DirName <> "." And DirName <> ".." Then
            If GetAttr(Path + DirName) = ATTR_DIRECTORY Then
                If (Count Mod 10) = 0 Then
                    ReDim Preserve D(Count + 10)
                End If
                Count = Count + 1
                D(Count) = DirName
            End If
        End If
        DirName = Dir
    Loop
    For i = 1 To Count
        List1.AddItem LCase(Path & D(i))
        ListaDirectorios Path & D(i) & "\"
    Next i
    Screen.MousePointer = 0
End Sub

Private Sub mArchivo_opcion_Click(Index As Integer)
    Select Case Index
        Case ARCHIVO_COMPRIMIR: F_COMPRIMIR = -1: A_Comprimir
        Case ARCHIVO_DESCOMPRIMIR: A_Descomprimir
        Case ARCHIVO_CONVERTIR: A_Convertir
        Case ARCHIVO_CREAREXE: A_CrearExe
        Case ARCHIVO_LISTAR: A_Listar
        Case ARCHIVO_MOVER: F_COMPRIMIR = 0: A_Comprimir
        Case ARCHIVO_REPARAR: A_Reparar
        Case ARCHIVO_VERIFICAR: A_Verificar
        Case ARCHIVO_BORRAR: A_Borrar
        Case ARCHIVO_EDITAR: A_Editar
        Case ARCHIVO_IMPRIMIR: A_imprimir
        Case ARCHIVO_INFORMACION: A_Información
        Case ARCHIVO_EJECUCION: A_EjecucionDirecta
        'Case ARCHIVO_MODULO: A_Modulo
        Case ARCHIVO_SALIR: Salir
    End Select
End Sub

Private Sub mAyuda_opcion_Click(Index As Integer)
    On Error Resume Next
    Select Case Index
        Case AYUDA_INDICE
            Ret = WinHelp(hWnd, PATH_HELP, HELP_CONTENTS, 1&)
        Case AYUDA_BUSCAR
            Ret = WinHelp(hWnd, PATH_HELP, HELP_PARTIALKEY, "")
        Case AYUDA_USAR
            Ret = WinHelp(hWnd, "", HELP_HELPONHELP, 0&)
        Case AYUDA_COMO
            Ret = WinHelp(hWnd, PATH_HELP, HELP_KEY, "Cómo registrarse")
        Case AYUDA_REQUERIMIENTOS
            Ret = WinHelp(hWnd, PATH_HELP, HELP_KEY, "Requerimientos")
        Case AYUDA_ACERCA
            facerca.Show 1
    End Select
End Sub

Private Sub MDIForm_Activate()
    StatusMsg "Listo"
End Sub

Private Sub MDIForm_Load()
    Turn3DOn
    AppRunning
    On Error Resume Next
    Dim i As Integer
    LLenarPlantilla
    ChDir "\"
    Dir1.Path = CurDir$
    IndexHelp = 0
    SugeHelp cSuge(IndexHelp)
    COMPRESS_ACTIVE = Usar.List(Usar.ListIndex)
    MULTIPLE = 0
    ReDim UNIDADES(0 To drive1.ListCount - 1) As String
    For i = 0 To drive1.ListCount - 1
        UNIDADES(i) = LCase(Left(drive1.List(i), 2))
        TOT_UNIDADES = i
    Next i
    
    Open PATH_HELP For Binary As #1
        mAyuda.Enabled = IIf(LOF(1) > 0, -1, 0)
        ButtonBar1.ButtonEnabled(11) = IIf(LOF(1) > 0, 0, -1)
    Close #1
    ChequearSetup
    If Val(LISTA_DIR) = -1 Then
        fmensaje.Show
        fmensaje.Refresh
        ListaDirectorios LCase(Left(App.Path, 3))
        DoEvents
        Open PATH_DIRS For Output As #1
        For i = 0 To List1.ListCount - 1: Print #1, List1.List(i): Next i
        Close #1
        List1.Clear
        Unload fmensaje
    End If
    ButtonBar2.ControlHwnd = drive1.hWnd
    ButtonBar2.ControlHint = "Seleccione unidad"
    ButtonBar2.ControlMessage = "Selección de la unidad a activar"
    ButtonBar2.ControlHwnd = Usar.hWnd
    ButtonBar2.ControlHint = "Seleccione compresor"
    ButtonBar2.ControlMessage = "Seleccione el compresor a usar en la acción a realizar"
    ButtonBar2.ControlHwnd = txtsearch.hWnd
    ButtonBar2.ControlHint = "Seleccione archivos"
    ButtonBar2.ControlMessage = "Selección rápida de archivos. Use los caracteres especiales como ""*"" y ""?"""
    ButtonBar2.hWndStatusBar = StatusBar1.hWnd
    ButtonBar2.StatusField = 0
    fbienve.Show 1
    Dir1_Change
    SW_CHANGE = 0
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
     GrabarInformacionAIniFile
     EliminarArchivosDeTrabajo
     End
End Sub

Private Sub mEdicion_opcion_Click(Index As Integer)
    Select Case Index
        Case EDICION_AGREGAR_TODOS: E_AgregarTodos
        Case EDICION_QUITAR_TODOS: E_QuitarTodos
        Case EDICION_SELECCION: E_Seleccion
    End Select
End Sub

Private Sub mOpcion_opcion_Click(Index As Integer)
    Select Case Index
        Case OPCIONES_CONFIGURACION: O_Configuracion
        Case OPCIONES_LISTA_DIR: O_ActualizarPosibles
    End Select
End Sub

Private Sub mUtilitarios_opcion_Click(Index As Integer)
    Select Case Index
        Case UTILITARIOS_ARJSORT: U_Arjsort
        Case UTILITARIOS_RCVT: U_Rcvt
        Case UTILITARIOS_REARJ: U_Rearj
    End Select
End Sub

Private Sub O_ActualizarPosibles()
     If mOpcion_opcion(OPCIONES_LISTA_DIR).Checked = 0 Then
    mOpcion_opcion(OPCIONES_LISTA_DIR).Checked = -1
    flistadir.Show 1
     Else
    mOpcion_opcion(OPCIONES_LISTA_DIR).Checked = 0
     End If
     LISTA_DIR = Str(mOpcion_opcion(OPCIONES_LISTA_DIR).Checked)
End Sub

Private Sub O_Configuracion()
     fsetup.Show 1
End Sub

Private Sub Restaurar_Click(Value As Integer)
    Select Case Value
        Case -1
        Dir1.Enabled = -1: Dir1.TabStop = -1: Dir1.TabIndex = List3.TabIndex
        List1.Enabled = -1: List1.TabStop = -1
        List3.Visible = 0: List3.TabStop = 0
        Restaurar.Visible = 0
        Restaurar.Value = 0
        List3.Clear
        List3.ListIndex = -1
        MULTIPLE = 0
        Dir1_Change
        ResetearMenu 0
    End Select
End Sub

Private Sub RutinaError(MsgShow As String, Accion As String)
     Msg = MsgShow + " no esta instalado." + CRLF
     Msg = Msg + CRLF
     Msg = Msg + "En el menú opciones seleccione" + CRLF
     Msg = Msg + """Configuración"",después escriba el path" + CRLF
     Msg = Msg + "en el cuadro de texto " + MsgShow + " ó " + CRLF
     Msg = Msg + "haga click en el botón ""Buscar"" para" + CRLF
     Msg = Msg + "encontrar en su computador el compresor " + MsgShow + "." + CRLF
     Msg = Msg + "Esto es necesario para realizar con" + CRLF
     Msg = Msg + "éxito la acción de " + Accion + "." + CRLF
     MsgBox Msg, 48, "WinReady"
End Sub

Private Sub Salir()
     On Error Resume Next
     GrabarInformacionAIniFile
     Ret = WinHelp(hWnd, PATH_HELP, HELP_QUIT, 0&)
     EliminarArchivosDeTrabajo
     DoEvents
     Turn3DOff
     DoEvents
     End
End Sub

Private Sub ShowErr(MsgShow As String)
     Msg = "Extensión no válida." + CRLF
     Msg = Msg + CRLF
     Msg = Msg + "Para" + MsgShow + "archivos la extensión" + CRLF
     Msg = Msg + "debe ser arj, lzh, pak, rar o zip." + CRLF
     MsgBox Msg, 48, "WinReady"
End Sub

Private Sub Spin1_SpinDown()
    IndexHelp = IndexHelp - 1
    If IndexHelp < 0 Then IndexHelp = 0: Exit Sub
    SugeHelp cSuge(IndexHelp)
End Sub

Private Sub Spin1_SpinUp()
    IndexHelp = IndexHelp + 1
    If IndexHelp > 10 Then IndexHelp = 10: Exit Sub
    SugeHelp cSuge(IndexHelp)
End Sub

Private Sub StatusBar1_SysMenuBrowse(MenuID As Integer, Message As String)
    Select Case MenuID
        Case Is > 0: Message = "Quita, cambia el tamaño o mueve WinReady"
        Case -4096: Message = "Cambia el tamaño de la ventana"
        Case -4080: Message = "Cambia la posición de la ventana"
        Case -4064: Message = "Reduce el tamaño de la ventana a un icono"
        Case -4048: Message = "Cambia el tamaño de la ventana a su maximo tamaño"
        Case -4000: Message = "Sale de WinReady"
        Case -3808: Message = "Cambia el tamaño de la ventana a su tamaño normal"
        Case -3792: Message = "Cambia a otra aplicacion"
    End Select
End Sub

Private Sub SugeHelp(Msg As String)
     lblsuge.Caption = Msg
End Sub

Private Sub Timer1_Timer()
     fselcomp.Show 1
End Sub

Private Sub txtsearch_KeyPress(keyascii As Integer)
    Dim Pos As Integer
    Dim Char, Ingre, Aux As String
    Dim Invalid As String
    Dim TextoAnt As String
    Invalid = ";,:·^<>=/\+[]|@!ª#¬ "
    'Pos = txtsearch.SelStart
    If keyascii = 13 Then
        keyascii = 0
        'If Pos = 0 Or txtsearch.Text = "" Then Exit Sub Else Ingre = Mid(txtsearch.Text, Pos, 1)
        For i = 1 To Len(Invalid)
            Char = Mid(Invalid, i, 1)
            If Char = Ingre Then Exit Sub
            'Aux = Mid(txtsearch.Text, 1, Len(txtsearch.Text) - 1)
            'txtsearch.SelStart = txtsearch.SelStart - 1
            'txtsearch.Text = Aux
            'txtsearch.Refresh
            'Pos = txtsearch.SelStart
        'End If
        Next i
        File1.Pattern = txtsearch.Text: Dir1_Change
        If List1.ListCount = 0 Then
            Apagar
        Else
            EncenderMenu
            txtsearch.AddItem File1.Pattern
        End If
        SW_CHANGE = -1
     End If
End Sub

Private Sub U_Arjsort()
     Dim Copiar, Archivo, Destino As String
     If List1.ListIndex = -1 Then Exit Sub
     Ruta = Cual()
     Select Case Right(Ruta, 3)
    Case "arj": If VerCompresor("Arj") Then farjsort.Show 1 Else RutinaError "Arj", "Ordenar": Exit Sub
    Case Else
         Msg = "Extensión no válida." + CRLF
         Msg = Msg + CRLF
         Msg = Msg + "Para ordenar archivos la extensión" + CRLF
         Msg = Msg + "debe ser arj." + CRLF
         MsgBox Msg, 48, "WinReady"
         Exit Sub
     End Select
End Sub

Private Sub U_Rcvt()
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
         Msg = "Extensión no válida." + CRLF
         Msg = Msg + CRLF
         Msg = Msg + "Para convertir archivos la extensión" + CRLF
         Msg = Msg + "debe ser arj, lzh, rar o zip." + CRLF
         MsgBox Msg, 48, "WinReady"
         Exit Sub
     End Select
     Origen = QuitSlash(DIR_DE_RCVT)
     Archivo = GetFileName(JobPath)
     Destino = Origen + Archivo
     
     Start = DIR_DE_RCVT + " " + Ruta
     StatusMsg "WinReady convertira el siguiente archivo a formato RAR : " & Cual() & ". Espere un momento por favor. "
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
     StatusMsg "Listo"
End Sub

Private Sub U_Rearj()
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
         Msg = "Extensión no válida." + CRLF
         Msg = Msg + CRLF
         Msg = Msg + "Para convertir archivos la extensión" + CRLF
         Msg = Msg + "debe ser arj, lzh, pak, rar o zip." + CRLF
         MsgBox Msg, 48, "WinReady"
         Exit Sub
     End Select
     Origen = QuitSlash(DIR_DE_REARJ)
     Archivo = GetFileName(JobPath)
     Destino = Origen + Archivo
     Start = DIR_DE_REARJ + " " + Ruta
     StatusMsg "WinReady convertira el siguiente archivo a formato ARJ : " & Cual() & ". Espere un momento por favor. "
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
    StatusMsg "Listo"
End Sub

Private Sub Usar_Click()
    COMPRESS_ACTIVE = Usar.List(Usar.ListIndex)
End Sub

Private Sub Usar_Change()
    COMPRESS_ACTIVE = Usar.List(Usar.ListIndex)
End Sub

Private Function VerCompresor(Cual As String)
     Dim i As Integer

     For i = 0 To Usar.ListCount - 1
    If WinJob.Usar.List(i) = Cual Then
         VerCompresor = True
         COMPRESS_ACTIVE = WinJob.Usar.List(Usar.ListIndex)
         Exit Function
    End If
     Next i
     VerCompresor = False
End Function

