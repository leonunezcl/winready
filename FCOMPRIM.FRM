VERSION 5.00
Begin VB.Form fcomprimir 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4320
   ClientLeft      =   1440
   ClientTop       =   1470
   ClientWidth     =   6735
   ClipControls    =   0   'False
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
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4320
   ScaleWidth      =   6735
   Begin VB.PictureBox Excluir 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      ScaleHeight     =   225
      ScaleWidth      =   3105
      TabIndex        =   9
      Top             =   3120
      Width           =   3135
   End
   Begin VB.PictureBox btn 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   5160
      ScaleHeight     =   345
      ScaleWidth      =   1425
      TabIndex        =   16
      Top             =   1920
      Width           =   1455
   End
   Begin VB.PictureBox btn 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   5160
      ScaleHeight     =   345
      ScaleWidth      =   1425
      TabIndex        =   15
      Top             =   1440
      Width           =   1455
   End
   Begin VB.PictureBox btn 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   5160
      ScaleHeight     =   345
      ScaleWidth      =   1425
      TabIndex        =   14
      Top             =   960
      Width           =   1455
   End
   Begin VB.PictureBox btn 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   5160
      ScaleHeight     =   345
      ScaleWidth      =   1425
      TabIndex        =   13
      Top             =   480
      Width           =   1455
   End
   Begin VB.ComboBox Posibles 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   3915
      Width           =   4935
   End
   Begin VB.PictureBox Posdes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      ScaleHeight     =   225
      ScaleWidth      =   1545
      TabIndex        =   11
      Top             =   3600
      Width           =   1575
   End
   Begin VB.PictureBox Usar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      ScaleHeight     =   225
      ScaleWidth      =   2265
      TabIndex        =   10
      Top             =   3360
      Width           =   2295
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2880
      TabIndex        =   8
      Top             =   2720
      Width           =   2175
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   2880
      TabIndex        =   7
      Top             =   1440
      Width           =   2175
   End
   Begin VB.PictureBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1590
      Left            =   120
      ScaleHeight     =   1590
      ScaleWidth      =   2655
      TabIndex        =   5
      Top             =   1440
      Width           =   2655
   End
   Begin VB.TextBox txthacia 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Top             =   480
      Width           =   3495
   End
   Begin VB.TextBox txtdesde 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Comprimir &hacia el Directorio/unidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   960
      Width           =   1335
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "A&rchivos seleccionados a comprimir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Nombre del archivo :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   500
      Width           =   1470
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Desde :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   555
   End
End
Attribute VB_Name = "fcomprimir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim N, SW
Dim C, JobFile, FinalFile As String

Private Sub Arj()
    On Error Resume Next
    Dim i As Integer
    If PASSWORD = -1 Then Mensaje
    If Excluir.Value = -1 Then ExcluirFiles
    PREDEF_ARJ = IIf(Excluir.Value = -1, PREDEF_ARJ + " -x!" + PATH_EXC + " ", PREDEF_ARJ)
    If MULTIPLE = -1 Then
        FillFile fcomprimir.List1
        Start = DIR_DE_ARJ + C + txthacia.Text + PREDEF_ARJ + "!" + PATH_FILES
    Else
        If N Then
            Start = DIR_DE_ARJ + C + txthacia.Text + PREDEF_ARJ + Path(WinJob.Dir1) + WinJob.File1.Pattern
        Else
            FillFile fcomprimir.List1
            Start = DIR_DE_ARJ + C + txthacia.Text + PREDEF_ARJ + "!" + PATH_FILES
        End If
    End If
    MsgBarra
    Go Start, MsgError, MsgOk
    For i = 0 To 7: ARJ_TAB(i) = "": Next i
    If MULTIPLE = -1 Then Kill PATH_FILES
    If Not N Then Kill PATH_FILES
    Kill PATH_EXC
    
    Start = ""
    Unload fcomprimir
End Sub

Private Sub btn_Click(Index As Integer)
    Dim Ret
    Select Case Index
        Case 0: Compresion
        Case 1: Unload fcomprimir
        Case 2: fopciones.Show 1
        Case 3:
        Select Case F_COMPRIMIR
            Case -1: Ret = WinHelp(hWnd, PATH_HELP, HELP_CONTENTS, 1010&)
            Case 0: Ret = WinHelp(hWnd, PATH_HELP, HELP_CONTENTS, 1060&)
        End Select
    End Select
End Sub

Private Sub btn_KeyPress(Index As Integer, keyascii As Integer)
    If keyascii = 13 Then
        keyascii = 0
    Select Case Index
        Case 0: Compresion
        Case 1: Unload fcomprimir
        Case 2: fopciones.Show 1
        Case 3: Ret = IIf(F_COMPRIMIR = -1, WinHelp(hWnd, PATH_HELP, HELP_CONTENTS, 1010&), WinHelp(hWnd, PATH_HELP, HELP_CONTENTS, 1060&))
    End Select
    End If
End Sub

Private Sub Compresion()
    On Error GoTo RutinaError
    Open txthacia.Text For Output As #1
        Print #1, "Leo y Jovi"
    Close #1
    Kill txthacia.Text
    Select Case COMPRESS_ACTIVE
        Case "Arj": Arj
        Case "Lha": Lha
        Case "Pak": Pak
        Case "Pkzip": Pkzip
        Case "Rar": Rar
    End Select
Exit Sub
RutinaError:
    Select Case Err
        Case 61
            Msg = "No hay espacio para comprimir en la " + CRLF
            Msg = Msg + "unidad " + Left(drive1.Drive, 2) + CRLF
            Msg = Msg + CRLF
            Msg = Msg + "Inserte un disco con espacio y vuelva" + CRLF
            Msg = Msg + "a intentar." + CRLF
            MsgBox Msg, 48, "WinReady"
        Case 64, 68, 70, 71, 75, 76
            Msg = "El nombre del archivo es incorrecto." + CRLF
            Msg = Msg + CRLF
            Msg = Msg + "Digite el nombre del archivo, según la normas" + CRLF
            Msg = Msg + "del DOS." + CRLF
            Msg = Msg + "Si es hacia una unidad de disco flexible" + CRLF
            Msg = Msg + "asegurese que este desprotegida, que exista" + CRLF
            Msg = Msg + "un disquette en ella y que la unidad este disponible."
            MsgBox Msg, 48, "WinReady"
            txthacia.SetFocus
    End Select
Exit Sub
End Sub

Private Sub ChangeNameFile()
    txthacia.Text = ""
    txthacia.Text = Posibles.List(Posibles.ListIndex) + JobFile
    If Val(LISTA_DIR) = -1 Then Dir1.Path = Posibles.List(Posibles.ListIndex)
End Sub
                        

Private Sub Dir1_Change()
    txthacia.Text = ""
    If Posibles.Enabled = -1 Then txthacia.Text = Path(fcomprimir.Dir1) + JobFile
End Sub

Private Sub Drive1_Change()
    On Error GoTo ErrorUnidad
    Dir1.Path = drive1.Drive
    Exit Sub
ErrorUnidad:
    Msg = "La unidad " + drive1.Drive + " no está preparada." + CRLF
    Msg = Msg + CRLF
    Msg = Msg + "Si comprime hacia unidades de disquete" + CRLF
    Msg = Msg + "inserte un disquete en ella, asegurese" + CRLF
    Msg = Msg + "que el disquete este formateado y que no" + CRLF
    Msg = Msg + "este protegido contra escritura." + CRLF
    MsgBox Msg, 48, "WinReady"
    drive1.Drive = App.Path
    Exit Sub
End Sub

Private Sub ExcluirFiles()
    On Error Resume Next
    Open PATH_EXC For Output As #1
        Print #1, "*.arj"
        Print #1, "*.lzh"
        Print #1, "*.pak"
        Print #1, "*.zip"
        Print #1, "*.rar"
    Close #1
End Sub

Private Sub Form_Load()
    Dim i As Single
    CentrarForma fcomprimir
    C = IIf(F_COMPRIMIR = -1, " a ", " m ")
    fcomprimir.Caption = IIf(C = " a ", "Comprimir -[" + COMPRESS_ACTIVE + "]", "Mover -[" + COMPRESS_ACTIVE + "]")
    fcomprimir.HelpContextID = IIf(F_COMPRIMIR = -1, 1010, 1060)
    txtdesde.Text = WinJob.Dir1.Path
    PASSWORD = 0
    Select Case COMPRESS_ACTIVE
        Case "Arj": JobFile = QuitSlash(WinJob.Dir1) & ".arj": PREDEF_ARJ = " -jm -js -va ": Paso = PREDEF_ARJ
        Case "Lha": JobFile = QuitSlash(WinJob.Dir1) & ".lzh": PREDEF_LHA = " -z2 ": LHA_TAB(1) = IIf(C = " a ", " a ", " m "): Paso = PREDEF_LHA
        Case "Pak": JobFile = QuitSlash(WinJob.Dir1) & ".pak": PREDEF_PAK = " /sec /p ": Paso = PREDEF_PAK
        Case "Pkzip": JobFile = QuitSlash(WinJob.Dir1) & ".zip": PREDEF_PKZIP = " -ex ": Paso = PREDEF_PKZIP
        Case "Rar": JobFile = QuitSlash(WinJob.Dir1) & ".rar": PREDEF_RAR = " -s -m5 -ehs -va -v1457 ": Paso = PREDEF_RAR
    End Select
    If MULTIPLE = -1 Then
        Select Case WinJob.List3.ListIndex
            Case -1: AddXlist WinJob.List3, fcomprimir.List1, 0
            Case Else: AddXlist WinJob.List3, fcomprimir.List1, 1
        End Select
    Else
        Select Case WinJob.List1.ListIndex
            Case -1: AddXlist WinJob.List1, fcomprimir.List1, 0: N = True
            Case Else: AddXlist WinJob.List1, fcomprimir.List1, 1: N = False
        End Select
    End If
    
    PosDes.Enabled = IIf(Val(LISTA_DIR) = -1, -1, 0)
    PosDes.Value = IIf(Val(LISTA_DIR) = -1, -1, 0)
    Posibles.Enabled = IIf(Val(LISTA_DIR) = -1, -1, 0)
    If Val(LISTA_DIR) = -1 Then ChangeNameFile Else txthacia.Text = Path(fcomprimir.Dir1) + JobFile
    If Val(LISTA_DIR) = -1 Then LLenarDirs fcomprimir.Posibles
    If Val(LISTA_DIR) = -1 Then Posibles.ListIndex = 0
    ' Get instance handle of project
    hInstance = GetInstance(Me)
    ' start the 3D functions
    Start3D
    ' show this form as 3D
    Make3D Me
    MsgError = "Error en la compresión > " + PATH_ERROR
    MsgOk = "Exito en la compresión >" + PATH_ERROR
End Sub

Private Sub Lha()
    On Error Resume Next
    If MULTIPLE = -1 Then
        FillFile fcomprimir.List1
        Start = DIR_DE_LHA + LHA_TAB(1) + txthacia.Text + PREDEF_LHA + "@" + PATH_FILES
    Else
        If N Then
            Start = DIR_DE_LHA + LHA_TAB(1) + txthacia.Text + PREDEF_LHA + Path(WinJob.Dir1) + WinJob.File1.Pattern
        Else
            FillFile fcomprimir.List1
            Start = DIR_DE_LHA + LHA_TAB(1) + txthacia.Text + PREDEF_LHA + "@" + PATH_FILES
        End If
    End If
    MsgBarra
    Go Start, MsgError, MsgOk
    If MULTIPLE = -1 Then Kill PATH_FILES
    If Not N Then Kill PATH_FILES
    Start = ""
    Unload fcomprimir
End Sub

Private Sub Mensaje()
    Msg = ""
    Msg = "Aviso de WinReady." + CRLF
    Msg = Msg + CRLF
    Msg = Msg + "El siguiente archivo esta siendo" + CRLF
    Msg = Msg + "comprimido con la siguiente contraseña" + CRLF
    Msg = Msg + UCase(CONTRASENA) + CRLF
    Msg = Msg + "Anote el nombre del archivo y la " + CRLF
    Msg = Msg + "contraseña asignada a este." + CRLF
    Msg = Msg + "El no recordarla al descomprimir, " + CRLF
    Msg = Msg + "significa no poder recuperar los " + CRLF
    Msg = Msg + "archivos comprimidos." + CRLF
    MsgBox Msg, 16, "WinReady"
End Sub

Private Sub MsgBarra()
    StatusMsg "WinReady irá al DOS para comprimir con el compresor : " & COMPRESS_ACTIVE & " el ó los archivos seleccionados. Espere un momento por favor."
End Sub

Private Sub Pak()
    On Error Resume Next
    If PASSWORD = -1 Then Mensaje
    If MULTIPLE = -1 Then
        FillFile fcomprimir.List1
        Start = DIR_DE_PAK + C + txthacia.Text + PREDEF_PAK + " @" + PATH_FILES
    Else
        If N Then
            Start = DIR_DE_PAK + C + txthacia.Text + PREDEF_PAK + Path(WinJob.Dir1) + WinJob.File1.Pattern
        Else
            FillFile fcomprimir.List1
            Start = DIR_DE_PAK + C + txthacia.Text + PREDEF_PAK + " @" + PATH_FILES
        End If
    End If
    MsgBarra
    Go Start, MsgError, MsgOk
    If MULTIPLE = -1 Then Kill PATH_FILES
    If Not N Then Kill PATH_FILES
    Start = ""
    Unload fcomprimir
End Sub

Private Sub Pkzip()
    On Error Resume Next
    Dim A As String
    If PASSWORD = -1 Then Mensaje
    If Excluir.Value = -1 Then ExcluirFiles
    PREDEF_PKZIP = IIf(Excluir.Value = -1, PREDEF_PKZIP + " -x@" + PATH_EXC + " ", PREDEF_PKZIP)
    A = IIf(C = " a ", " -a ", " -m ")
    If MULTIPLE = -1 Then
        FillFile fcomprimir.List1
        Start = DIR_DE_PKZIP + A + txthacia.Text + PREDEF_PKZIP + "@" + PATH_FILES
    Else
        If N Then
            Start = DIR_DE_PKZIP + A + txthacia.Text + PREDEF_PKZIP + Path(WinJob.Dir1) + WinJob.File1.Pattern
        Else
            FillFile fcomprimir.List1
            Start = DIR_DE_PKZIP + A + txthacia.Text + PREDEF_PKZIP + "@" + PATH_FILES
        End If
    End If
    MsgBarra
    Go Start, MsgError, MsgOk
    Kill PATH_EXC
    If MULTIPLE = -1 Then Kill PATH_FILES
    If Not N Then Kill PATH_FILES
    Start = ""
    Unload fcomprimir
End Sub

Private Sub Posdes_Click(Value As Integer)
    Posibles.Enabled = IIf(Value = -1, -1, 0)
    If Value = 0 Then Posibles.Enabled = 0 ': txthacia.SetFocus
End Sub

Private Sub Posibles_Click()
    ChangeNameFile
End Sub

Private Sub Posibles_Change()
    ChangeNameFile
End Sub

Private Function QuitSlash(Ruta As DirListBox) As String
    Dim j As Integer
    If Mid(Ruta, 2) = ":\" Then QuitSlash = "archivo": Exit Function
    For j = Len(Ruta) To 1 Step -1
        If Mid(Ruta, j, 1) = "\" Then
            QuitSlash = Mid(Ruta, j + 1)
            Exit Function
        End If
    Next j
End Function

Private Sub Rar()
    On Error Resume Next
    Dim i As Integer
    If PASSWORD = -1 Then Mensaje
    If Excluir.Value = -1 Then ExcluirFiles
    PREDEF_RAR = IIf(Excluir.Value = -1, PREDEF_RAR + " -x@" + PATH_EXC + " ", PREDEF_RAR)
    If MULTIPLE = -1 Then
        FillFile fcomprimir.List1
        Start = DIR_DE_RAR + C + txthacia.Text + PREDEF_RAR + "@" + PATH_FILES
    Else
        If N Then
            Start = DIR_DE_RAR + C + txthacia.Text + PREDEF_RAR + Path(WinJob.Dir1) + WinJob.File1.Pattern
        Else
            FillFile fcomprimir.List1
            Start = DIR_DE_RAR + C + txthacia.Text + PREDEF_RAR + "@" + PATH_FILES
        End If
    End If
    MsgBarra
    Go Start, MsgError, MsgOk
    Kill PATH_EXC
    If MULTIPLE = -1 Then Kill PATH_FILES
    If Not N Then Kill PATH_FILES
    Start = ""
    Unload fcomprimir
End Sub

Private Sub txthacia_Click()
    btn(0).Enabled = IIf(txthacia.Text = "", 0, -1)
End Sub

Private Sub txthacia_Change()
    btn(0).Enabled = IIf(txthacia.Text = "", 0, -1)
End Sub

Private Sub txthacia_KeyPress(keyascii As Integer)
    If keyascii = 13 Then keyascii = 0: If btn(0).Enabled = -1 Then Compresion
End Sub

Private Sub Usar_Click(Value As Integer)
    
    Static Aux As String
    Select Case COMPRESS_ACTIVE
        Case "Arj": Aux = PREDEF_ARJ
        Case "Lha": Aux = PREDEF_LHA
        Case "Pak": Aux = PREDEF_PAK
        Case "Pkzip": Aux = PREDEF_PKZIP
        Case "Rar": Aux = PREDEF_RAR
    End Select
    
    Select Case Value
        Case -1:
        Select Case COMPRESS_ACTIVE
            Case "Arj": PREDEF_ARJ = Paso
            Case "Lha": PREDEF_LHA = Paso
            Case "Pak": PREDEF_PAK = Paso
            Case "Pkzip": PREDEF_PKZIP = Paso
            Case "Rar": PREDEF_RAR = Paso
        End Select
        Case 0
        Select Case COMPRESS_ACTIVE
            Case "Arj": PREDEF_ARJ = Aux
            Case "Lha": PREDEF_LHA = Aux
            Case "Pak": PREDEF_PAK = Aux
            Case "Pkzip": PREDEF_PKZIP = Aux
            Case "Rar": PREDEF_RAR = Aux
        End Select
    End Select
End Sub

