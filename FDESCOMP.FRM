VERSION 5.00
Begin VB.Form fdescomprimir 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6090
   ClientLeft      =   1455
   ClientTop       =   780
   ClientWidth     =   6810
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
   HelpContextID   =   1020
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6090
   ScaleWidth      =   6810
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
      Left            =   5280
      ScaleHeight     =   345
      ScaleWidth      =   1425
      TabIndex        =   22
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
      Left            =   5280
      ScaleHeight     =   345
      ScaleWidth      =   1425
      TabIndex        =   21
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
      Left            =   5280
      ScaleHeight     =   345
      ScaleWidth      =   1425
      TabIndex        =   20
      Top             =   480
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   225
      Left            =   5280
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox Frame3d 
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
      ForeColor       =   &H00000000&
      Height          =   1575
      Index           =   1
      Left            =   120
      ScaleHeight     =   1545
      ScaleWidth      =   5025
      TabIndex        =   14
      Top             =   4440
      Width           =   5055
      Begin VB.PictureBox Sobre 
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
         Left            =   2400
         ScaleHeight     =   225
         ScaleWidth      =   2145
         TabIndex        =   19
         Top             =   360
         Width           =   2175
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
         TabIndex        =   18
         Top             =   1120
         Width           =   4815
      End
      Begin VB.PictureBox PosDes 
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
         TabIndex        =   17
         Top             =   840
         Width           =   1575
      End
      Begin VB.PictureBox Mantener 
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
         ScaleWidth      =   1785
         TabIndex        =   16
         Top             =   600
         Width           =   1815
      End
      Begin VB.PictureBox todos 
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
         ScaleWidth      =   1665
         TabIndex        =   15
         Top             =   360
         Width           =   1695
      End
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
      TabIndex        =   13
      Top             =   4030
      Width           =   2295
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
      TabIndex        =   12
      Top             =   2760
      Width           =   2295
   End
   Begin VB.PictureBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
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
      ForeColor       =   &H80000008&
      Height          =   1590
      Left            =   120
      ScaleHeight     =   1590
      ScaleWidth      =   2655
      TabIndex        =   10
      Top             =   2760
      Width           =   2655
   End
   Begin VB.TextBox txtcontra 
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
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   1920
      Width           =   4095
   End
   Begin VB.PictureBox Frame3d 
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
      ForeColor       =   &H00000000&
      Height          =   975
      Index           =   0
      Left            =   120
      ScaleHeight     =   945
      ScaleWidth      =   5025
      TabIndex        =   5
      Top             =   840
      Width           =   5055
      Begin VB.Label Label 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Si el archivo fue comprimido con contraseña digite la contraseña, esto es necesario para que se puedan descomprimir sus archivos."
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
         Height          =   390
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   4650
         WordWrap        =   -1  'True
      End
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
      Left            =   720
      TabIndex        =   3
      Top             =   480
      Width           =   4455
   End
   Begin VB.TextBox txtnombre 
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
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Descomprimir hacia el Directorio/unidad"
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
      Height          =   390
      Index           =   5
      Left            =   2880
      TabIndex        =   11
      Top             =   2280
      Width           =   1815
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Seleccione los archivos a descomprimir"
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
      Height          =   390
      Index           =   4
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   2100
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Co&ntraseña:"
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
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   1940
      Width           =   855
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Hacia"
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
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   420
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Nombre del archivo :"
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
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1470
   End
End
Attribute VB_Name = "fdescomprimir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OverWrite As String
Dim SW

Private Sub Arj()
    On Error Resume Next
    OverWrite = IIf(txtcontra.Text <> "", " -g" + txtcontra.Text, OverWrite)
    Start = IIf(List2.ListIndex = -1, DIR_DE_ARJ + Ext + txtnombre.Text + OverWrite + " " + txthacia.Text, DIR_DE_ARJ + Ext + txtnombre.Text + OverWrite + " -!@ " + txthacia.Text + " @" + PATH_FILES)
    SW = IIf(List2.ListIndex = -1, 0, -1)
    If SW = -1 Then Llenar
    Go Start, MsgError, MsgOk
    If SW Then Kill PATH_FILES
    Start = ""
    Unload fdescomprimir
End Sub

Private Sub btn_Click(Index As Integer)
    Select Case Index
        Case 0: Descompresion
        Case 1: Unload fdescomprimir
        Case 2: Ret = WinHelp(hWnd, PATH_HELP, HELP_CONTENTS, 1020&)
    End Select
End Sub

Private Sub btn_KeyPress(Index As Integer, keyascii As Integer)
    If keyascii = 13 Then
        keyascii = 0
    Select Case Index
        Case 0: Descompresion
        Case 1: Unload fdescomprimir
        Case 2: Ret = WinHelp(hWnd, PATH_HELP, HELP_CONTENTS, 1020&)
    End Select
    End If
End Sub

Private Sub Descompresion()
    Dim Path As String
    Path = txthacia.Text
    If Not IsValidPath(Path, Left(Path, 2)) Then
        GoTo RutinaError
    Else
        If Right(txthacia.Text, 1) <> "\" Then txthacia.Text = txthacia.Text + "\"
    End If
    Select Case COMPRESS_ACTIVE
        Case "Arj": Arj
        Case "Lha": Lha
        Case "Pak": Pak
        Case "Pkunzip": Pkunzip
        Case "Rar": Rar
    End Select
Exit Sub
RutinaError:
    Msg = "Path no encontrado." + CRLF
    Msg = Msg + CRLF
    Msg = Msg + "Digite el path nuevamente, siguiendo las" + CRLF
    Msg = Msg + "normas de DOS. Si esta es una unidad de " + CRLF
    Msg = Msg + "disquete, inserte el disquete en la unidad," + CRLF
    Msg = Msg + "retire el protector si estuviera protegido y " + CRLF
    Msg = Msg + "vuelva a intentar." + CRLF
    MsgBox Msg, 48, "WinReady"
    txthacia.SetFocus
    Exit Sub
End Sub

Private Sub Dir1_Change()
    txthacia.Text = Path(fdescomprimir.Dir1)
End Sub

Private Sub Drive1_Change()
    On Error GoTo ErrorUnidad
    Dir1.Path = drive1.Drive
    Exit Sub
ErrorUnidad:
    Msg = "La unidad " + drive1.Drive + " no está preparada." + CRLF
    Msg = Msg + CRLF
    Msg = Msg + "Si descomprime hacia unidades de disquete" + CRLF
    Msg = Msg + "inserte un disquete en ella, asegurese" + CRLF
    Msg = Msg + "que el disquete este formateado y que no" + CRLF
    Msg = Msg + "este protegido contra escritura." + CRLF
    MsgBox Msg, 48, "WinReady"
    drive1.Drive = App.Path
    Exit Sub
End Sub

Private Sub Form_Load()
    Dim i As Single
    Dim Total As Long
    fdescomprimir.Caption = "Descomprimir -[" + COMPRESS_ACTIVE + "]"
    CentrarForma fdescomprimir
    txtnombre.Text = FileJob
    
    LeerArchivo fdescomprimir.List1
    Mantener.Enabled = IIf(COMPRESS_ACTIVE <> "Lha", -1, 0)
    Select Case COMPRESS_ACTIVE
        Case "Arj": Addlist 8, 3, fdescomprimir.List1, fdescomprimir.List2: Ext = " e -v "
        Case "Lha": Addlist 5, 3, fdescomprimir.List1, fdescomprimir.List2: Ext = " e "
        Case "Pak": Addlist 5, 4, fdescomprimir.List1, fdescomprimir.List2: Ext = " e "
        Case "Rar": Addlist 8, 4, fdescomprimir.List1, fdescomprimir.List2: Ext = " e -v "
        Case Else
        Total = GetCount(List1) - 1
        For i = 0 To Total: List2.AddItem LCase(List1.List(i)): Next i: Ext = " -e "
        List1.Clear
    End Select
    
    txtcontra.Enabled = IIf(COMPRESS_ACTIVE = "Lha", 0, -1)
    PosDes.Enabled = IIf(Val(LISTA_DIR) = -1, -1, 0)
    PosDes.Value = IIf(Val(LISTA_DIR) = -1, -1, 0)
    Posibles.Enabled = IIf(Val(LISTA_DIR) = -1, -1, 0)
    txthacia.Text = Path(fdescomprimir.Dir1) + JobFile
    If Val(LISTA_DIR) = -1 Then LLenarDirs fdescomprimir.Posibles
    Posibles.ListIndex = IIf(Val(LISTA_DIR) = -1, 0, -1)
    
    hInstance = GetInstance(Me)
    ' start the 3D functions
    Start3D
    ' show this form as 3D
    Make3D Me
    
    MsgError = "Error en la descompresión > " + PATH_ERROR
    MsgOk = "Exito en la descompresión >" + PATH_ERROR
    OverWrite = " "
End Sub

Private Sub Lha()
    Start = IIf(List2.ListIndex = -1, DIR_DE_LHA + Ext + txtnombre.Text + OverWrite + txthacia.Text, DIR_DE_LHA + Ext + txtnombre.Text + OverWrite + txthacia.Text + " @" + PATH_FILES)
    If List2.ListIndex > -1 Then Llenar
    Go Start, MsgError, MsgOk
    Start = ""
    Unload fdescomprimir
End Sub

Private Sub List2_Click()
    btn(0).Enabled = IIf(txthacia.Text = "" Or Todos.Value = 0 And List2.SelCount = 0, 0, -1)
End Sub

Private Sub Llenar()
    Dim i As Single
    Open PATH_FILES For Output As #1
        For i = 0 To List2.ListCount - 1
            If List2.Selected(i) Then
                Print #1, List2.List(i)
            End If
        Next i
    Close #1
End Sub

Private Sub Mantener_Click(Value As Integer)
    Select Case COMPRESS_ACTIVE
        Case "Arj": Ext = IIf(Value = -1, " x -v -y ", " e -v ")
        Case "Lha": Ext = IIf(Value = -1, " x ", " e ")
        Case "Pak": Ext = IIf(Value = -1, " e /path ", " e ")
        Case "Pkunzip": Ext = IIf(Value = -1, " -d ", " -e ")
        Case "Rar": Ext = IIf(Value = -1, " x -v -y ", " e -v ")
    End Select
End Sub

Private Sub Pak()
    OverWrite = IIf(txtcontra.Text <> "", OverWrite + " /g=" + txtcontra.Text, OverWrite)
    Start = IIf(List2.ListIndex = -1, DIR_DE_PAK + Ext + txtnombre.Text + OverWrite + " " + txthacia.Text, DIR_DE_PAK + Ext + txtnombre.Text + OverWrite + txthacia.Text + " @" + PATH_FILES)
    SW = IIf(List2.ListIndex = -1, 0, -1)
    If SW = -1 Then Llenar
    Go Start, MsgError, MsgOk
    If SW Then Kill PATH_FILES
    Start = ""
    Unload fdescomprimir
End Sub

Private Sub Pkunzip()
    OverWrite = IIf(txtcontra.Text <> "", OverWrite + " -s" + txtcontra.Text, OverWrite)
    Start = IIf(List2.ListIndex = -1, DIR_DE_PKUNZIP + Ext + txtnombre.Text + OverWrite + " " + txthacia.Text, DIR_DE_PKUNZIP + Ext + txtnombre.Text + OverWrite + txthacia.Text + " @" + PATH_FILES)
    SW = IIf(List2.ListIndex = -1, 0, -1)
    If SW = -1 Then Llenar
    Go Start, MsgError, MsgOk
    If SW Then Kill PATH_FILES
    Start = ""
    Unload fdescomprimir
End Sub

Private Sub Posdes_Click(Value As Integer)
    Select Case Value
        Case -1: Posibles.Enabled = -1: ' Posibles.SetFocus
        Case 0: Posibles.Enabled = 0 ': txthacia.SetFocus
    End Select
End Sub

Private Sub Posibles_Click()
    txthacia.Text = Posibles.List(Posibles.ListIndex)
    Dir1.Path = Posibles.List(Posibles.ListIndex)
End Sub

Private Sub Posibles_Change()
    txthacia.Text = Posibles.List(Posibles.ListIndex)
    Dir1.Path = Posibles.List(Posibles.ListIndex)
End Sub

Private Sub Rar()
    On Error Resume Next
    OverWrite = IIf(txtcontra.Text <> "", " -p" + txtcontra.Text, OverWrite)
    Start = IIf(List2.ListIndex = -1, DIR_DE_RAR + Ext + txtnombre.Text + OverWrite + " " + txthacia.Text, DIR_DE_RAR + Ext + txtnombre.Text + OverWrite + txthacia.Text + " @" + PATH_FILES)
    MsgBox Start
    SW = IIf(List2.ListIndex = -1, 0, -1)
    If SW = -1 Then Llenar
    Go Start, MsgError, MsgOk
    If SW Then Kill PATH_FILES
    Start = ""
    Unload fdescomprimir
End Sub

Private Sub Sobre_Click(Value As Integer)
    Select Case COMPRESS_ACTIVE
        Case "Arj": OverWrite = IIf(Value = -1, " -jo ", OverWrite = " ")
        Case "Lha": OverWrite = IIf(Value = -1, " -c1 ", OverWrite = " ")
        Case "Pak": OverWrite = IIf(Value = -1, " /wa ", OverWrite = " ")
        Case "Pkunzip": OverWrite = IIf(Value = -1, " -o ", " ")
        Case "Rar": OverWrite = IIf(Value = -1, " o+ ", " ")
    End Select
End Sub

Private Sub todos_Click(Value As Integer)
    List2.Enabled = IIf(Value = -1, 0, -1)
    Select Case Value
        Case -1
        List2.ListIndex = -1
        btn(0).Enabled = IIf(txthacia.Text = "", 0, -1)
        Case 0
        btn(0).Enabled = 0
    End Select
End Sub

Private Sub txthacia_Change()
    btn(0).Enabled = IIf(txthacia.Text = "" Or Todos.Value = 0 And List2.SelCount = 0, 0, -1)
End Sub

Private Sub txthacia_KeyPress(keyascii As Integer)
    If keyascii = 13 Then keyascii = 0: If btn(0).Enabled = -1 Then Descompresion
End Sub

