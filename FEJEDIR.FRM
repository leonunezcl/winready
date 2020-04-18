VERSION 5.00
Begin VB.Form fejedir 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ejecución directa"
   ClientHeight    =   1545
   ClientLeft      =   1365
   ClientTop       =   1500
   ClientWidth     =   6720
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
   HelpContextID   =   1130
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1545
   ScaleWidth      =   6720
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
      TabIndex        =   8
      Top             =   1080
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
      TabIndex        =   7
      Top             =   600
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
      TabIndex        =   6
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox txteje 
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
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   4935
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
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   480
      Width           =   3615
   End
   Begin VB.ComboBox Archivos 
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
      Left            =   600
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Ejecutar"
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
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   585
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Posibles destinos"
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
      Top             =   510
      Width           =   1215
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Usar"
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
      Top             =   150
      Width           =   330
   End
End
Attribute VB_Name = "fejedir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Paths(0 To 5) As String

Private Sub archivos_Click()
    txteje.Text = ""
    txteje.Text = Paths(archivos.ListIndex)
End Sub

Private Sub archivos_Change()
    txteje.Text = ""
    txteje.Text = Paths(archivos.ListIndex)
End Sub

Private Sub btn_Click(Index As Integer)
    Select Case Index
        Case 0: Ejecutar
        Case 1: Unload fejedir
        Case 2: Ret = WinHelp(hWnd, PATH_HELP, HELP_CONTENTS, 1130&)
    End Select
End Sub

Private Sub btn_KeyPress(Index As Integer, keyascii As Integer)
    If keyascii = 13 Then
        keyascii = 0
    Select Case Index
        Case 0: Ejecutar
        Case 1: Unload fejedir
        Case 2: Ret = WinHelp(hWnd, PATH_HELP, HELP_CONTENTS, 1130&)
    End Select
    End If
End Sub

Private Sub Ejecutar()
    Start = txteje.Text
    Go Start, MsgError, MsgOk
    Start = ""
    Unload fejedir
End Sub

Private Sub Form_Load()
    Dim i As Integer
    CentrarForma fejedir
    
    If DIR_DE_ARJ <> "" Then archivos.AddItem "arj.exe"
    If DIR_DE_LHA <> "" Then archivos.AddItem "lha.exe"
    If DIR_DE_PAK <> "" Then archivos.AddItem "pak.exe"
    If DIR_DE_PKZIP <> "" Then archivos.AddItem "pkzip.exe"
    If DIR_DE_PKUNZIP <> "" Then archivos.AddItem "pkunzip.exe"
    If DIR_DE_RAR <> "" Then archivos.AddItem "rar.exe"
    
    Paths(0) = DIR_DE_ARJ
    Paths(1) = DIR_DE_LHA
    Paths(2) = DIR_DE_PAK
    Paths(3) = DIR_DE_PKZIP
    Paths(4) = DIR_DE_PKUNZIP
    Paths(5) = DIR_DE_RAR

    For i = 0 To archivos.ListCount - 1
        If archivos.List(i) <> "" Then
            archivos.ListIndex = i
            txteje.Text = Paths(i)
            Exit For
        End If
    Next i
    Posibles.Enabled = IIf(Val(LISTA_DIR) = -1, -1, 0)
    If Val(LISTA_DIR) = -1 Then LLenarDirs fejedir.Posibles
    If Val(LISTA_DIR) = -1 Then Posibles.ListIndex = 0
    hInstance = GetInstance(Me)
    ' start the 3D functions
    Start3D
    ' show this form as 3D
    Make3D Me
    
    MsgError = "Error en la ejecución > " + PATH_ERROR
    MsgOk = "Exito en la  ejecución >" + PATH_ERROR
End Sub

Private Sub txteje_Change()
    btn(0).Enabled = IIf(txteje.Text <> "", -1, 0)
End Sub

Private Sub txteje_KeyPress(keyascii As Integer)
    If keyascii = 13 And btn(0).Enabled = -1 Then keyascii = 0: Ejecutar
End Sub

