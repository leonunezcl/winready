VERSION 5.00
Begin VB.Form fimprimir 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4185
   ClientLeft      =   1575
   ClientTop       =   1440
   ClientWidth     =   6930
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
   HelpContextID   =   1110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4185
   ScaleWidth      =   6930
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
      Left            =   5400
      ScaleHeight     =   345
      ScaleWidth      =   1425
      TabIndex        =   12
      Top             =   1800
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
      Left            =   5400
      ScaleHeight     =   345
      ScaleWidth      =   1425
      TabIndex        =   11
      Top             =   1320
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
      Left            =   5400
      ScaleHeight     =   345
      ScaleWidth      =   1425
      TabIndex        =   10
      Top             =   840
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   225
      Left            =   5400
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2400
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox Frame3D1 
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
      Height          =   1695
      Left            =   120
      ScaleHeight     =   1665
      ScaleWidth      =   5145
      TabIndex        =   4
      Top             =   2400
      Width           =   5175
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
         TabIndex        =   9
         Top             =   1250
         Width           =   4935
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
         ScaleWidth      =   1785
         TabIndex        =   8
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox txtarchivo 
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
         Left            =   120
         TabIndex        =   7
         Top             =   650
         Width           =   4935
      End
      Begin VB.PictureBox Impresora 
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
         Left            =   1200
         ScaleHeight     =   225
         ScaleWidth      =   1065
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
      Begin VB.PictureBox Archivo 
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
         ScaleWidth      =   945
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.PictureBox List2 
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
      Height          =   1395
      Left            =   120
      ScaleHeight     =   1395
      ScaleWidth      =   5175
      TabIndex        =   3
      Top             =   840
      Width           =   5175
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
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Seleccione los archivos a imprimir"
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
      Top             =   600
      Width           =   2370
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Nombre del archivo"
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
      Width           =   1380
   End
End
Attribute VB_Name = "fimprimir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Archivo_Click(Value As Integer)
    btn(0).Enabled = IIf(txtarchivo.Text <> "", -1, 0)
    Select Case Value
        Case -1: txtarchivo.Enabled = -1: txtarchivo.SetFocus
        Case 0: txtarchivo.Enabled = 0
    End Select
End Sub

Private Sub btn_Click(Index As Integer)
    Select Case Index
        Case 0: Imprimir
        Case 1: Unload fimprimir
        Case 2: Ret = WinHelp(hWnd, PATH_HELP, HELP_CONTENTS, 1110&)
    End Select
End Sub

Private Sub btn_KeyPress(Index As Integer, keyascii As Integer)
    If keyascii = 13 Then
        keyascii = 0
    Select Case Index
        Case 0: Imprimir
        Case 1: Unload fimprimir
        Case 2: Ret = WinHelp(hWnd, PATH_HELP, HELP_CONTENTS, 1110&)
    End Select
    End If
End Sub

Private Sub Form_Load()
    CentrarForma fimprimir
    fimprimir.Caption = "Imprimir -[" + COMPRESS_ACTIVE + "]"
    txtnombre.Text = FileJob
    LeerArchivo fimprimir.List1

    Select Case COMPRESS_ACTIVE
        Case "Arj": Addlist 8, 3, fimprimir.List1, fimprimir.List2
        Case "Lha": Addlist 5, 3, fimprimir.List1, fimprimir.List2
        Case "Pak": Addlist 5, 4, fimprimir.List1, fimprimir.List2
        Case "Rar": Addlist 8, 4, fimprimir.List1, fimprimir.List2
        Case Else
        For i = 0 To List1.ListCount - 1: List2.AddItem LCase(List1.List(i)): Next i
        List1.Clear
    End Select
    
    List2.Selected(0) = -1
    PosDes.Enabled = IIf(Val(LISTA_DIR) = -1, -1, 0)
    PosDes.Value = IIf(Val(LISTA_DIR) = -1, -1, 0)
    Posibles.Enabled = IIf(Val(LISTA_DIR) = -1, -1, 0)
    If Val(LISTA_DIR) = -1 Then LLenarDirs fimprimir.Posibles
    If Val(LISTA_DIR) = -1 Then Posibles.ListIndex = 0
    hInstance = GetInstance(Me)
    ' start the 3D functions
    Start3D
    ' show this form as 3D
    Make3D Me
    
    MsgError = "Error en la impresión > " + PATH_ERROR
    MsgOk = "Exito en la impresión > " + PATH_ERROR
End Sub

Private Sub Impresora_Click(Value As Integer)
    btn(0).Enabled = IIf(Value = -1, -1, 0)
    txtarchivo.Enabled = 0
End Sub

Private Sub Imprimir()
    Dim SW, i As Integer
    On Error GoTo RutinaErrores
    SW = IIf(Archivo.Value = -1, -1, 0)
    If Archivo.Value = -1 Then
        Open txtarchivo.Text For Output As #1
        Close #1
        Open PATH_FILES For Output As #1
            For i = 0 To List2.ListCount - 1
                If List2.Selected(i) Then Print #1, List2.List(i)
            Next i
        Close #1
    End If
    Select Case COMPRESS_ACTIVE
        Case "Arj"
            Start = DIR_DE_ARJ + " p " + txtnombre.Text + " -v -!@ @"
            Start = IIf(SW = -1, Start + PATH_FILES + " > " + txtarchivo.Text, Start + PATH_FILES + " > lpt1 ")
        Case "Lha"
            Start = DIR_DE_LHA + " p " + txtnombre.Text + " @"
            Start = IIf(SW = -1, Start + PATH_FILES + " > " + txtarchivo.Text, Start + PATH_FILES + " > lpt1 ")
        Case "Pak"
            Start = DIR_DE_PAK + " p " + txtnombre.Text + " @"
            Start = IIf(SW = -1, Start + PATH_FILES + " > " + txtarchivo.Text, Start + PATH_FILES + " > lpt1 ")
        Case "Pkunzip"
            Start = DIR_DE_PKUNZIP
            Start = IIf(SW = -1, Start + " -c " + txtnombre.Text + " @" + PATH_FILES + " > " + txtarchivo.Text, Start + " -pa " + txtnombre.Text + " @" + PATH_FILES)
        Case "Rar"
            Start = DIR_DE_RAR + " p " + txtnombre.Text + " @"
            Start = IIf(SW = -1, Start + PATH_FILES + " > " + txtarchivo.Text, Start + PATH_FILES + " > lpt1 ")
    End Select
    If Not SW Then MsgBox "Encienda la impresora", 64, "WinReady"
    Go Start, MsgError, MsgOk
    Start = ""
    Kill PATH_FILES
    Unload fimprimir
Exit Sub

RutinaErrores:
    Select Case Err
        Case 61
            Msg = "No hay espacio para imprimir en la " + CRLF
            Msg = Msg + "unidad " + Left(txtarchivo.Text, 2) + CRLF
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
            txtarchivo.SetFocus
    End Select
    Exit Sub
End Sub

Private Sub List2_Click()
    btn(0).Enabled = IIf(List2.SelCount = 0, 0, -1)
End Sub

Private Sub txtarchivo_Change()
    btn(0).Enabled = IIf(txtarchivo.Text <> "", -1, 0)
End Sub

