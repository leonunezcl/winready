VERSION 5.00
Begin VB.Form fborrar 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Borrar"
   ClientHeight    =   2940
   ClientLeft      =   180
   ClientTop       =   750
   ClientWidth     =   7095
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
   HelpContextID   =   1070
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2940
   ScaleWidth      =   7095
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
      Left            =   5520
      ScaleHeight     =   345
      ScaleWidth      =   1425
      TabIndex        =   7
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
      Left            =   5520
      ScaleHeight     =   345
      ScaleWidth      =   1425
      TabIndex        =   6
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
      Left            =   5520
      ScaleHeight     =   345
      ScaleWidth      =   1425
      TabIndex        =   5
      Top             =   120
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   225
      Left            =   5640
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1680
      Visible         =   0   'False
      Width           =   1455
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
      ScaleWidth      =   5295
      TabIndex        =   4
      Top             =   1440
      Width           =   5295
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
      Width           =   3735
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Seleccione los archivos "
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
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Nota : No se pueden borrar archivos, si estos están comprimidos en múltiples vólumenes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   390
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   5295
      WordWrap        =   -1  'True
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
      Top             =   120
      Width           =   1380
   End
End
Attribute VB_Name = "fborrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Arj()
    LlenarBorrar
    Start = DIR_DE_ARJ + " d " + txtnombre.Text + " !" + PATH_FILES
    Go Start, MsgError, MsgOk
    Kill PATH_FILES
    Start = ""
    Unload fborrar
End Sub

Private Sub Borrar()
    Select Case COMPRESS_ACTIVE
        Case "Arj": Arj
        Case "Lha": Lha
        Case "Pak": Pak
        Case "Pkzip": Pkzip
        Case "Rar": Rar
    End Select
End Sub

Private Sub btn_Click(Index As Integer)
    Select Case Index
        Case 0: Borrar
        Case 1: Unload fborrar
        Case 2: Ret = WinHelp(hWnd, PATH_HELP, HELP_CONTENTS, 1090&)
    End Select
End Sub

Private Sub btn_KeyPress(Index As Integer, keyascii As Integer)
    If keyascii = 13 Then
    keyascii = 0
    Select Case Index
        Case 0: Borrar
        Case 1: Unload fborrar
        Case 2: Ret = WinHelp(hWnd, PATH_HELP, HELP_CONTENTS, 1090&)
    End Select
    End If
End Sub

Private Sub Form_Load()
    CentrarForma fborrar
    fborrar.Caption = "Borrar -[" + COMPRESS_ACTIVE + "]"
    txtnombre.Text = ACTUAL
    
    LeerArchivo fborrar.List1

    Select Case COMPRESS_ACTIVE
        Case "Arj": Addlist 8, 3, fborrar.List1, fborrar.List2
        Case "Lha": Addlist 5, 3, fborrar.List1, fborrar.List2
        Case "Pak": Addlist 6, 4, fborrar.List1, fborrar.List2
        Case "Pkzip": Addlist 14, 3, fborrar.List1, fborrar.List2
        Case "Rar": Addlist 8, 4, fborrar.List1, fborrar.List2
    End Select
    List2.Selected(0) = -1
    hInstance = GetInstance(Me)
    ' start the 3D functions
    Start3D
    ' show this form as 3D
    Make3D Me
    
    MsgError = "Error al borrar archivos > " + PATH_ERROR
    MsgOk = "Exito al borrar archivos >" + PATH_ERROR

End Sub

Private Sub Lha()
    LlenarBorrar
    Start = DIR_DE_LHA + " d " + txtnombre.Text + " @" + PATH_FILES
    Go Start, MsgError, MsgOk
    Kill PATH_FILES
    Start = ""
    Unload fborrar
End Sub

Private Sub List2_Click()
    btn(0).Enabled = IIf(List2.SelCount > 0, -1, 0)
End Sub

Private Sub LlenarBorrar()
    Dim i As Single
    On Error Resume Next
    Open PATH_FILES For Output As #1
        For i = 0 To List2.ListCount - 1
            If List2.Selected(i) Then Print #1, List2.List(i)
        Next i
    Close #1
End Sub

Private Sub Pak()
    LlenarBorrar
    Start = DIR_DE_PAK + " d " + txtnombre.Text + " @" + PATH_FILES
    Go Start, MsgError, MsgOk
    Kill PATH_FILES
    Start = ""
    Unload fborrar
End Sub

Private Sub Pkzip()
    LlenarBorrar
    Start = DIR_DE_PKZIP + " -d " + txtnombre.Text + " @" + PATH_FILES
    Go Start, MsgError, MsgOk
    Kill PATH_FILES
    Start = ""
    Unload fborrar
End Sub

Private Sub Rar()
    LlenarBorrar
    Start = DIR_DE_RAR + " d " + txtnombre.Text + " @" + PATH_FILES
    Go Start, MsgError, MsgOk
    Kill PATH_FILES
    Start = ""
    Unload fborrar
End Sub

