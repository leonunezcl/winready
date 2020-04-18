VERSION 5.00
Begin VB.Form fselmul 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selección múltiple"
   ClientHeight    =   4260
   ClientLeft      =   960
   ClientTop       =   1560
   ClientWidth     =   6120
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
   ScaleHeight     =   4260
   ScaleWidth      =   6120
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
      Height          =   1590
      Left            =   120
      ScaleHeight     =   1590
      ScaleWidth      =   1815
      TabIndex        =   1
      Top             =   360
      Width           =   1815
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
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   3
      Left            =   120
      ScaleHeight     =   345
      ScaleWidth      =   1065
      TabIndex        =   6
      Top             =   2040
      Width           =   1095
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
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   2
      Left            =   4560
      ScaleHeight     =   345
      ScaleWidth      =   1425
      TabIndex        =   14
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
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   1
      Left            =   4560
      ScaleHeight     =   345
      ScaleWidth      =   1425
      TabIndex        =   13
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
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   0
      Left            =   4560
      ScaleHeight     =   345
      ScaleWidth      =   1425
      TabIndex        =   12
      Top             =   840
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
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   5
      Left            =   2280
      ScaleHeight     =   345
      ScaleWidth      =   1065
      TabIndex        =   8
      Top             =   2040
      Width           =   1095
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
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   6
      Left            =   3360
      ScaleHeight     =   345
      ScaleWidth      =   1065
      TabIndex        =   9
      Top             =   2040
      Width           =   1095
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
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   4
      Left            =   1200
      ScaleHeight     =   345
      ScaleWidth      =   1065
      TabIndex        =   7
      Top             =   2040
      Width           =   1095
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
      Height          =   1395
      Left            =   120
      ScaleHeight     =   1395
      ScaleWidth      =   4335
      TabIndex        =   11
      Top             =   2760
      Width           =   4335
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
      Left            =   2040
      TabIndex        =   5
      Top             =   780
      Width           =   2415
   End
   Begin VB.FileListBox File1 
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
      Height          =   420
      Left            =   4560
      MultiSelect     =   2  'Extended
      TabIndex        =   2
      Top             =   2280
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "A&rchivos seleccionados"
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
      TabIndex        =   10
      Top             =   2520
      Width           =   1680
   End
   Begin VB.Label DirActivo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "DirActivo"
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
      Left            =   2040
      TabIndex        =   4
      Top             =   360
      Width           =   645
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Directorios"
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
      Left            =   2040
      TabIndex        =   3
      Top             =   120
      Width           =   750
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Arc&hivos"
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
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "fselmul"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Ruta As String
Dim NewPattern, OldPattern As String

Private Sub Agregar()
    Dim i As Single
    Ruta = Path(Dir1)
    If File1.ListIndex = 0 Then
        List1.AddItem Ruta + File1.List(i)
        btn(4).Enabled = -1: btn(6).Enabled = -1
        Exit Sub
    End If
    Screen.MousePointer = 11
    For i = 0 To File1.ListCount - 1
        If File1.Selected(i) Then List1.AddItem Ruta + File1.List(i)
    Next i
    Screen.MousePointer = 0
    btn(0).Enabled = -1: btn(4).Enabled = -1: btn(6).Enabled = -1
End Sub

Private Sub AgregarTodos()
    Dim i As Single
    Screen.MousePointer = 11
    Ruta = Path(Dir1)
    For i = 0 To File1.ListCount - 1
        List1.AddItem Ruta + File1.List(i)
    Next i
    btn(0).Enabled = -1: btn(4).Enabled = -1: btn(6).Enabled = -1
    Screen.MousePointer = 0
End Sub

Private Sub btn_Click(Index As Integer)
    Select Case Index
        Case 0: Seleccionar
        Case 1: Unload fselmul
        Case 2: Ret = WinHelp(hWnd, PATH_HELP, HELP_CONTENTS, 2010&)
        Case 3: Agregar
        Case 4: Quitar
        Case 5: AgregarTodos
        Case 6: QuitarTodos
    End Select
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
    DirActivo.Caption = Dir1.Path
    VerBotones
End Sub

Private Sub Dir1_KeyPress(keyascii As Integer)
    If keyascii = 13 Then
        keyascii = 0
        File1.Path = Dir1.List(Dir1.ListIndex)
        VerBotones
    End If
End Sub

Private Sub Form_Load()
    CentrarForma fselmul
    List1.Clear
    List2.Clear
    OldPattern = WinJob.File1.Pattern
    NewPattern = "*.*"
    Dir1_Change
    hInstance = GetInstance(Me)
    ' start the 3D functions
    Start3D
    ' show this form as 3D
    Make3D Me

End Sub

Private Sub List1_Click()
    btn(4).Enabled = IIf(List1.SelCount = 0, 0, -1)
    btn(6).Enabled = IIf(List1.SelCount = 0, 0, -1)
End Sub

Private Sub Quitar()
    Dim i As Single
    Ruta = Path(Dir1)
    If List1.ListIndex = -1 Then Exit Sub
    Screen.MousePointer = 11
Remover:
    For i = 0 To List1.ListCount - 1
        If List1.Selected(i) Then
            List1.ListIndex = i
            List1.RemoveItem i
            GoTo Remover
        End If
    Next i
    Screen.MousePointer = 0
    If List1.ListCount = 0 Then btn(0).Enabled = 0: btn(4).Enabled = 0: btn(6).Enabled = 0
    If File1.ListCount = 0 Then
        For i = 3 To 6: btn(i).Enabled = 0: Next i
    Else
        btn(3).Enabled = -1: btn(5).Enabled = -1
    End If
    List1.ListIndex = -1
End Sub

Private Sub QuitarTodos()
    Screen.MousePointer = 11
    List1.Clear
    Screen.MousePointer = 0
    btn(0).Enabled = 0: btn(4).Enabled = 0: btn(6).Enabled = 0
    If File1.ListCount = 0 Then
        For i = 3 To 6: btn(i).Enabled = 0: Next i
    Else
        btn(3).Enabled = -1: btn(5).Enabled = -1
    End If
    List1.ListIndex = -1
End Sub

Private Sub Seleccionar()
    Dim i As Single
    
    WinJob.Dir1.Enabled = IIf(List1.ListCount = 0, -1, 0)
    WinJob.Dir1.TabStop = IIf(List1.ListCount = 0, -1, 0)
    WinJob.Dir1.TabIndex = IIf(List1.ListCount = 0, WinJob.List3.TabIndex, WinJob.Dir1.TabIndex)
    MULTIPLE = IIf(List1.ListCount = 0, 0, -1)
    WinJob.Restaurar.Visible = IIf(List1.ListCount = 0, 0, -1)
    WinJob.List1.Enabled = IIf(List1.ListCount = 0, -1, 0)
    WinJob.List1.TabStop = IIf(List1.ListCount = 0, -1, 0)
    WinJob.List3.Visible = IIf(List1.ListCount = 0, 0, -1)
    WinJob.List3.TabStop = IIf(List1.ListCount = 0, 0, -1)
    WinJob.List3.Clear
    
    If List1.ListCount = 0 Then
        WinJob.Restaurar.Value = 0
        ResetearMenu 0
        WinJob.Caption = "WinReady-[" + Path(WinJob.Dir1) + File1.Pattern + "]"
        WinJob.ButtonBar2.ButtonEnabled(1) = IIf(WinJob.File1.Pattern = "*.*", 0, -1)
        WinJob.ButtonBar2.ButtonVisible(1) = IIf(WinJob.File1.Pattern = "*.*", 0, -1)
        Else
        For i = 0 To List1.ListCount - 1: WinJob.List3.AddItem List1.List(i): Next i
        WinJob.File1.Pattern = OldPattern
        WinJob.List3.ListIndex = 0
        WinJob.List3.Selected(WinJob.List3.ListIndex) = -1
        WinJob.Caption = "WinReady-[Selección múltiple de archivos]"
        WinJob.ButtonBar2.ButtonEnabled(1) = 0
        WinJob.ButtonBar2.ButtonVisible(1) = 0
        ResetearMenu 1
    End If
    Unload fselmul
End Sub

Private Sub VerBotones()
    Dim i As Single
    List2.Clear
    For i = 0 To File1.ListCount - 1
        List2.AddItem File1.List(i)
    Next i
    'Vacios archivos y lista
    btn(0).Enabled = IIf(File1.ListCount = 0 And List1.ListCount = 0, 0, -1)
    If File1.ListCount = 0 And List1.ListCount = 0 Then
        For i = 3 To 6: btn(i).Enabled = 0: Next i
    'Vacio archivo y lista con algo
    ElseIf File1.ListCount = 0 And List1.ListCount > 0 Then
        btn(3).Enabled = 0: btn(5).Enabled = 0
    'Archivo con algo y lista vacia
    ElseIf File1.ListCount > 0 And List1.ListCount = 0 Then
        btn(3).Enabled = -1: btn(5).Enabled = -1
        btn(4).Enabled = 0: btn(6).Enabled = 0
    'Archivo con algo
    ElseIf File1.ListCount > 0 Then
        btn(3).Enabled = -1: btn(5).Enabled = -1
    End If
End Sub

