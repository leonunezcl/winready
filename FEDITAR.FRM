VERSION 5.00
Begin VB.Form feditar 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Editar"
   ClientHeight    =   2850
   ClientLeft      =   225
   ClientTop       =   600
   ClientWidth     =   6855
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
   HelpContextID   =   1100
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2850
   ScaleWidth      =   6855
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
      TabIndex        =   11
      Top             =   1560
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
      TabIndex        =   10
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
      Index           =   0
      Left            =   5280
      ScaleHeight     =   345
      ScaleWidth      =   1425
      TabIndex        =   9
      Top             =   600
      Width           =   1455
   End
   Begin VB.PictureBox Frame3D 
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
      Height          =   1335
      Index           =   1
      Left            =   120
      ScaleHeight     =   1305
      ScaleWidth      =   5025
      TabIndex        =   4
      Top             =   1440
      Width           =   5055
      Begin VB.TextBox txteditor 
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
         TabIndex        =   8
         Top             =   890
         Width           =   4815
      End
      Begin VB.PictureBox Otro 
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
         ScaleWidth      =   1185
         TabIndex        =   7
         Top             =   600
         Width           =   1215
      End
      Begin VB.PictureBox Editor 
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
         Left            =   1920
         ScaleHeight     =   225
         ScaleWidth      =   1425
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
      Begin VB.PictureBox Block 
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
         TabIndex        =   5
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.PictureBox Frame3D 
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
      Height          =   855
      Index           =   0
      Left            =   120
      ScaleHeight     =   825
      ScaleWidth      =   5025
      TabIndex        =   2
      Top             =   480
      Width           =   5055
      Begin VB.Label Label 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Si selecciona otro editor, digite el path donde este se encuentra."
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
         TabIndex        =   3
         Top             =   360
         Width           =   4815
         WordWrap        =   -1  'True
      End
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
      Width           =   3615
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
Attribute VB_Name = "feditar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PathBlock As String
Dim PathEdit As String

Private Sub Block_Click(Value As Integer)
    Select Case Value
        Case -1: btn(0).Enabled = -1
        Case 0: btn(0).Enabled = 0
    End Select
End Sub

Private Sub btn_Click(Index As Integer)
    Select Case Index
        Case 0: Editar
        Case 1: Unload feditar
        Case 2: Ret = WinHelp(hWnd, PATH_HELP, HELP_CONTENTS, 1100&)
    End Select
End Sub

Private Sub btn_KeyPress(Index As Integer, keyascii As Integer)
    If keyascii = 13 Then
        keyascii = 0
    Select Case Index
        Case 0: Editar
        Case 1: Unload feditar
        Case 2: Ret = WinHelp(hWnd, PATH_HELP, HELP_CONTENTS, 1100&)
    End Select
    End If
End Sub

Private Sub Editar()
    Dim Accion As String
    
    On Error GoTo ErrorEditor
    If Block.Value = -1 Then
        Accion = PathBlock + txtnombre.Text
    ElseIf Editor.Value = -1 Then
        Accion = PathEdit + txtnombre.Text
    Else
        Open txteditor.Text For Binary As #1
        Accion = IIf(LOF(1) > 0, txteditor.Text + " " + txtnombre.Text, "")
        Close #1
    End If
    DoShell Accion, SW_NORMAL
    Unload feditar
Exit Sub
ErrorEditor:
    Select Case Err
        Case 64, 68, 70, 71, 75, 76
            Msg = "El nombre del archivo es incorrecto." + CRLF
            Msg = Msg + CRLF
            Msg = Msg + "Digite el nombre del archivo, según la normas" + CRLF
            Msg = Msg + "del DOS." + CRLF
            Msg = Msg + "Si es hacia una unidad de disco flexible" + CRLF
            Msg = Msg + "asegurese que este desprotegida, que exista" + CRLF
            Msg = Msg + "un disquette en ella y que la unidad este disponible."
            MsgBox Msg, 48, "WinReady"
            txteditor.SetFocus
    End Select
    Exit Sub
End Sub

Private Sub Editor_Click(Value As Integer)
    Select Case Value
        Case -1: btn(0).Enabled = -1
        Case 0: btn(0).Enabled = 0
    End Select
End Sub

Private Sub Form_Load()
    CentrarForma feditar
    txtnombre.Text = PATH_EDITAR
    PathBlock = Windir() + "\notepad.exe "
    PathEdit = Left(App.Path, 3) + "dos\edit.com "
    hInstance = GetInstance(Me)
    ' start the 3D functions
    Start3D
    ' show this form as 3D
    Make3D Me
End Sub

Private Sub Otro_Click(Value As Integer)
    Select Case Value
        Case -1
        btn(0).Enabled = -1
        Block.Value = 0: Block.Enabled = 0
        Editor.Value = 0: Editor.Enabled = 0
        txteditor.Enabled = -1: txteditor.SetFocus
        Case 0
        Block.Enabled = -1
        Block.SetFocus
        btn(0).Enabled = 0
        Editor.Enabled = -1
        txteditor.Enabled = 0
    End Select
End Sub

Private Sub txteditor_Change()
    btn(0).Enabled = IIf(txteditor.Text <> "", -1, 0)
End Sub

Private Sub txteditor_KeyPress(keyascii As Integer)
    If keyascii = 13 And btn(0).Enabled = -1 Then keyascii = 0: Editar
End Sub

