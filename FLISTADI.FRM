VERSION 5.00
Begin VB.Form flistadir 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lista de directorios"
   ClientHeight    =   2835
   ClientLeft      =   960
   ClientTop       =   2130
   ClientWidth     =   5985
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
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2835
   ScaleWidth      =   5985
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
      Left            =   4440
      ScaleHeight     =   345
      ScaleWidth      =   1425
      TabIndex        =   4
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
      Index           =   1
      Left            =   4440
      ScaleHeight     =   345
      ScaleWidth      =   1425
      TabIndex        =   3
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
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   4440
      ScaleHeight     =   345
      ScaleWidth      =   1425
      TabIndex        =   2
      Top             =   360
      Width           =   1455
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
      Height          =   2370
      Left            =   120
      ScaleHeight     =   2370
      ScaleWidth      =   4215
      TabIndex        =   0
      Top             =   360
      Width           =   4215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Directorios encontrados :"
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
      TabIndex        =   1
      Top             =   120
      Width           =   1770
   End
End
Attribute VB_Name = "flistadir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Aceptar()
    Dim i As Single
    WinJob.mOpcion_opcion(OPCIONES_LISTA_DIR).Checked = -1
    Open PATH_DIRS For Output As #1
        For i = 0 To List1.ListCount - 1
            Print #1, List1.List(i)
        Next i
    Close #1
    List1.Clear
    LISTA_DIR = Str(WinJob.mOpcion_opcion(OPCIONES_LISTA_DIR).Checked)
    Unload flistadir
End Sub

Private Sub btn_Click(Index As Integer)
    Select Case Index
        Case 0: Aceptar
        Case 1: Cancelar
        Case 2: Ret = WinHelp(hWnd, PATH_HELP, HELP_CONTENTS, 6&)
    End Select
End Sub

Private Sub btn_KeyPress(Index As Integer, keyascii As Integer)
    If keyascii = 13 Then
        keyascii = 0
    Select Case Index
        Case 0: Aceptar
        Case 1: Cancelar
        Case 2: Ret = WinHelp(hWnd, PATH_HELP, HELP_CONTENTS, 6&)
    End Select
    End If
End Sub

Private Sub Cancelar()
    WinJob.mOpcion_opcion(OPCIONES_LISTA_DIR).Checked = 0
    LISTA_DIR = Str(WinJob.mOpcion_opcion(OPCIONES_LISTA_DIR).Checked)
    Unload flistadir
End Sub

Private Sub Form_Load()
    Load fmensaje
    fmensaje.Show
    fmensaje.Refresh
    CentrarForma flistadir
    Actualizar LCase(Left(App.Path, 3))
    Unload fmensaje
    hInstance = GetInstance(Me)
    ' start the 3D functions
    Start3D
    ' show this form as 3D
    Make3D Me

End Sub

