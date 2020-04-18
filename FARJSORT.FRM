VERSION 5.00
Begin VB.Form farjsort 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ArjSort"
   ClientHeight    =   2610
   ClientLeft      =   690
   ClientTop       =   1770
   ClientWidth     =   6240
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2610
   ScaleWidth      =   6240
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
      Left            =   4680
      ScaleHeight     =   345
      ScaleWidth      =   1425
      TabIndex        =   15
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
      Left            =   4680
      ScaleHeight     =   345
      ScaleWidth      =   1425
      TabIndex        =   14
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
      Left            =   4680
      ScaleHeight     =   345
      ScaleWidth      =   1425
      TabIndex        =   13
      Top             =   360
      Width           =   1455
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
      ScaleWidth      =   4425
      TabIndex        =   2
      Top             =   720
      Width           =   4455
      Begin VB.PictureBox Ordenar 
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
         Index           =   9
         Left            =   2160
         ScaleHeight     =   225
         ScaleWidth      =   945
         TabIndex        =   12
         Top             =   1320
         Width           =   975
      End
      Begin VB.PictureBox Ordenar 
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
         Index           =   8
         Left            =   2160
         ScaleHeight     =   225
         ScaleWidth      =   1065
         TabIndex        =   11
         Top             =   1080
         Width           =   1095
      End
      Begin VB.PictureBox Ordenar 
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
         Index           =   7
         Left            =   2160
         ScaleHeight     =   225
         ScaleWidth      =   1305
         TabIndex        =   10
         Top             =   840
         Width           =   1335
      End
      Begin VB.PictureBox Ordenar 
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
         Index           =   6
         Left            =   2160
         ScaleHeight     =   225
         ScaleWidth      =   1785
         TabIndex        =   9
         Top             =   600
         Width           =   1815
      End
      Begin VB.PictureBox Ordenar 
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
         Index           =   5
         Left            =   2160
         ScaleHeight     =   225
         ScaleWidth      =   2025
         TabIndex        =   8
         Top             =   360
         Width           =   2055
      End
      Begin VB.PictureBox Ordenar 
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
         Index           =   4
         Left            =   120
         ScaleHeight     =   225
         ScaleWidth      =   2025
         TabIndex        =   7
         Top             =   1320
         Width           =   2055
      End
      Begin VB.PictureBox Ordenar 
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
         Index           =   3
         Left            =   120
         ScaleHeight     =   225
         ScaleWidth      =   1185
         TabIndex        =   6
         Top             =   1080
         Width           =   1215
      End
      Begin VB.PictureBox Ordenar 
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
         Index           =   2
         Left            =   120
         ScaleHeight     =   225
         ScaleWidth      =   1305
         TabIndex        =   5
         Top             =   840
         Width           =   1335
      End
      Begin VB.PictureBox Ordenar 
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
         Index           =   1
         Left            =   120
         ScaleHeight     =   225
         ScaleWidth      =   945
         TabIndex        =   4
         Top             =   600
         Width           =   975
      End
      Begin VB.PictureBox Ordenar 
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
         Index           =   0
         Left            =   120
         ScaleHeight     =   225
         ScaleWidth      =   1185
         TabIndex        =   3
         Top             =   360
         Width           =   1215
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
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4455
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1380
   End
End
Attribute VB_Name = "farjsort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Orden(0 To 9)
Dim i As Integer

Private Sub btn_Click(Index As Integer)
    Select Case Index
        Case 0: Sort
        Case 1: Unload farjsort
        Case 2: Ret = WinHelp(hWnd, PATH_HELP, HELP_CONTENTS, 4010&)
    End Select
End Sub

Private Sub btn_KeyPress(Index As Integer, keyascii As Integer)
    If keyascii = 13 Then
        keyascii = 0
        Select Case Index
            Case 0: Sort
            Case 1: Unload farjsort
            Case 2: Ret = WinHelp(hWnd, PATH_HELP, HELP_CONTENTS, 4010&)
        End Select
    End If
End Sub

Private Sub Form_Load()
    CentrarForma farjsort

    txtnombre.Text = ACTUAL
    Orden(0) = " /a "
    Orden(1) = " /c "
    Orden(2) = " /d "
    Orden(3) = " /e "
    Orden(4) = " /f "
    Orden(5) = " /o "
    Orden(6) = " /p "
    Orden(7) = " /s "
    Orden(8) = " /t "
    Orden(9) = " /r "

    hInstance = GetInstance(Me)
    ' start the 3D functions
    Start3D
    ' show this form as 3D
    Make3D Me
    
    MsgError = "Error al ordenar archivos > " + PATH_ERROR
    MsgOk = "Exito al ordenar archivos >" + PATH_ERROR
End Sub

Private Sub Ordenar_Click(Index As Integer, Value As Integer)
    i = Index
End Sub

Private Sub Sort()
    On Error Resume Next
    Dim Archivo, Destino As String
    
    Origen = QuitSlash(DIR_DE_ARJSORT)
    Start = DIR_DE_ARJSORT + " " + txtnombre.Text + " " + Orden(i)
    Archivo = GetFileName(DIR_DE_ARJ)
    Destino = Origen + Archivo
    Open Destino For Binary As #2
    If LOF(2) = 0 Then
        FileCopy DIR_DE_ARJ, Destino
        Go Start, MsgError, MsgOk
        Close #2
        Kill Destino
    Else
        Go Start, MsgError, MsgOk
    End If
    Close #2
    Start = ""
    Unload farjsort
End Sub

