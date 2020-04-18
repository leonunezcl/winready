VERSION 5.00
Begin VB.Form fconvertir 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Convertir"
   ClientHeight    =   2775
   ClientLeft      =   105
   ClientTop       =   1425
   ClientWidth     =   6600
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
   HelpContextID   =   1030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2775
   ScaleWidth      =   6600
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
      Left            =   5040
      ScaleHeight     =   345
      ScaleWidth      =   1425
      TabIndex        =   10
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
      Left            =   5040
      ScaleHeight     =   345
      ScaleWidth      =   1425
      TabIndex        =   9
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
      Left            =   5040
      ScaleHeight     =   345
      ScaleWidth      =   1425
      TabIndex        =   8
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
      Height          =   735
      Index           =   1
      Left            =   120
      ScaleHeight     =   705
      ScaleWidth      =   4785
      TabIndex        =   5
      Top             =   1920
      Width           =   4815
      Begin VB.PictureBox Final 
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
         Left            =   1320
         ScaleHeight     =   225
         ScaleWidth      =   825
         TabIndex        =   7
         Top             =   360
         Width           =   855
      End
      Begin VB.PictureBox Final 
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
         ScaleWidth      =   945
         TabIndex        =   6
         Top             =   360
         Width           =   975
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
      Height          =   1335
      Index           =   0
      Left            =   120
      ScaleHeight     =   1305
      ScaleWidth      =   4785
      TabIndex        =   2
      Top             =   480
      Width           =   4815
      Begin VB.Label Label 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Rearj : convierte a archivo ""ARJ"" los siguientes formatos : Arj, Lzh, Pak, Rar y Zip."
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
         TabIndex        =   4
         Top             =   840
         Width           =   4530
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Rcvt : convierte a archivo ""RAR"" los siguientes formatos : Arj, Lzh, Rar y Zip."
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
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   4545
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
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   3255
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
Attribute VB_Name = "fconvertir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btn_Click(Index As Integer)
    If Final(0).Value = 0 And Final(1).Value = 0 Then Exit Sub
    Select Case Index
        Case 0: Convertir
        Case 1: Unload fconvertir
        Case 2: Ret = WinHelp(hWnd, PATH_HELP, HELP_CONTENTS, 1030&)
    End Select
End Sub

Private Sub btn_KeyPress(Index As Integer, keyascii As Integer)
    If Final(0).Value = 0 And Final(1).Value = 0 Then Exit Sub
    If keyascii = 13 Then
        keyascii = 0
    Select Case Index
        Case 0: Convertir
        Case 1: Unload fconvertir
        Case 2: Ret = WinHelp(hWnd, PATH_HELP, HELP_CONTENTS, 1030&)
    End Select
    End If
End Sub

Private Sub Convertir()
    Dim i As Integer
    Dim Copiar, Archivo, Ext, Destino As String
     
    Origen = IIf(Final(0).Value = -1, QuitSlash(DIR_DE_RCVT), QuitSlash(DIR_DE_REARJ))
    Conversor = IIf(Final(0).Value = -1, DIR_DE_RCVT + " " + txtnombre.Text, DIR_DE_REARJ + " " + txtnombre.Text)
    Ext = Right(txtnombre.Text, 3)
    Select Case Right(txtnombre.Text, 3)
        Case "arj": Copiar = DIR_DE_ARJ
        Case "lzh": Copiar = DIR_DE_LHA
        Case "pak": Copiar = DIR_DE_PAK
        Case "rar": Copiar = DIR_DE_RAR
        Case "zip": Copiar = DIR_DE_PKUNZIP
    End Select
    If Final(0).Value = -1 And Ext = "pak" Then
        Msg = "Rcvt mensaje." + CRLF
        Msg = Msg + CRLF
        Msg = Msg + "Rcvt no puede convertir archivos" + CRLF
        Msg = Msg + "con extensión .pak a .rar" + CRLF
        Msg = Msg + "Descomprimalos y luego vuelva a" + CRLF
        Msg = Msg + "comprimirlos usando Rar." + CRLF
        MsgBox Msg, 48, "WinReady"
        Unload fconvertir
    End If
    Archivo = GetFileName(Copiar)
    Destino = Origen + Archivo
    FileCopy Copiar, Destino
    Go Conversor, MsgError, MsgOk
    Unload fconvertir
End Sub

Private Sub final_Click(Index As Integer, Value As Integer)
    btn(0).Enabled = -1
End Sub

Private Sub Form_Load()
    Dim i As Integer
    hInstance = GetInstance(Me)
    ' start the 3D functions
    Start3D
    ' show this form as 3D
    Make3D Me
    
    CentrarForma fconvertir
    txtnombre.Text = ACTUAL
    
    Final(0).Enabled = IIf(DIR_DE_RCVT <> "" And DIR_DE_RCVT_CFG <> "", -1, 0)
    Final(0).Value = IIf(Final(0).Enabled = -1, -1, 0)
    Final(1).Enabled = IIf(DIR_DE_REARJ <> "" And DIR_DE_REARJ_CFG <> "", -1, 0)
    Final(1).Value = IIf(Final(1).Enabled = -1, -1, 0)
    
    MsgError = "Error en la conversión > " + PATH_ERROR
    MsgOk = "Exito en la conversión >" + PATH_ERROR
End Sub

