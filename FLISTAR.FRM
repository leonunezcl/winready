VERSION 5.00
Begin VB.Form flistar 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3885
   ClientLeft      =   120
   ClientTop       =   585
   ClientWidth     =   7815
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
   HelpContextID   =   1050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3885
   ScaleWidth      =   7815
   Begin VB.PictureBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   2640
      ScaleHeight     =   420
      ScaleWidth      =   1215
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3540
      Left            =   0
      ScaleHeight     =   3540
      ScaleWidth      =   7815
      TabIndex        =   0
      Top             =   360
      Width           =   7815
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Radio"
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
      Left            =   4680
      TabIndex        =   7
      Top             =   120
      Width           =   420
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Nombre del archivo  "
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
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   1470
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fecha/Hora"
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
      Left            =   5400
      TabIndex        =   2
      Top             =   120
      Width           =   870
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Radio"
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
      Left            =   5400
      TabIndex        =   3
      Top             =   600
      Width           =   420
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Comprimido"
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
      Left            =   3720
      TabIndex        =   4
      Top             =   120
      Width           =   810
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Original"
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
      Left            =   2640
      TabIndex        =   5
      Top             =   120
      Width           =   525
   End
End
Attribute VB_Name = "flistar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub AddlistListar(Top As Integer, Down As Integer)
    Dim i As Single
    Dim Linea As String
    Dim Nombre As String
    Dim Original As String
    Dim Comprimido As String
    Dim Radio As String
    Dim Fecha As String
    
    Select Case COMPRESS_ACTIVE
        Case "Arj"
        For i = Top To List2.ListCount - Down
            List1.AddItem LCase(Left(List2.List(i), 58))
        Next i
        Case "Lha"
        For i = Top To List2.ListCount - Down
            List1.AddItem LCase(Mid(List2.List(i), 2, 59))
        Next i
        Case "Pak"
        For i = Top To List2.ListCount - Down
            Linea = LCase(Left(List2.List(i), Len(List2.List(i))))
            Nombre = Left(Linea, 22)
            Fecha = Mid(Linea, 25, 8)
            Radio = Mid(Linea, 53, 4)
            Comprimido = Mid(Linea, 59, 8)
            List1.AddItem Nombre + "    " + Comprimido + " " + Radio + "  " + Fecha
        Next i
        Case "Pkzip"
        For i = Top To List2.ListCount - Down
            Linea = LCase(Left(List2.List(i), Len(List2.List(i))))
            Nombre = Mid(Linea, 48)
            If Len(Nombre) < 12 Then Nombre = Nombre + Space(12 - Len(Nombre))
            Original = Left(Linea, 7)
            Comprimido = Mid(Linea, 17, 7)
            Radio = Mid(Linea, 25, 4)
            Fecha = Mid(Linea, 31, 16)
            List1.AddItem Nombre + "    " + Original + "    " + Comprimido + " " + Radio + "  " + Fecha
            Nombre = ""
        Next i
        Case "Rar"
        For i = Top To List2.ListCount - Down
            Linea = LCase(Left(List2.List(i), Len(List2.List(i))))
            Nombre = Mid(Linea, 2, 12)
            If Len(Nombre) < 12 Then Nombre = Nombre + Space(12 - Len(Nombre))
            Original = Mid(Linea, 16, 7)
            Comprimido = Mid(Linea, 25, 7)
            Radio = Mid(Linea, 34, 4)
            Fecha = Mid(Linea, 40, 14)
            List1.AddItem Nombre + "    " + Original + "    " + Comprimido + "  " + Radio + " " + Fecha
            Nombre = ""
        Next i
    End Select
    List2.Clear
    List1.ListIndex = 0
End Sub

Private Sub Form_Load()
    Dim Linea, Archivo As String
    
    CentrarForma flistar
    flistar.Caption = "Listar -[" + COMPRESS_ACTIVE + "]"
    
    Open PATH_LIST For Input As #1
    Do While Not EOF(1)
        Line Input #1, Linea
        List2.AddItem Left(Linea, Len(Linea))
    Loop
    Close #1
    Kill PATH_LIST
    
    Select Case COMPRESS_ACTIVE
        Case "Arj": AddlistListar 8, 3
        Case "Lha": AddlistListar 5, 3
        Case "Pak": AddlistListar 7, 3
        Case "Pkzip": AddlistListar 14, 3
        Case "Rar": AddlistListar 8, 4
    End Select
    hInstance = GetInstance(Me)
    ' start the 3D functions
    Start3D
    ' show this form as 3D
    Make3D Me

End Sub

