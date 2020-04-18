VERSION 5.00
Begin VB.Form fsetup 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración"
   ClientHeight    =   5850
   ClientLeft      =   1680
   ClientTop       =   855
   ClientWidth     =   5520
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
   ScaleHeight     =   5850
   ScaleWidth      =   5520
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      Height          =   705
      Left            =   3960
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   2640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Height          =   420
      Left            =   3960
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3960
      Top             =   3960
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
      Left            =   3960
      ScaleHeight     =   345
      ScaleWidth      =   1425
      TabIndex        =   30
      Top             =   1680
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
      Index           =   2
      Left            =   3960
      ScaleHeight     =   345
      ScaleWidth      =   1425
      TabIndex        =   29
      Top             =   1200
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
      Left            =   3960
      ScaleHeight     =   345
      ScaleWidth      =   1425
      TabIndex        =   28
      Top             =   720
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
      Left            =   3960
      ScaleHeight     =   345
      ScaleWidth      =   1425
      TabIndex        =   27
      Top             =   240
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
      Height          =   5055
      Left            =   120
      ScaleHeight     =   5025
      ScaleWidth      =   3705
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.TextBox txt 
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
         Index           =   12
         Left            =   840
         TabIndex        =   26
         Top             =   4680
         Width           =   2775
      End
      Begin VB.TextBox txt 
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
         Index           =   11
         Left            =   840
         TabIndex        =   24
         Top             =   4320
         Width           =   2775
      End
      Begin VB.TextBox txt 
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
         Index           =   10
         Left            =   840
         TabIndex        =   23
         Top             =   3960
         Width           =   2775
      End
      Begin VB.TextBox txt 
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
         Index           =   9
         Left            =   840
         TabIndex        =   22
         Top             =   3600
         Width           =   2775
      End
      Begin VB.TextBox txt 
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
         Index           =   8
         Left            =   840
         TabIndex        =   21
         Top             =   3240
         Width           =   2775
      End
      Begin VB.TextBox txt 
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
         Index           =   7
         Left            =   840
         TabIndex        =   20
         Top             =   2880
         Width           =   2775
      End
      Begin VB.TextBox txt 
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
         Index           =   6
         Left            =   840
         TabIndex        =   14
         Top             =   2520
         Width           =   2775
      End
      Begin VB.TextBox txt 
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
         Index           =   5
         Left            =   840
         TabIndex        =   13
         Top             =   2160
         Width           =   2775
      End
      Begin VB.TextBox txt 
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
         Index           =   4
         Left            =   840
         TabIndex        =   12
         Top             =   1800
         Width           =   2775
      End
      Begin VB.TextBox txt 
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
         Index           =   3
         Left            =   840
         TabIndex        =   8
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox txt 
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
         Index           =   2
         Left            =   840
         TabIndex        =   6
         Top             =   1080
         Width           =   2775
      End
      Begin VB.TextBox txt 
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
         Index           =   1
         Left            =   840
         TabIndex        =   4
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox txt 
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
         Index           =   0
         Left            =   840
         TabIndex        =   3
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Rcvt.cfg"
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
         TabIndex        =   25
         Top             =   4680
         Width           =   615
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Rcvt"
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
         TabIndex        =   19
         Top             =   4320
         Width           =   345
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Rar"
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
         TabIndex        =   18
         Top             =   3960
         Width           =   255
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Zip2exe"
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
         TabIndex        =   17
         Top             =   3600
         Width           =   570
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pkzipfix"
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
         TabIndex        =   16
         Top             =   3240
         Width           =   540
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pkunzip"
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
         TabIndex        =   15
         Top             =   2880
         Width           =   570
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pkzip"
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
         TabIndex        =   11
         Top             =   2520
         Width           =   390
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pak"
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
         Top             =   2160
         Width           =   285
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Lha"
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
         TabIndex        =   9
         Top             =   1800
         Width           =   270
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Rearj.cfg"
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
         TabIndex        =   7
         Top             =   1440
         Width           =   645
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Rearj"
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
         TabIndex        =   5
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Arjsort"
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
         TabIndex        =   2
         Top             =   720
         Width           =   435
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Arj"
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
         Top             =   360
         Width           =   180
      End
   End
   Begin VB.Label DirInSearch 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "DirInSearch"
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
      TabIndex        =   32
      Top             =   5520
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label Buscando 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Buscando"
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
      TabIndex        =   31
      Top             =   5280
      Visible         =   0   'False
      Width           =   720
   End
End
Attribute VB_Name = "fsetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Estado As Integer
Dim Compress(0 To 12), Dirs(0 To 12), Patron(0 To 4) As String


Private Sub btn_Click(Index As Integer)
    Select Case Index
    Case 0: GrabarSetup
    Case 1: Estado = 1: timer1.Enabled = -1
    Case 2: Unload fsetup
    Case 3: Ret = WinHelp(hWnd, PATH_HELP, HELP_CONTENTS, 3010&)
    End Select
End Sub

Private Sub btn_KeyPress(Index As Integer, keyascii As Integer)
    If keyascii = 13 Then
    keyascii = 0
    Select Case Index
    Case 0: GrabarSetup
    Case 1: Estado = 1: timer1.Enabled = -1
    Case 2: Unload fsetup
    Case 3: Ret = WinHelp(hWnd, PATH_HELP, HELP_CONTENTS, 3010&)
    End Select
    End If
End Sub

Private Function FileExits(ByVal FileToCheck As String) As Integer
    On Error Resume Next

    Open FileToCheck For Binary As #1
    If Err Or LOF(1) = 0 Then FileExits = 0 Else FileExits = -1
    Close #1

End Function

Private Sub FindFiles(PathInSearch As String)
    Dim i As Integer
    Dim Ruta As String
    File1.Path = Dir1.Path
    Ruta = Path(fsetup.Dir1)
    For j = 0 To 4
        File1.Pattern = Patron(j)
        File1.Refresh
        If File1.ListCount > 0 Then
            For i = 0 To 12
                If Found(Compress(i)) Then txt(i).Text = Ruta + Compress(i)
            Next i
        End If
    Next j
End Sub

Private Sub Form_Load()
    CentrarForma fsetup
    ' Get instance handle of project
    hInstance = GetInstance(Me)
    ' start the 3D functions
    Start3D
    ' show this form as 3D
    Make3D Me
    LLenarTextos
    Compress(0) = "arj.exe"
    Compress(1) = "arjsort.com"
    Compress(2) = "rearj.exe"
    Compress(3) = "rearj.cfg"
    Compress(4) = "lha.exe"
    Compress(5) = "pak.exe"
    Compress(6) = "pkzip.exe"
    Compress(7) = "pkunzip.exe"
    Compress(8) = "pkzipfix.exe"
    Compress(9) = "zip2exe.exe"
    Compress(10) = "rar.exe"
    Compress(11) = "rcvt.exe"
    Compress(12) = "rcvt.cfg"

    Patron(0) = "a*.*"
    Patron(1) = "r*.*"
    Patron(2) = "l*.*"
    Patron(3) = "p*.*"
    Patron(4) = "z*.*"

End Sub

Private Function Found(ByVal Archivo As String) As Integer
    Dim nInf As Integer
    Dim nSup, i As Single
    Dim File As String
    nInf = 0: nSup = File1.ListCount - 1
    Screen.MousePointer = 11
    Do While nInf <= nSup
        i = Int((nInf + nSup) / 2)
        File1.ListIndex = i
        File = File1.List(i)
        If Archivo = File Then Found = True: Screen.MousePointer = 0: Exit Function
        If File < Archivo Then nInf = i + 1 Else nSup = i - 1
    Loop
    Found = False
    Screen.MousePointer = 0
End Function

Private Sub GrabarSetup()
    Dim i, SW As Integer
    
    Msg = ""
    For i = 0 To 12: Dirs(i) = txt(i).Text: Next i
    
    For i = 0 To 12
        SW = i
        If Dirs(i) <> "" Then If Not FileExits(Dirs(i)) Then GoTo Salir
    Next i
    DIR_DE_ARJ = txt(0).Text
    DIR_DE_ARJSORT = txt(1).Text
    DIR_DE_REARJ = txt(2).Text
    DIR_DE_REARJ_CFG = txt(3).Text
    DIR_DE_LHA = txt(4).Text
    DIR_DE_PAK = txt(5).Text
    DIR_DE_PKZIP = txt(6).Text
    DIR_DE_PKUNZIP = txt(7).Text
    DIR_DE_PKZIPFIX = txt(8).Text
    DIR_DE_ZIP2EXE = txt(9).Text
    DIR_DE_RAR = txt(10).Text
    DIR_DE_RCVT = txt(11).Text
    DIR_DE_RCVT_CFG = txt(12).Text
    File1.Pattern = "*.*"
    GrabarInformacionAIniFile
    WinJob.Usar.Clear
    FillFiles
    Unload fsetup
Exit Sub
Salir:
    Msg = "Error de path o nombre en el archivo:" + CRLF
    Msg = Msg + Dirs(SW) + CRLF
    Msg = Msg + CRLF
    Msg = Msg + "El path o el nombre del archivo " + CRLF
    Msg = Msg + "escrito esta incorrecto." + CRLF
    Msg = Msg + "Digite el path o el nombre del archivo" + CRLF
    Msg = Msg + "según las normas del DOS." + CRLF
    Msg = Msg + "Si no sabe donde se encuentran sus" + CRLF
    Msg = Msg + "archivos haga clic en el botón" + CRLF
    Msg = Msg + """Buscar""." + CRLF
    MsgBox Msg, 48, "Ready"
    txt(SW).SetFocus
End Sub

Private Sub ListSubDirs(Path)
    Dim Count, D(), i, DirName
    Dim PathFind As String
    
    Screen.MousePointer = 11
    DirName = Dir(Path, ATTR_DIRECTORY)
    Dir1.Path = Path
    
    Do While DirName <> ""
    If DirName <> "." And DirName <> ".." Then
    If GetAttr(Path + DirName) = ATTR_DIRECTORY Then
    If (Count Mod 10) = 0 Then
        ReDim Preserve D(Count + 10)
    End If
    Count = Count + 1
    D(Count) = LCase(DirName)
    End If
    End If
    DirName = Dir
    Loop
    
    For i = 1 To Count
    PathFind = Path & D(i)
    DirInSearch.Caption = PathFind
    DirInSearch.Visible = True
    FindFiles PathFind
    ListSubDirs Path & D(i) & "\"
    Next i
    Screen.MousePointer = 0
End Sub

Private Sub LLenarTextos()
    txt(0).Text = DIR_DE_ARJ
    txt(1).Text = DIR_DE_ARJSORT
    txt(2).Text = DIR_DE_REARJ
    txt(3).Text = DIR_DE_REARJ_CFG
    txt(4).Text = DIR_DE_LHA
    txt(5).Text = DIR_DE_PAK
    txt(6).Text = DIR_DE_PKZIP
    txt(7).Text = DIR_DE_PKUNZIP
    txt(8).Text = DIR_DE_PKZIPFIX
    txt(9).Text = DIR_DE_ZIP2EXE
    txt(10).Text = DIR_DE_RAR
    txt(11).Text = DIR_DE_RCVT
    txt(12).Text = DIR_DE_RCVT_CFG
End Sub

Private Sub Timer1_Timer()
    If Estado = 1 Then
    For i = 0 To 3: btn(i).Enabled = 0: Next i
    Buscando.Visible = True
    ListSubDirs LCase(Left(App.Path, 3))
    End If
    For i = 0 To 3: btn(i).Enabled = -1: Next i
    timer1.Enabled = False
    Buscando.Visible = 0
    DirInSearch.Visible = 0
End Sub

