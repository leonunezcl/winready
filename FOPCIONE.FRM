VERSION 5.00
Begin VB.Form fopciones 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Opciones"
   ClientHeight    =   3315
   ClientLeft      =   630
   ClientTop       =   1875
   ClientWidth     =   7680
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
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3315
   ScaleWidth      =   7680
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
      Left            =   6120
      ScaleHeight     =   345
      ScaleWidth      =   1425
      TabIndex        =   20
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
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   6120
      ScaleHeight     =   345
      ScaleWidth      =   1425
      TabIndex        =   19
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
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   6120
      ScaleHeight     =   345
      ScaleWidth      =   1425
      TabIndex        =   18
      Top             =   240
      Width           =   1455
   End
   Begin VB.PictureBox CMDialog1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   6240
      ScaleHeight     =   450
      ScaleWidth      =   1170
      TabIndex        =   21
      Top             =   1800
      Width           =   1200
   End
   Begin VB.PictureBox Panel 
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
      Height          =   975
      Index           =   3
      Left            =   3240
      ScaleHeight     =   945
      ScaleWidth      =   2745
      TabIndex        =   15
      Top             =   2160
      Width           =   2775
      Begin VB.ComboBox Tipo_com 
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
         TabIndex        =   17
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label lblopciones 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo de compresión"
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
         TabIndex        =   16
         Top             =   360
         Width           =   1395
      End
   End
   Begin VB.PictureBox Panel 
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
      Height          =   1095
      Index           =   2
      Left            =   3240
      ScaleHeight     =   1065
      ScaleWidth      =   2745
      TabIndex        =   12
      Top             =   960
      Width           =   2775
      Begin VB.ComboBox Orden 
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
         TabIndex        =   14
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label lblopciones 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ordenar por"
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
         TabIndex        =   13
         Top             =   360
         Width           =   840
      End
   End
   Begin VB.PictureBox Panel 
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
      Left            =   3240
      ScaleHeight     =   705
      ScaleWidth      =   2745
      TabIndex        =   10
      Top             =   120
      Width           =   2775
      Begin VB.PictureBox Arj 
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
         ScaleWidth      =   2025
         TabIndex        =   11
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.PictureBox Panel 
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
      Height          =   3015
      Index           =   0
      Left            =   120
      ScaleHeight     =   2985
      ScaleWidth      =   2985
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      Begin VB.PictureBox General 
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
         ScaleWidth      =   2145
         TabIndex        =   9
         Top             =   2640
         Width           =   2175
      End
      Begin VB.ComboBox Tamano 
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
         TabIndex        =   8
         Top             =   2280
         Width           =   1695
      End
      Begin VB.PictureBox General 
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
         Left            =   120
         ScaleHeight     =   225
         ScaleWidth      =   2025
         TabIndex        =   6
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox txtcontra 
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
         Top             =   1360
         Width           =   2775
      End
      Begin VB.PictureBox General 
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
         ScaleWidth      =   1545
         TabIndex        =   4
         Top             =   1080
         Width           =   1575
      End
      Begin VB.PictureBox General 
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
         ScaleWidth      =   1785
         TabIndex        =   3
         Top             =   840
         Width           =   1815
      End
      Begin VB.PictureBox General 
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
         ScaleWidth      =   2265
         TabIndex        =   2
         Top             =   600
         Width           =   2295
      End
      Begin VB.PictureBox General 
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
         ScaleWidth      =   2745
         TabIndex        =   1
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label lblopciones 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Múltiples disquettes"
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
         TabIndex        =   7
         Top             =   2040
         Width           =   1380
      End
   End
End
Attribute VB_Name = "fopciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Multi(0 To 4) As String
Dim Tipo(0 To 4) As String
Dim ArjOp(0 To 7) As String
Dim LhaOp(0 To 5) As String
Dim PakOp(0 To 5) As String
Dim PkzipOp(0 To 7) As String
Dim RarOp(0 To 6) As String
Dim Ordenar(0 To 4) As String

Private Sub AcoplarArj()
    Dim i As Integer
    ArjOp(0) = " -a "
    ArjOp(1) = " -je "
    ArjOp(2) = " -z" + COMENTARIO
    ArjOp(3) = " -g" + txtcontra.Text
    ArjOp(4) = " -r "
    ArjOp(5) = " -e1 "
    ArjOp(6) = Right(Multi(Tamano.ListIndex), 8)
    ArjOp(7) = " -jt "
    
    For i = 0 To 5
        ARJ_TAB(i) = IIf(General(i).Value = -1, ArjOp(i), "")
    Next i
    PASSWORD = IIf(General(3).Value = -1, -1, 0)
    CONTRASENA = IIf(General(3).Value = -1, txtcontra.Text, "")
    ARJ_TAB(6) = ArjOp(6)
    ARJ_TAB(7) = IIf(Arj.Value = -1, ArjOp(7), "")
    For i = 0 To 7: PREDEF_ARJ = PREDEF_ARJ + ARJ_TAB(i): Next i
    Paso = PREDEF_ARJ
    Unload fopciones
End Sub

Private Sub AcoplarLha()
    Dim i As Integer
    LhaOp(0) = " -a1 "
    If General(1).Value = -1 Then
        LhaOp(1) = " s "
    Else
        LhaOp(1) = IIf(F_COMPRIMIR = -1, " a ", " m ")
    End If
    LhaOp(2) = " "
    LhaOp(3) = " "
    LhaOp(4) = " -r2x1 "
    LhaOp(5) = " -x0 "
    
    For i = 0 To 5
        LHA_TAB(i) = IIf(General(i).Value = -1, LhaOp(i), "")
    Next i
    PREDEF_LHA = PREDEF_LHA + LHA_TAB(0)
    For i = 2 To 5: PREDEF_LHA = PREDEF_LHA + LHA_TAB(i): Next i
    Paso = PREDEF_LHA
    Unload fopciones
End Sub

Private Sub AcoplarPak()
    Dim i As Integer
    
    PakOp(0) = " "
    PakOp(1) = " /exe "
    PakOp(2) = " "
    PakOp(3) = " /g=" + txtcontra.Text
    PakOp(4) = " /I "
    PakOp(5) = Right(Ordenar(Orden.ListIndex), 5)
    
    For i = 0 To 4
        PAK_TAB(i) = IIf(General(i).Value = -1, PakOp(i), "")
    Next i
    PAK_TAB(5) = PakOp(5)
    PASSWORD = IIf(General(3).Value = -1, -1, 0)
    CONTRASENA = IIf(General(3).Value = -1, txtcontra.Text, "")
    For i = 0 To 5: PREDEF_PAK = PREDEF_PAK + PAK_TAB(i): Next i
    Paso = PREDEF_PAK
    Unload fopciones
End Sub

Private Sub AcoplarPkzip()
    Dim i As Integer
    
    PkzipOp(0) = " -whs "
    PkzipOp(1) = " "
    PkzipOp(2) = " -z < " + COMENTARIO
    PkzipOp(3) = " -s" + txtcontra.Text
    PkzipOp(4) = " -rp "
    PkzipOp(5) = " -p "
    PkzipOp(6) = Right(Tipo(Tipo_com.ListIndex), 5)
    PkzipOp(7) = IIf(Tamano.Enabled = -1, Right(Multi(Tamano.ListIndex), 6), "")
    
    For i = 0 To 5
        PKZIP_TAB(i) = IIf(General(i).Value = -1, PkzipOp(i), "")
    Next i
    PASSWORD = IIf(General(3).Value = -1, -1, 0)
    CONTRASENA = IIf(General(3).Value = -1, txtcontra.Text, "")
    PKZIP_TAB(6) = PkzipOp(6)
    PKZIP_TAB(7) = PkzipOp(7)
    For i = 0 To 7: PREDEF_PKZIP = PREDEF_PKZIP + PKZIP_TAB(i): Next i
    Paso = PREDEF_PKZIP
    Unload fopciones
End Sub

Private Sub AcoplarPredef()
    Select Case COMPRESS_ACTIVE
        Case "Arj": AcoplarArj
        Case "Lha": AcoplarLha
        Case "Pak": AcoplarPak
        Case "Pkzip": AcoplarPkzip
        Case "Rar": AcoplarRar
    End Select
End Sub

Private Sub AcoplarRar()
    Dim i As Integer
    RarOp(0) = " "
    RarOp(1) = " -sfx "
    RarOp(2) = " -z " + COMENTARIO
    RarOp(3) = " -p" + txtcontra.Text
    RarOp(4) = " -r "
    RarOp(5) = " -ep1 "
    RarOp(6) = Right(Multi(Tamano.ListIndex), 8)

    For i = 0 To 5
        RAR_TAB(i) = IIf(General(i).Value = -1, RarOp(i), "")
    Next i
    RAR_TAB(6) = RarOp(6)
    PASSWORD = IIf(General(3).Value = -1, -1, 0)
    CONTRASENA = IIf(General(3).Value = -1, txtcontra.Text, "")
    For i = 0 To 6: PREDEF_RAR = PREDEF_RAR + RAR_TAB(i): Next i
    Paso = PREDEF_RAR
    Unload fopciones
End Sub

Private Sub btn_Click(Index As Integer)
    Select Case Index
        Case 0: AcoplarPredef
        Case 1: Unload fopciones
        Case 2: Ret = WinHelp(hWnd, PATH_HELP, HELP_CONTENTS, 7&)
    End Select
End Sub

Private Sub btn_KeyPress(Index As Integer, keyascii As Integer)
    If keyascii = 13 Then
        keyascii = 0
    Select Case Index
        Case 0: AcoplarPredef
        Case 1: Unload fopciones
        Case 2: Ret = WinHelp(hWnd, PATH_HELP, HELP_CONTENTS, 7&)
    End Select
    End If
End Sub

Private Sub ComentarArchivo()
    On Error Resume Next
    If COMENTARIO <> "" Then Exit Sub
    CMDialog1.DialogTitle = "Agregar comentario"
    CMDialog1.Filter = "Archivos de texto (*.txt)| *.txt| Todos los archivos (*.*)| *.*"
    CMDialog1.Action = 1
    If Err <> 32755 Then
        COMENTARIO = CMDialog1.FILENAME
    Else
        General(2).Value = 0
    End If
End Sub

Private Function ExistsDrive(ByVal Drive_Search As String) As Integer
    Dim i As Integer
    For i = 0 To TOT_UNIDADES
        If UNIDADES(i) = Drive_Search Then ExistsDrive = -1: Exit Function
    Next i
    ExistsDrive = 0
End Function

Private Sub Form_Activate()
    Select Case COMPRESS_ACTIVE
        Case "Arj": RestaurarArj
        Case "Lha": RestaurarLha
        Case "Pak": RestaurarPak
        Case "Pkzip": RestaurarPkzip
        Case "Rar": RestaurarRar
    End Select
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim path_to_compress, Path_ready As String
    
    path_to_compress = Left(fcomprimir.txthacia.Text, 2)
    Path_ready = LCase(Left(App.Path, 2))
    
    CentrarForma fopciones
    Select Case COMPRESS_ACTIVE
        Case "Arj": For i = 2 To 3: Panel(i).Enabled = 0: Next i
                    For i = 1 To 2: lblopciones(i).Enabled = 0: Next i
                    If Right(path_to_compress, 1) = ":" Then
                        lblopciones(0).Enabled = IIf(Not ExistsDrive(path_to_compress), 0, -1)
                        Tamano.Enabled = IIf(lblopciones(0).Enabled = -1, -1, 0)
                    End If
                    Orden.Enabled = 0
                    Tipo_com.Enabled = 0
        Case "Lha": For i = 1 To 3: Panel(i).Enabled = 0: Next i
                    For i = 2 To 3: General(i).Enabled = 0: Next i
                    For i = 0 To 2: lblopciones(i).Enabled = 0: Next i
                    Tamano.Enabled = 0
                    Arj.Enabled = 0
                    Orden.Enabled = 0
                    Tipo_com.Enabled = 0
        Case "Pak": For i = 1 To 3 Step 2: Panel(i).Enabled = 0: Next i
                    For i = 0 To 2 Step 2: General(i).Enabled = 0: Next i
                    General(5).Enabled = 0
                    For i = 0 To 2 Step 2: lblopciones(i).Enabled = 0: Next i
                    Tamano.Enabled = 0
                    Arj.Enabled = 0
                    Tipo_com.Enabled = 0
        Case "Pkzip": For i = 1 To 2: Panel(i).Enabled = 0: Next i
                    General(1).Enabled = 0
                    If Right(path_to_compress, 1) = ":" Then
                        lblopciones(0).Enabled = IIf(Path_ready = path_to_compress Or Not ExistsDrive(path_to_compress), 0, -1)
                        Tamano.Enabled = IIf(lblopciones(0).Enabled = -1, -1, 0)
                    Else
                        lblopciones(0).Enabled = 0
                        Tamano.Enabled = 0
                    End If
                    Arj.Enabled = 0
                    Orden.Enabled = 0
                    lblopciones(1).Enabled = 0
        Case "Rar": For i = 1 To 3: Panel(i).Enabled = 0: Next i
                    If Right(path_to_compress, 1) = ":" Then
                        lblopciones(0).Enabled = IIf(Not ExistsDrive(path_to_compress), 0, -1)
                        Tamano.Enabled = IIf(lblopciones(0).Enabled = -1, -1, 0)
                    End If
                    Arj.Enabled = 0
                    For i = 1 To 2: lblopciones(i).Enabled = 0: Next i
                    Orden.Enabled = 0
                    Tipo_com.Enabled = 0
    End Select
                           
    
    Tipo(0) = "Solo juntar  -e0 "
    Tipo(1) = "Super rápida -es "
    Tipo(2) = "Rápida       -ef "
    Tipo(3) = "Normal       -en "
    Tipo(4) = "Máxima       -ex "
                  
    Ordenar(0) = "Nombre     /sn "
    Ordenar(1) = "Extensión  /se "
    Ordenar(2) = "Hora       /st "
    Ordenar(3) = "Tamaño     /ss "
    Ordenar(4) = "No ordenar /ns "
    
    Select Case COMPRESS_ACTIVE
        Case "Arj", "Rar", "Pak", "Lha"
            Multi(0) = "360             -v360  "
            Multi(1) = "720             -v720  "
            Multi(2) = "1220            -v1220 "
            Multi(3) = "1440            -v1440 "
            Multi(4) = "Detectar        -va    "
            GoSub Llenar
        Case Else
            Multi(0) = "No formatear    -&   "
            Multi(1) = "Borrar archivos -&w  "
            Multi(2) = "Formato rápido  -&fl "
            Multi(3) = "Full formato    -&fu "
            Multi(4) = "Detectar        -&   "
            GoSub Llenar
    End Select
    hInstance = GetInstance(Me)
    ' start the 3D functions
    Start3D
    ' show this form as 3D
    Make3D Me

Exit Sub
Llenar:
    For i = 0 To 4
        Tamano.AddItem Left(Multi(i), 15)
        Tipo_com.AddItem Left(Tipo(i), 13)
        Orden.AddItem Left(Ordenar(i), 10)
    Next i
    Tamano.ListIndex = 4
    Tipo_com.ListIndex = 4
    Orden.ListIndex = 4
Return
End Sub

Private Sub General_Click(Index As Integer, Value As Integer)
    Select Case Index
        Case 2: If Value = -1 Then ComentarArchivo Else COMENTARIO = ""
        Case 3:
        Select Case Value
            Case -1: txtcontra.Enabled = -1: txtcontra.SetFocus
            Case 0: txtcontra.Enabled = 0: 'txtcontra.Text = "": btn(0).SetFocus
        End Select
    End Select
End Sub

Private Sub RestaurarArj()
    Dim i As Integer
    For i = 0 To 5
        General(i).Value = IIf(ARJ_TAB(i) <> "", -1, 0)
    Next i
    If ARJ_TAB(2) = "" Then COMENTARIO = ""
    txtcontra.Enabled = IIf(ARJ_TAB(3) <> "", -1, 0)
    txtcontra.Text = IIf(ARJ_TAB(3) <> "", CONTRASENA, "")
    
    Select Case ARJ_TAB(6)
        Case " -v360  ": Tamano.ListIndex = 0
        Case " -v720  ": Tamano.ListIndex = 1
        Case " -v1220 ": Tamano.ListIndex = 2
        Case " -v1440 ": Tamano.ListIndex = 3
        Case " -va    ": Tamano.ListIndex = 4
    End Select
    Arj.Value = IIf(ARJ_TAB(7) <> "", -1, 0)
End Sub

Private Sub RestaurarLha()
    Dim i As Integer
    For i = 0 To 5
        General(i).Value = IIf(LHA_TAB(i) <> "", -1, 0)
    Next i
End Sub

Private Sub RestaurarPak()
    Dim i As Integer
    For i = 0 To 4
        General(i).Value = IIf(PAK_TAB(i) <> "", -1, 0)
    Next i
    txtcontra.Enabled = IIf(PAK_TAB(3) <> "", -1, 0)
    txtcontra.Text = IIf(PAK_TAB(3) <> "", CONTRASENA, "")
    Select Case PAK_TAB(5)
        Case " /sn ": Orden.ListIndex = 0
        Case " /se ": Orden.ListIndex = 1
        Case " /st ": Orden.ListIndex = 2
        Case " /ss ": Orden.ListIndex = 3
        Case " /ns ": Orden.ListIndex = 4
    End Select
End Sub

Private Sub RestaurarPkzip()
    Dim i As Integer
    For i = 0 To 5
        General(i).Value = IIf(PKZIP_TAB(i) <> "", -1, 0)
    Next i
    If PKZIP_TAB(2) = "" Then COMENTARIO = ""
    txtcontra.Enabled = IIf(PKZIP_TAB(3) <> "", -1, 0)
    txtcontra.Text = IIf(PKZIP_TAB(3) <> "", CONTRASENA, "")
    
    Select Case PKZIP_TAB(6)
        Case " -e0 ": Tipo_com.ListIndex = 0
        Case " -es ": Tipo_com.ListIndex = 1
        Case " -ef ": Tipo_com.ListIndex = 2
        Case " -en ": Tipo_com.ListIndex = 3
        Case " -ex ": Tipo_com.ListIndex = 4
    End Select
    Select Case PKZIP_TAB(7)
        Case " -&   ": Tamano.ListIndex = 0
        Case " -&w  ": Tamano.ListIndex = 1
        Case " -&fl ": Tamano.ListIndex = 2
        Case " -&fu ": Tamano.ListIndex = 3
        Case " -&   ": Tamano.ListIndex = 4
    End Select
End Sub

Private Sub RestaurarRar()
    Dim i As Integer
    For i = 0 To 5
        General(i).Value = IIf(RAR_TAB(i) <> "", -1, 0)
    Next i
    If RAR_TAB(2) = "" Then COMENTARIO = ""
    txtcontra.Enabled = IIf(RAR_TAB(3) <> "", -1, 0)
    txtcontra.Text = IIf(RAR_TAB(3) <> "", CONTRASENA, "")
    
    Select Case RAR_TAB(6)
        Case " -v360  ": Tamano.ListIndex = 0
        Case " -v720  ": Tamano.ListIndex = 1
        Case " -v1220 ": Tamano.ListIndex = 2
        Case " -v1440 ": Tamano.ListIndex = 3
        Case " -va    ": Tamano.ListIndex = 4
    End Select
End Sub

Private Sub txtcontra_Change()
    If txtcontra.Text = "" Then General(3).Value = 0
End Sub

Private Sub txtcontra_LostFocus()
    If txtcontra.Text = "" Then General(3).Value = 0
End Sub

