VERSION 5.00
Begin VB.Form finfor 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Información"
   ClientHeight    =   5490
   ClientLeft      =   105
   ClientTop       =   570
   ClientWidth     =   9075
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
   HelpContextID   =   1120
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5490
   ScaleWidth      =   9075
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
      Height          =   1455
      Left            =   0
      ScaleHeight     =   1425
      ScaleWidth      =   8985
      TabIndex        =   1
      Top             =   3960
      Width           =   9015
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "S = Sistema : Identifica al archivo como archivo de sistema de MS-DOS."
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
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   5100
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "O = Oculto. Impide que el archivo aparezca en una lista directorios a continuación del símbolo del sistema MS-DOS."
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
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   8190
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "N = Normal. El archivo puede ser modificado, borrado, leido,etc. "
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
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   4560
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "L = Archivo de solo lectura. Evita que el archivo pueda ser modificado."
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
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   4995
      End
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   225
      Left            =   3600
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox MList1 
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
      ScaleWidth      =   9075
      TabIndex        =   0
      Top             =   360
      Width           =   9075
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fecha y hora"
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
      Left            =   7320
      TabIndex        =   10
      Top             =   120
      Width           =   930
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tamaño en Bytes"
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
      Left            =   5880
      TabIndex        =   9
      Top             =   120
      Width           =   1245
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tipo de archivo"
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
      TabIndex        =   8
      Top             =   120
      Width           =   1110
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Directorio del archivo"
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
      TabIndex        =   7
      Top             =   120
      Width           =   1500
   End
   Begin VB.Label Label1 
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
      Left            =   480
      TabIndex        =   6
      Top             =   120
      Width           =   1380
   End
End
Attribute VB_Name = "finfor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const ATTR_READONLY = 1, ATTR_HIDDEN = 2
Const ATTR_SYSTEM = 4, ATTR_ARCHIVE = 32

Private Sub Cerrar_Click()
    Unload finfor
End Sub

Private Sub Form_Load()
    Dim i As Single
    Dim Nombre, Origen2, Tamano, Fecha, Origen As String
    Dim Attr, Fname, Tipo
    CentrarForma finfor
    List1.Clear
    Mlist1.Clear
    Origen = Path(WinJob.Dir1)
    Select Case WinJob.List1.ListIndex
        Case -1: For i = 0 To WinJob.List1.ListCount - 1: List1.AddItem WinJob.List1.List(i): Next i
        Case Else
        For i = 0 To WinJob.List1.ListCount - 1
            If WinJob.List1.Selected(i) Then List1.AddItem WinJob.List1.List(i)
        Next i
    End Select
    For i = 0 To List1.ListCount - 1
        Fname = Origen + List1.List(i)
        Attr = GetAttr(Fname)
        If Attr > 7 Then Attr = Attr Xor ATTR_ARCHIVE
        Select Case Attr
            Case 0: Tipo = "N     "
            Case ATTR_READONLY: Tipo = "L     "
            Case ATTR_HIDDEN: Tipo = "O     "
            Case ATTR_HIDDEN + ATTR_READONLY: Tipo = "LO    "
            Case ATTR_SYSTEM: Tipo = "S     "
            Case ATTR_READONLY + ATTR_SYSTEM: Tipo = "LS    "
            Case ATTR_HIDDEN + ATTR_SYSTEM: Tipo = "OS    "
            Case ATTR_READONLY + ATTR_HIDDEN + ATTR_SYSTEM: Tipo = "LOS   "
        End Select
        Nombre = List1.List(i)
        Origen2 = Origen
        If Len(Nombre) < 12 Then Nombre = Nombre + Space(12 - Len(Nombre))
        Origen2 = Right(Origen2, 23)
        If Len(Origen2) < 25 Then Origen2 = Origen2 + Space(26 - Len(Origen2))
        Tamano = LTrim(Str(FileLen(Fname)))
        If Len(Tamano) < 9 Then Tamano = Tamano + Space(12 - Len(Tamano))
        Fecha = FileDateTime(Fname)
        Fecha = Format(Fecha, "dd-mm-yy") & " "
        Fecha = Fecha & Format(Fecha, "h:mm AM/PM")
        Mlist1.AddItem Nombre + " " + Origen2 + Tipo + Tamano + Fecha
    Next i
    List1.Clear
    hInstance = GetInstance(Me)
    ' start the 3D functions
    Start3D
    ' show this form as 3D
    Make3D Me

End Sub

