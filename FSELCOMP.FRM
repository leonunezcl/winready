VERSION 5.00
Begin VB.Form fselcomp 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleccione el descompresor"
   ClientHeight    =   1920
   ClientLeft      =   2295
   ClientTop       =   2820
   ClientWidth     =   4035
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
   ScaleHeight     =   1920
   ScaleWidth      =   4035
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
      Left            =   2760
      ScaleHeight     =   345
      ScaleWidth      =   1185
      TabIndex        =   8
      Top             =   1440
      Width           =   1215
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
      Left            =   1440
      ScaleHeight     =   345
      ScaleWidth      =   1185
      TabIndex        =   7
      Top             =   1440
      Width           =   1215
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
      Left            =   120
      ScaleHeight     =   345
      ScaleWidth      =   1185
      TabIndex        =   6
      Top             =   1440
      Width           =   1215
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
      Height          =   1215
      Left            =   120
      ScaleHeight     =   1185
      ScaleWidth      =   3825
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      Begin VB.PictureBox Option3D 
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
         Left            =   3120
         ScaleHeight     =   225
         ScaleWidth      =   585
         TabIndex        =   5
         Top             =   360
         Width           =   615
      End
      Begin VB.PictureBox Option3D 
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
         Left            =   1680
         ScaleHeight     =   225
         ScaleWidth      =   945
         TabIndex        =   4
         Top             =   840
         Width           =   975
      End
      Begin VB.PictureBox Option3D 
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
         Left            =   1680
         ScaleHeight     =   225
         ScaleWidth      =   945
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.PictureBox Option3D 
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
         ScaleWidth      =   585
         TabIndex        =   2
         Top             =   840
         Width           =   615
      End
      Begin VB.PictureBox Option3D 
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
         ScaleWidth      =   705
         TabIndex        =   1
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "fselcomp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btn_Click(Index As Integer)
    Select Case Index
        Case 0: nSELCOMP = 0: GoSub VerExt: Unload fselcomp
        Case 1: nSELCOMP = -1: Unload fselcomp
        Case 2
    End Select
Exit Sub
VerExt:
    MsgBox Ruta
    If Option3D(0).Value = -1 Then Start = DIR_DE_ARJ + " l " + Ruta + " > " + PATH_LIST: Ext = "Arj"
    If Option3D(1).Value = -1 Then Start = DIR_DE_LHA + " l " + Ruta + " > " + PATH_LIST: Ext = "Lha"
    If Option3D(2).Value = -1 Then Start = DIR_DE_PAK + " t " + Ruta + " > " + PATH_LIST: Ext = "Pak"
    If Option3D(3).Value = -1 Then Start = DIR_DE_PKUNZIP + " " + Ruta + " -@" + PATH_LIST: Ext = "Pkunzip"
    If Option3D(4).Value = -1 Then Start = DIR_DE_RAR + " l " + Ruta + " > " + PATH_LIST: Ext = "Rar"
Return
End Sub

Private Sub Form_Load()
    CentrarForma fselcomp
    Option3D(0).Enabled = IIf(DIR_DE_ARJ <> "", -1, 0)
    Option3D(1).Enabled = IIf(DIR_DE_LHA <> "", -1, 0)
    Option3D(2).Enabled = IIf(DIR_DE_PAK <> "", -1, 0)
    Option3D(3).Enabled = IIf(DIR_DE_PKUNZIP <> "", -1, 0)
    Option3D(4).Enabled = IIf(DIR_DE_RAR <> "", -1, 0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    nSELCOMP = 0
End Sub

