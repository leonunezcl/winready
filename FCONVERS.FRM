VERSION 2.00
Begin Form fconversores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleccione el conversor"
   ClientHeight    =   2280
   ClientLeft      =   1200
   ClientTop       =   1350
   ClientWidth     =   6975
   Height          =   2685
   Left            =   1140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   6975
   Top             =   1005
   Width           =   7095
   Begin SSPanel Panel3D1 
      Align           =   1  'Align Top
      BackColor       =   &H00C0C0C0&
      Font3D          =   0  'None
      Height          =   2295
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6975
      Begin SSCommand btn 
         Caption         =   "A&yuda"
         Font3D          =   0  'None
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   2
         Left            =   5400
         TabIndex        =   5
         Top             =   1200
         Width           =   1455
      End
      Begin SSCommand btn 
         Caption         =   "&Cancelar"
         Font3D          =   0  'None
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   1
         Left            =   5400
         TabIndex        =   4
         Top             =   720
         Width           =   1455
      End
      Begin SSCommand btn 
         Caption         =   "&Aceptar"
         Font3D          =   0  'None
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   0
         Left            =   5400
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
      Begin SSOption Rearj 
         Caption         =   "R&earj"
         Font3D          =   0  'None
         Height          =   255
         Left            =   2880
         TabIndex        =   2
         Top             =   1920
         Width           =   735
      End
      Begin SSOption Rcvt 
         Caption         =   "&Rcvt"
         Font3D          =   0  'None
         Height          =   255
         Left            =   2160
         TabIndex        =   1
         Top             =   1920
         Width           =   735
      End
      Begin Label lbl 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Conversor a usar  :"
         Height          =   195
         Index           =   4
         Left            =   360
         TabIndex        =   0
         Top             =   1920
         Width           =   1635
      End
      Begin Label lbl 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "arj,lzh,rar,zip"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   1110
      End
      Begin Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Rcvt : Convierte a ""rar"" los siguientes formatos de compresión:"
         Height          =   495
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Top             =   120
         Width           =   4695
      End
      Begin Label lbl 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "lzh,pak,zip"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   1560
         Width           =   930
      End
      Begin Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Rearj : Convierte a ""arj"" los siguientes formatos de compresión:"
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   4335
      End
   End
End

Sub btn_Click (Index As Integer)
    Select Case Index
        Case 1: fconvertir.Show 1
        Case 2: Unload fconvertir
        Case 3
    End Select
End Sub

Sub btn_KeyPress (Index As Integer, keyAscii As Integer)
    If keyAscii = 13 Then
        keyAscii = 0
        Select Case Index
            Case 1: fconvertir.Show 1
            Case 2: Unload fconvertir
            Case 3
        End Select
    End If
End Sub

Sub Form_Load ()
    CentrarForma fconversores
    If DIR_DE_RCVT <> "" And DIR_DE_RCVT_CFG <> "" Then rcvt.Enabled = -1 Else rcvt.Enabled = 0
    If DIR_DE_REARJ <> "" And DIR_DE_REARJ_CFG <> "" Then rearj.Enabled = -1 Else rearj.Enabled = 0
End Sub

