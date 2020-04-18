VERSION 2.00
Begin Form fmodulo 
   BackColor       =   &H8000000F&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Módulo de instalación"
   ClientHeight    =   3240
   ClientLeft      =   900
   ClientTop       =   2430
   ClientWidth     =   7200
   Height          =   3645
   HelpContextID   =   1140
   Left            =   840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   7200
   Top             =   2085
   Width           =   7320
   Begin SSPanel Panel3D1 
      Align           =   1  'Align Top
      Height          =   3255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7200
      Begin SSCommand btn 
         Caption         =   "&Aceptar"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   5640
         TabIndex        =   0
         Top             =   360
         Width           =   1455
      End
      Begin SSFrame Frame3D1 
         Caption         =   "Mensaje"
         ForeColor       =   &H00000000&
         Height          =   2655
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   5415
         Begin PictureBox Picture1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   480
            Left            =   120
            Picture         =   FMODULO.FRX:0000
            ScaleHeight     =   480
            ScaleWidth      =   480
            TabIndex        =   8
            Top             =   1680
            Width           =   480
         End
         Begin Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Avda Benozzo Gozolli # 5227, comuna de San Joaquín Chile o al telefono 5520818."
            Height          =   390
            Left            =   120
            TabIndex        =   5
            Top             =   2160
            Width           =   5055
            WordWrap        =   -1  'True
         End
         Begin Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Ready Software"
            FontBold        =   -1  'True
            FontItalic      =   0   'False
            FontName        =   "MS Sans Serif"
            FontSize        =   9.75
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   720
            TabIndex        =   7
            Top             =   1800
            Width           =   1665
         End
         Begin Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Para más información contactese a Ready Software a la siguiente dirección:"
            Height          =   390
            Left            =   120
            TabIndex        =   6
            Top             =   1080
            Width           =   5175
            WordWrap        =   -1  'True
         End
         Begin Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Usando módulos de instalación le permite a Ud. distribuir sus programas en formato comprimido desde uno a varios discos de una manera elegante y profesional."
            Height          =   585
            Left            =   120
            TabIndex        =   4
            Top             =   360
            Width           =   5250
            WordWrap        =   -1  'True
         End
      End
      Begin Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Disponible solo en la versión registrada"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   -1  'True
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   4170
      End
   End
End

Sub btn_Click ()
    Unload fmodulo
End Sub

Sub Form_Load ()
    CentrarForma fmodulo
End Sub

