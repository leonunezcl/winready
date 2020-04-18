VERSION 2.00
Begin Form fready 
   BackColor       =   &H8000000F&
   Caption         =   "WinReady"
   ClientHeight    =   3870
   ClientLeft      =   2700
   ClientTop       =   840
   ClientWidth     =   3645
   Height          =   4845
   HelpContextID   =   1
   Icon            =   FREADY.FRX:0000
   Left            =   2640
   LinkTopic       =   "Form1"
   ScaleHeight     =   540
   ScaleWidth      =   540
   Top             =   -75
   Width           =   3765
   WindowState     =   2  'Maximized
   Begin SSPanel panel 
      Align           =   2  'Align Bottom
      Alignment       =   1  'Left Justify - MIDDLE
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   390
      Index           =   2
      Left            =   0
      Outline         =   -1  'True
      RoundedCorners  =   0   'False
      TabIndex        =   6
      Top             =   3480
      Width           =   3645
      Begin Label lbllisto 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Listo"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   240
         Left            =   45
         TabIndex        =   51
         Top             =   60
         Width           =   420
      End
      Begin Label lblbytes 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "bytes"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   195
         Left            =   9000
         TabIndex        =   33
         Top             =   90
         Width           =   375
      End
      Begin Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Bytes:"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   195
         Left            =   7560
         TabIndex        =   34
         Top             =   90
         Width           =   435
      End
      Begin Line Line8 
         BorderColor     =   &H00FFFFFF&
         X1              =   9435
         X2              =   9435
         Y1              =   60
         Y2              =   320
      End
      Begin Line Line7 
         BorderColor     =   &H00FFFFFF&
         X1              =   7485
         X2              =   9435
         Y1              =   315
         Y2              =   315
      End
      Begin Line Line6 
         BorderColor     =   &H00808080&
         X1              =   7485
         X2              =   9435
         Y1              =   60
         Y2              =   60
      End
      Begin Line Line5 
         BorderColor     =   &H00808080&
         X1              =   7485
         X2              =   7485
         Y1              =   60
         Y2              =   320
      End
      Begin Label lbltot 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "archi"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   195
         Left            =   6990
         TabIndex        =   35
         Top             =   90
         Width           =   360
      End
      Begin Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Archivos:"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   195
         Left            =   5685
         TabIndex        =   36
         Top             =   90
         Width           =   660
      End
      Begin Line Line4 
         BorderColor     =   &H00FFFFFF&
         X1              =   7440
         X2              =   7440
         Y1              =   60
         Y2              =   320
      End
      Begin Line Line3 
         BorderColor     =   &H00FFFFFF&
         X1              =   5640
         X2              =   7440
         Y1              =   315
         Y2              =   315
      End
      Begin Line Line2 
         BorderColor     =   &H00808080&
         X1              =   5640
         X2              =   7440
         Y1              =   60
         Y2              =   60
      End
      Begin Line Line1 
         BorderColor     =   &H00808080&
         X1              =   5640
         X2              =   5640
         Y1              =   60
         Y2              =   320
      End
   End
   Begin SSPanel panel 
      Align           =   1  'Align Top
      Alignment       =   0  'Left Justify - TOP
      Height          =   4560
      Index           =   4
      Left            =   0
      TabIndex        =   39
      Top             =   1680
      Width           =   3645
      Begin ListBox List3 
         Height          =   4320
         Left            =   120
         MultiSelect     =   2  'Extended
         TabIndex        =   53
         Top             =   120
         Visible         =   0   'False
         Width           =   9375
      End
      Begin ListBox List1 
         Columns         =   5
         Height          =   4320
         Left            =   2280
         MultiSelect     =   2  'Extended
         TabIndex        =   52
         Top             =   120
         Width           =   7215
      End
      Begin FileListBox File1 
         Height          =   225
         Left            =   2280
         TabIndex        =   45
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin DirListBox Dir1 
         Height          =   4305
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   2055
      End
   End
   Begin SSPanel panel 
      Align           =   1  'Align Top
      Height          =   780
      Index           =   3
      Left            =   0
      TabIndex        =   32
      Top             =   900
      Width           =   3645
      Begin SpinButton Spin1 
         BackColor       =   &H00C0C0C0&
         BorderThickness =   0
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   9260
         ShadeColor      =   &H00000000&
         ShadowBackColor =   &H00000000&
         ShadowForeColor =   &H00000000&
         ShadowThickness =   1
         SpinForeColor   =   &H80000008&
         TdThickness     =   1
         Top             =   75
         Width           =   255
      End
      Begin PictureBox Picture2 
         Height          =   615
         Left            =   120
         ScaleHeight     =   585
         ScaleWidth      =   9045
         TabIndex        =   7
         Top             =   75
         Width           =   9080
         Begin PictureBox PicSuge 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   480
            Left            =   120
            Picture         =   FREADY.FRX:0302
            ScaleHeight     =   480
            ScaleWidth      =   480
            TabIndex        =   38
            Top             =   60
            Width           =   480
         End
         Begin Label LblSuge 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Sugerencias"
            FontBold        =   0   'False
            FontItalic      =   0   'False
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            Height          =   195
            Left            =   720
            TabIndex        =   37
            Top             =   120
            Width           =   8295
            WordWrap        =   -1  'True
         End
      End
   End
   Begin SSPanel panel 
      Align           =   1  'Align Top
      Height          =   450
      Index           =   1
      Left            =   0
      TabIndex        =   31
      Top             =   450
      Width           =   3645
      Begin SSCheck Restaurar 
         Caption         =   "Restaurar lista"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   5880
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   1575
      End
      Begin DriveListBox Drive1 
         Height          =   315
         Left            =   1080
         TabIndex        =   3
         Top             =   75
         Width           =   2295
      End
      Begin ComboBox Usar 
         Height          =   300
         Left            =   3960
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   75
         Width           =   1695
      End
      Begin Label lblinfo 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Unidades"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   810
      End
      Begin Label lblinfo 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "U&sar "
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   0
         Left            =   3480
         TabIndex        =   4
         Top             =   120
         Width           =   465
      End
   End
   Begin SSPanel panel 
      Align           =   1  'Align Top
      Height          =   450
      Index           =   0
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   3645
      Begin PictureBox EditarDis 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   330
         Left            =   1920
         Picture         =   FREADY.FRX:057C
         ScaleHeight     =   330
         ScaleWidth      =   360
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   60
         Visible         =   0   'False
         Width           =   360
      End
      Begin PictureBox AyudaDis 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   330
         Left            =   2640
         Picture         =   FREADY.FRX:06FE
         ScaleHeight     =   330
         ScaleWidth      =   360
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   60
         Width           =   360
      End
      Begin PictureBox ImprimirDis 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   330
         Left            =   2280
         Picture         =   FREADY.FRX:0880
         ScaleHeight     =   330
         ScaleWidth      =   360
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   60
         Width           =   360
      End
      Begin PictureBox BorrarDis 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   330
         Left            =   1560
         Picture         =   FREADY.FRX:0A02
         ScaleHeight     =   330
         ScaleWidth      =   360
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   60
         Visible         =   0   'False
         Width           =   360
      End
      Begin PictureBox AyudaUp 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   4320
         Picture         =   FREADY.FRX:0B84
         ScaleHeight     =   330
         ScaleWidth      =   360
         TabIndex        =   48
         Top             =   120
         Visible         =   0   'False
         Width           =   360
      End
      Begin PictureBox AyudaDown 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   4320
         Picture         =   FREADY.FRX:1B88
         ScaleHeight     =   330
         ScaleWidth      =   360
         TabIndex        =   47
         Top             =   120
         Visible         =   0   'False
         Width           =   360
      End
      Begin PictureBox VerificarDis 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   330
         Left            =   1200
         Picture         =   FREADY.FRX:2B8C
         ScaleHeight     =   330
         ScaleWidth      =   360
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   60
         Visible         =   0   'False
         Width           =   360
      End
      Begin PictureBox CrearExeDis 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   330
         Left            =   840
         Picture         =   FREADY.FRX:2D0E
         ScaleHeight     =   330
         ScaleWidth      =   360
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   60
         Visible         =   0   'False
         Width           =   360
      End
      Begin PictureBox AbrirDis 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   330
         Left            =   480
         Picture         =   FREADY.FRX:2E90
         ScaleHeight     =   330
         ScaleWidth      =   360
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   60
         Visible         =   0   'False
         Width           =   360
      End
      Begin PictureBox NuevoDis 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   330
         Left            =   120
         Picture         =   FREADY.FRX:3012
         ScaleHeight     =   330
         ScaleWidth      =   360
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   60
         Visible         =   0   'False
         Width           =   360
      End
      Begin PictureBox ImprimirDown 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   4680
         Picture         =   FREADY.FRX:3194
         ScaleHeight     =   330
         ScaleWidth      =   360
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   360
      End
      Begin PictureBox ImprimirUp 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   4680
         Picture         =   FREADY.FRX:3316
         ScaleHeight     =   330
         ScaleWidth      =   360
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   360
      End
      Begin PictureBox EditarDown 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   6840
         Picture         =   FREADY.FRX:3498
         ScaleHeight     =   330
         ScaleWidth      =   360
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   360
      End
      Begin PictureBox EditarUp 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   6840
         Picture         =   FREADY.FRX:361A
         ScaleHeight     =   330
         ScaleWidth      =   360
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   360
      End
      Begin PictureBox BorrarDown 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   6480
         Picture         =   FREADY.FRX:379C
         ScaleHeight     =   330
         ScaleWidth      =   360
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   360
      End
      Begin PictureBox BorrarUp 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   6480
         Picture         =   FREADY.FRX:391E
         ScaleHeight     =   330
         ScaleWidth      =   360
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   360
      End
      Begin PictureBox VerificarDown 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   6120
         Picture         =   FREADY.FRX:3AA0
         ScaleHeight     =   330
         ScaleWidth      =   360
         TabIndex        =   24
         Top             =   120
         Visible         =   0   'False
         Width           =   360
      End
      Begin PictureBox VerificarUp 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   6120
         Picture         =   FREADY.FRX:3C22
         ScaleHeight     =   330
         ScaleWidth      =   360
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   360
      End
      Begin PictureBox CrearExeDown 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   5280
         Picture         =   FREADY.FRX:3DA4
         ScaleHeight     =   330
         ScaleWidth      =   360
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   360
      End
      Begin PictureBox CrearExeUp 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   5280
         Picture         =   FREADY.FRX:3F26
         ScaleHeight     =   330
         ScaleWidth      =   360
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   360
      End
      Begin PictureBox AbrirDown 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   7200
         Picture         =   FREADY.FRX:40A8
         ScaleHeight     =   330
         ScaleWidth      =   360
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   360
      End
      Begin PictureBox AbrirUp 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   7200
         Picture         =   FREADY.FRX:422A
         ScaleHeight     =   330
         ScaleWidth      =   360
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   360
      End
      Begin PictureBox NuevoDown 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   5760
         Picture         =   FREADY.FRX:43AC
         ScaleHeight     =   330
         ScaleWidth      =   360
         TabIndex        =   18
         Top             =   120
         Visible         =   0   'False
         Width           =   360
      End
      Begin PictureBox NuevoUp 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   5760
         Picture         =   FREADY.FRX:452E
         ScaleHeight     =   330
         ScaleWidth      =   360
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   360
      End
      Begin PictureBox Imprimir 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   2280
         Picture         =   FREADY.FRX:46B0
         ScaleHeight     =   330
         ScaleWidth      =   360
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   60
         Width           =   360
      End
      Begin PictureBox Editar 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   1920
         Picture         =   FREADY.FRX:4832
         ScaleHeight     =   330
         ScaleWidth      =   360
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   60
         Width           =   360
      End
      Begin PictureBox Borrar 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   1560
         Picture         =   FREADY.FRX:49B4
         ScaleHeight     =   330
         ScaleWidth      =   360
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   60
         Width           =   360
      End
      Begin PictureBox Ayuda 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   2640
         Picture         =   FREADY.FRX:4B36
         ScaleHeight     =   330
         ScaleWidth      =   360
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   60
         Width           =   360
      End
      Begin PictureBox Verificar 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   1200
         Picture         =   FREADY.FRX:4CB8
         ScaleHeight     =   330
         ScaleWidth      =   360
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   60
         Width           =   360
      End
      Begin PictureBox CrearExe 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   840
         Picture         =   FREADY.FRX:4E3A
         ScaleHeight     =   330
         ScaleWidth      =   360
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   60
         Width           =   360
      End
      Begin PictureBox Abrir 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   480
         Picture         =   FREADY.FRX:4FBC
         ScaleHeight     =   330
         ScaleWidth      =   360
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   60
         Width           =   360
      End
      Begin PictureBox Nuevo 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   120
         Picture         =   FREADY.FRX:513E
         ScaleHeight     =   330
         ScaleWidth      =   360
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   60
         Width           =   360
      End
   End
   Begin Menu mArchivo 
      Caption         =   "&Archivo"
      HelpContextID   =   1000
      Begin Menu mArchivo_Opcion 
         Caption         =   "&Comprimir"
         HelpContextID   =   1010
         Index           =   0
         Shortcut        =   ^C
      End
      Begin Menu mArchivo_Opcion 
         Caption         =   "&Descomprimir"
         HelpContextID   =   1020
         Index           =   1
         Shortcut        =   ^D
      End
      Begin Menu mArchivo_Opcion 
         Caption         =   "-"
         Index           =   2
      End
      Begin Menu mArchivo_Opcion 
         Caption         =   "Conver&tir "
         HelpContextID   =   1030
         Index           =   3
         Shortcut        =   ^T
      End
      Begin Menu mArchivo_Opcion 
         Caption         =   "Crear E&XE"
         HelpContextID   =   1040
         Index           =   4
         Shortcut        =   ^X
      End
      Begin Menu mArchivo_Opcion 
         Caption         =   "&Listar"
         HelpContextID   =   1050
         Index           =   5
         Shortcut        =   ^L
      End
      Begin Menu mArchivo_Opcion 
         Caption         =   "&Mover"
         HelpContextID   =   1060
         Index           =   6
         Shortcut        =   ^M
      End
      Begin Menu mArchivo_Opcion 
         Caption         =   "R&eparar"
         HelpContextID   =   1070
         Index           =   7
         Shortcut        =   ^E
      End
      Begin Menu mArchivo_Opcion 
         Caption         =   "&Verificar"
         HelpContextID   =   1080
         Index           =   8
         Shortcut        =   ^V
      End
      Begin Menu mArchivo_Opcion 
         Caption         =   "-"
         Index           =   9
      End
      Begin Menu mArchivo_Opcion 
         Caption         =   "Bo&rrar"
         HelpContextID   =   1090
         Index           =   10
         Shortcut        =   {DEL}
      End
      Begin Menu mArchivo_Opcion 
         Caption         =   "E&ditar"
         HelpContextID   =   1100
         Index           =   11
         Shortcut        =   {F2}
      End
      Begin Menu mArchivo_Opcion 
         Caption         =   "Im&primir"
         HelpContextID   =   1110
         Index           =   12
         Shortcut        =   ^P
      End
      Begin Menu mArchivo_Opcion 
         Caption         =   "&Información"
         HelpContextID   =   1120
         Index           =   13
         Shortcut        =   ^I
      End
      Begin Menu mArchivo_Opcion 
         Caption         =   "-"
         Index           =   14
      End
      Begin Menu mArchivo_Opcion 
         Caption         =   "E&jecución directa"
         HelpContextID   =   1130
         Index           =   15
         Shortcut        =   ^J
      End
      Begin Menu mArchivo_Opcion 
         Caption         =   "Mód&ulo de instalación"
         HelpContextID   =   1140
         Index           =   16
         Shortcut        =   ^U
      End
      Begin Menu mArchivo_Opcion 
         Caption         =   "-"
         Index           =   17
      End
      Begin Menu mArchivo_Opcion 
         Caption         =   "&Salir"
         Index           =   18
      End
   End
   Begin Menu mEdicion 
      Caption         =   "&Edición"
      HelpContextID   =   2000
      Begin Menu mEdicion_Opcion 
         Caption         =   "A&gregar todos"
         HelpContextID   =   2000
         Index           =   0
      End
      Begin Menu mEdicion_Opcion 
         Caption         =   "Q&uitar todos"
         HelpContextID   =   2000
         Index           =   1
      End
      Begin Menu mEdicion_Opcion 
         Caption         =   "-"
         Index           =   2
      End
      Begin Menu mEdicion_Opcion 
         Caption         =   "Selección múltiple"
         HelpContextID   =   2010
         Index           =   3
      End
   End
   Begin Menu mOpciones 
      Caption         =   "&Opciones"
      HelpContextID   =   3000
      Begin Menu mOpciones_Opciones 
         Caption         =   "&Configuración"
         HelpContextID   =   3010
         Index           =   0
      End
      Begin Menu mOpciones_Opciones 
         Caption         =   "&Lista de directorios"
         HelpContextID   =   3020
         Index           =   1
      End
   End
   Begin Menu mUtilitarios 
      Caption         =   "&Utilitarios"
      HelpContextID   =   4000
      Begin Menu mUtilitarios_Opcion 
         Caption         =   "&Arjsort          "
         Enabled         =   0   'False
         HelpContextID   =   4010
         Index           =   0
         Shortcut        =   {F9}
      End
      Begin Menu mUtilitarios_Opcion 
         Caption         =   "&Rcvt           "
         Enabled         =   0   'False
         HelpContextID   =   1030
         Index           =   1
         Shortcut        =   {F11}
      End
      Begin Menu mUtilitarios_Opcion 
         Caption         =   "R&earj           "
         Enabled         =   0   'False
         HelpContextID   =   1030
         Index           =   2
         Shortcut        =   {F12}
      End
   End
   Begin Menu mAyuda 
      Caption         =   "A&yuda"
      Begin Menu mAyuda_Opcion 
         Caption         =   "&Indice"
         HelpContextID   =   1
         Index           =   1
         Shortcut        =   {F1}
      End
      Begin Menu mAyuda_Opcion 
         Caption         =   "&Buscar ayuda"
         Index           =   2
      End
      Begin Menu mAyuda_Opcion 
         Caption         =   "&Usar ayuda"
         Index           =   3
      End
      Begin Menu mAyuda_Opcion 
         Caption         =   "-"
         Index           =   4
      End
      Begin Menu mAyuda_Opcion 
         Caption         =   "&Cómo registrarse"
         Index           =   5
      End
      Begin Menu mAyuda_Opcion 
         Caption         =   "&Requerimientos"
         Index           =   6
      End
      Begin Menu mAyuda_Opcion 
         Caption         =   "-"
         Index           =   7
      End
      Begin Menu mAyuda_Opcion 
         Caption         =   "Acerca &de ..."
         Index           =   8
      End
   End
End

Sub A_Borrar ()
    Dim i As Single
    
    If List1.ListCount < 0 Then Exit Sub
    Ruta = Cual()
    Select Case Right(Ruta, 3)
	Case "arj": If Not VerCompresor("Arj") Then RutinaError "Arj", "Borrar": Exit Sub Else Start = DIR_DE_ARJ + " l " + Ruta + " > " + PATH_LIST: Ext = "Arj"
	Case "lzh": If Not VerCompresor("Lha") Then RutinaError "Lha", "Borrar": Exit Sub Else Start = DIR_DE_LHA + " l " + Ruta + " > " + PATH_LIST: Ext = "Lha"
	Case "pak": If Not VerCompresor("Pak") Then RutinaError "Pak", "Borrar": Exit Sub Else Start = DIR_DE_PAK + " t " + Ruta + " > " + PATH_LIST: Ext = "Pak"
	Case "zip": If Not VerCompresor("Pkzip") Then RutinaError "Pkzip", "Borrar": Exit Sub Else Start = DIR_DE_PKZIP + " -vb " + Ruta + " > " + PATH_LIST: Ext = "Pkzip"
	Case "rar": If Not VerCompresor("Rar") Then RutinaError "Rar", "Borrar": Exit Sub Else Start = DIR_DE_RAR + " l " + Ruta + " > " + PATH_LIST: Ext = "Rar"
	Case Else
	ShowErr " borrar "
	Exit Sub
    End Select
	
    Colocar Ext
    Acoplar Start, " 1 > " + PATH_ERROR, " 0 > " + PATH_ERROR
    MsgBox Start
    NumTareas = GetNumTasks()
    
    Id = Shell(PATH_BAT, 3)
    Do While GetNumTasks() <> NumTareas
	Id = DoEvents()
    Loop
    ChDir app.Path
    Open PATH_ERROR For Input As #1
	Cerr = LTrim(Input$(LOF(1), 1))
    Close #1
    Kill PATH_ERROR
    COMPRESS_ACTIVE = usar.List(usar.ListIndex)
    Select Case Cerr
	Case "1"
	    MsgBox "Error en el archivo " + LCase(Ruta), 48, "Ready"
	Case Else
	    ACTUAL = Ruta
	    fborrar.Show 1
    End Select
End Sub

Sub A_Comprimir ()
    If usar.Text = "Pkunzip" Then
	Msg = "Error de selección." + CRLF
	Msg = Msg + CRLF
	Msg = Msg + "Pkunzip es un descompresor de archivos" + CRLF
	Msg = Msg + "en formato zip." + CRLF
	Msg = Msg + "Seleccione en la lista Usar otro " + CRLF
	Msg = Msg + "compresor y vuelva a intentar." + CRLF
	MsgBox Msg, 48, "Ready"
	Exit Sub
    End If
    If List1.ListCount > 0 Then fcomprimir.Show 1
End Sub

Sub A_Convertir ()
    If List1.ListCount < 0 Then Exit Sub
    Ruta = Cual()
    Select Case Right(Ruta, 3)
	Case "arj": If Not VerCompresor("Arj") Then RutinaError "Arj", "Convertir": Exit Sub
	Case "lzh": If Not VerCompresor("Lha") Then RutinaError "Lha", "Convertir": Exit Sub
	Case "pak": If Not VerCompresor("Pak") Then RutinaError "Pak", "Convertir": Exit Sub
	Case "rar": If Not VerCompresor("Rar") Then RutinaError "Rar", "Convertir": Exit Sub
	Case "zip": : If Not VerCompresor("Pkunzip") Then RutinaError "Pkunzip", "Convertir": Exit Sub
	Case Else
	ShowErr " convertir "
	Exit Sub
    End Select
    fconvertir.Show 1
End Sub

Sub A_CrearExe ()
    If List1.ListCount < 0 Then Exit Sub
    Ruta = Cual()
    Select Case Right(Ruta, 3)
	Case "arj": If VerCompresor("Arj") Then Start = DIR_DE_ARJ + " y - je1 " + Ruta Else RutinaError "Arj", "Crear exe": Exit Sub
	Case "lzh": If VerCompresor("Lha") Then Start = DIR_DE_LHA + " s " + Ruta Else RutinaError "Lha", "Crear exe": Exit Sub
	Case "pak": If VerCompresor("Pak") Then Start = DIR_DE_PAK + " c /exe " + Ruta Else RutinaError "Pak", "Crear exe": Exit Sub
	Case "rar": If VerCompresor("Rar") Then Start = DIR_DE_RAR + " s " + Ruta Else RutinaError "Rar", "Crear exe": Exit Sub
	Case "zip": : If DIR_DE_ZIP2EXE <> "" Then Start = DIR_DE_ZIP2EXE + " " + Ruta Else RutinaError "Zip2exe", "Crear exe": Exit Sub
	Case Else
	ShowErr " crear exe "
	Exit Sub
    End Select
    MsgError = "Error al crear archivo EXE > " + PATH_ERROR
    MsgOk = "Exito al crear archivo EXE >" + PATH_ERROR
    Go Start, MsgError, MsgOk
    Start = ""
End Sub

Sub A_Descomprimir ()
    Dim i As Single
    Dim Archivo As String
    If List1.ListCount < 0 Then Exit Sub
    Ruta = Cual()
    Ext = LCase(Right(Ruta, 3))
    Select Case Ext
	Case "arj": If Not VerCompresor("Arj") Then RutinaError "Arj", "Descomprimir": Exit Sub:  Else Start = DIR_DE_ARJ + " l " + Ruta + " > " + PATH_LIST: Ext = "Arj"
	Case "lzh": If Not VerCompresor("Lha") Then RutinaError "Lha", "Descomprimir": Exit Sub:  Else Start = DIR_DE_LHA + " l " + Ruta + " > " + PATH_LIST: Ext = "Lha"
	Case "pak": If Not VerCompresor("Pak") Then RutinaError "Pak", "Descomprimir": Exit Sub:  Else Start = DIR_DE_PAK + " t " + Ruta + " > " + PATH_LIST: Ext = "Pak"
	Case "zip": If Not VerCompresor("Pkunzip") Then RutinaError "Pkunzip", "Descomprimir": Exit Sub:  Else Start = DIR_DE_PKUNZIP + " " + Ruta + " -@" + PATH_LIST: Ext = "Pkunzip"
	Case "rar": If Not VerCompresor("Rar") Then RutinaError "Rar", "Descomprimir": Exit Sub:  Else Start = DIR_DE_RAR + " l " + Ruta + " > " + PATH_LIST: Ext = "Rar"
	Case Else
	ShowErr " descomprimir "
	Exit Sub
    End Select
    
    Colocar Ext
    Acoplar Start, " 1 > " + PATH_ERROR, " 0 > " + PATH_ERROR
    NumTareas = GetNumTasks()
    Id = Shell(PATH_BAT, 1)
    Do While GetNumTasks() <> NumTareas
	Id = DoEvents()
    Loop
    ChDir app.Path
    Open PATH_ERROR For Input As #1
	Cerr = LTrim(Input$(LOF(1), 1))
    Close #1
    Kill PATH_ERROR
    Kill PATH_BAT
    COMPRESS_ACTIVE = usar.Text
    Select Case Cerr
	Case "1"
	    MsgBox "Error en el archivo " + LCase(Ruta), 48, "Ready"
	Case Else
	    FileJob = Ruta
	    fdescomprimir.Show 1
    End Select
End Sub

Sub A_Editar ()
    If List1.ListCount > 0 Then PATH_EDITAR = Cual(): feditar.Show 1
End Sub

Sub A_EjecucionDirecta ()
    fejedir.Show 1
End Sub

Sub A_imprimir ()
    Dim i As Single
    If List1.ListCount < 0 Then Exit Sub
    Ruta = Cual()
    Ext = LCase(Right(Ruta, 3))
    Select Case Ext
	Case "arj": If Not VerCompresor("Arj") Then RutinaError "Arj", "Imprimir": Exit Sub Else Start = DIR_DE_ARJ + " l " + Ruta + " > " + PATH_LIST: Ext = "Arj"
	Case "lzh": If Not VerCompresor("Lha") Then RutinaError "Lha", "Imprimir": Exit Sub Else Start = DIR_DE_LHA + " l " + Ruta + " > " + PATH_LIST: Ext = "Lha"
	Case "pak": If Not VerCompresor("Pak") Then RutinaError "Pak", "Imprimir": Exit Sub Else Start = DIR_DE_PAK + " t " + Ruta + " > " + PATH_LIST: Ext = "Pak"
	Case "zip": If Not VerCompresor("Pkunzip") Then RutinaError "Pkunzip", "Listar": Exit Sub Else Start = DIR_DE_PKUNZIP + " " + Ruta + " -@" + PATH_LIST: Ext = "Pkunzip"
	Case "rar": If Not VerCompresor("Rar") Then RutinaError "Rar", "Imprimir": Exit Sub Else Start = DIR_DE_RAR + " l " + Ruta + " > " + PATH_LIST: Ext = "Rar"
    Case Else
	ShowErr " Imprimir "
	Exit Sub
    End Select
    
    Colocar Ext
    
    Acoplar Start, " 1 > " + PATH_ERROR, " 0 > " + PATH_ERROR
    NumTareas = GetNumTasks()
    Id = Shell(PATH_BAT, 1)
    Do While GetNumTasks() <> NumTareas
	Id = DoEvents()
    Loop
    ChDir app.Path
    Open PATH_ERROR For Input As #1
	Cerr = LTrim(Input$(LOF(1), 1))
    Close #1
    Kill PATH_ERROR
    COMPRESS_ACTIVE = usar.List(usar.ListIndex)
    Select Case Cerr
	Case "1"
	    MsgBox "Error en el archivo " + LCase(Ruta), 48, "Ready"
	Case Else
	    FileJob = Ruta: fImprimir.Show 1
    End Select
End Sub

Sub A_Información ()
    If List1.ListCount > 0 Then finfor.Show 1
End Sub

Sub A_Listar ()
    Dim i As Single
    If List1.ListCount < 0 Then Exit Sub
    Ruta = Cual()
    Ext = LCase(Right(Ruta, 3))
    Select Case Ext
	Case "arj": If Not VerCompresor("Arj") Then RutinaError "Arj", "Listar": Exit Sub Else Start = DIR_DE_ARJ + " l " + Ruta + " > " + PATH_LIST: Ext = "Arj"
	Case "lzh": If Not VerCompresor("Lha") Then RutinaError "Lha", "Listar": Exit Sub Else Start = DIR_DE_LHA + " l " + Ruta + " > " + PATH_LIST: Ext = "Lha"
	Case "pak": If Not VerCompresor("Pak") Then RutinaError "Pak", "Listar": Exit Sub Else Start = DIR_DE_PAK + " l " + Ruta + " > " + PATH_LIST: Ext = "Pak"
	Case "zip": If Not VerCompresor("Pkzip") Then RutinaError "Pkzip", "Listar": Exit Sub Else Start = DIR_DE_PKZIP + " -vb " + Ruta + " > " + PATH_LIST: Ext = "Pkzip"
	Case "rar": If Not VerCompresor("Rar") Then RutinaError "Rar", "Listar": Exit Sub Else Start = DIR_DE_RAR + " l " + Ruta + " > " + PATH_LIST: Ext = "Rar"
    Case Else
	ShowErr " listar "
	Exit Sub
    End Select
    
    Colocar Ext
    
    Acoplar Start, " 1 > " + PATH_ERROR, " 0 > " + PATH_ERROR
    NumTareas = GetNumTasks()
    
    Id = Shell(PATH_BAT, 1)
    Do While GetNumTasks() <> NumTareas
	Id = DoEvents()
    Loop
    ChDir app.Path
    Open PATH_ERROR For Input As #1
	Cerr = LTrim(Input$(LOF(1), 1))
    Close #1
    Kill PATH_ERROR
    COMPRESS_ACTIVE = usar.List(usar.ListIndex)
    Select Case Cerr
	Case "1"
	    MsgBox "Error en el archivo " + LCase(Ruta), 48, "Ready"
	Case Else
	    FileJob = Ruta: flistar.Show 1
    End Select
End Sub

Sub A_Modulo ()
    fmodulo.Show 1
End Sub

Sub A_Reparar ()
    If List1.ListCount < 0 Then Exit Sub
    Ruta = Cual()
    Select Case Right(Ruta, 3)
	Case "arj", "lzh", "pak"
	Msg = "Extensión no válida" + CRLF
	Msg = Msg + CRLF
	Msg = Msg + "Para reparar archivos la extensión" + CRLF
	Msg = Msg + "debe ser rar o zip." + CRLF
	MsgBox Msg, 48, "Ready"
	Exit Sub
	Case "zip": If DIR_DE_PKZIPFIX <> "" Then Start = DIR_DE_PKZIPFIX + Ruta Else RutinaError "Pkzipfix", "Reparar": Exit Sub
	Case "rar": If VerCompresor("Rar") Then Start = DIR_DE_RAR + " r " + Ruta Else RutinaError "Rar", "Reparar": Exit Sub
    End Select
    MsgError = "Error al reparar archivo > " + PATH_ERROR
    MsgOk = "Exito al reparar archivo >" + PATH_ERROR
    Go Start, MsgError, MsgOk
    Start = ""
End Sub

Sub A_Verificar ()
    If List1.ListCount = -1 Then Exit Sub
    Ruta = Cual()
    Select Case Right(Ruta, 3)
	Case "arj": If VerCompresor("Arj") Then Start = DIR_DE_ARJ + " t -jr " + Ruta: Ext = "Arj" Else RutinaError "Arj", "Verificar": Exit Sub
	Case "lzh": If VerCompresor("Lha") Then Start = DIR_DE_LHA + " t " + Ruta: Ext = "Lha" Else RutinaError "Lha", "Verificar": Exit Sub
	Case "pak": If VerCompresor("Pak") Then Start = DIR_DE_PAK + " t " + Ruta: Ext = "Pak" Else RutinaError "Pak", "Verificar": Exit Sub
	Case "rar": If VerCompresor("Rar") Then Start = DIR_DE_RAR + " t " + Ruta: Ext = "Rar" Else RutinaError "Rar", "Verificar": Exit Sub
	Case "zip": : If VerCompresor("Pkunzip") Then Start = DIR_DE_PKUNZIP + " -t " + Ruta: Ext = "Pkunzip" Else RutinaError "Pkunzip", "Verificar": Exit Sub
	Case Else
	ShowErr " verificar "
	Exit Sub
    End Select
    Colocar Ext
    MsgError = "Error al verificar archivo > " + PATH_ERROR
    MsgOk = "Exito al verificar archivo >" + PATH_ERROR
    Go Start, MsgError, MsgOk
    Start = ""
End Sub

Sub Abrir_Click ()
    Abrir.Refresh
    A_Descomprimir
End Sub

Sub Abrir_MouseDown (Button As Integer, Shift As Integer, X As Single, Y As Single)
    Abrir.Picture = AbrirDown.Picture
End Sub

Sub Abrir_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
    lbllisto.Caption = "Descomprime archivos hacia la unidad especificada"
    Select Case Button
    Case 1
	If X <= 0 Or X > Abrir.Width Or Y < 0 Or Y > Abrir.Height Then
	    Abrir.Picture = AbrirUp.Picture
	Else
	    Abrir.Picture = AbrirDown.Picture
	End If
    End Select
End Sub

Sub Abrir_MouseUp (Button As Integer, Shift As Integer, X As Single, Y As Single)
    Abrir.Picture = AbrirUp.Picture
End Sub

Sub Ayuda_Click ()
    On Error Resume Next
    Ayuda.Refresh
    Ret = WinHelp(hWnd, PATH_HELP, HELP_CONTENTS, 1&)
End Sub

Sub Ayuda_MouseDown (Button As Integer, Shift As Integer, X As Single, Y As Single)
    Ayuda.Picture = AyudaDown.Picture
End Sub

Sub Ayuda_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
    lbllisto.Caption = "Ayuda de WinReady"
    Select Case Button
    Case 1
	If X <= 0 Or X > Ayuda.Width Or Y < 0 Or Y > Ayuda.Height Then
	    Ayuda.Picture = AyudaUp.Picture
	Else
	    Ayuda.Picture = AyudaDown.Picture
	End If
    End Select
End Sub

Sub Ayuda_MouseUp (Button As Integer, Shift As Integer, X As Single, Y As Single)
    Ayuda.Picture = AyudaUp.Picture
End Sub

Sub Borrar_Click ()
    Borrar.Refresh
    A_Borrar
End Sub

Sub Borrar_MouseDown (Button As Integer, Shift As Integer, X As Single, Y As Single)
    Borrar.Picture = BorrarDown.Picture
End Sub

Sub Borrar_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
    lbllisto.Caption = "Borra archivos del archivo comprimido"
    Select Case Button
    Case 1
	If X <= 0 Or X > Borrar.Width Or Y < 0 Or Y > Borrar.Height Then
	    Borrar.Picture = BorrarUp.Picture
	Else
	    Borrar.Picture = BorrarDown.Picture
	End If
    End Select
End Sub

Sub Borrar_MouseUp (Button As Integer, Shift As Integer, X As Single, Y As Single)
    Borrar.Picture = BorrarUp.Picture
End Sub

Sub Colocar (Who As String)
    Dim i As Integer
    For i = 0 To usar.ListCount - 1
	If usar.List(i) = Who Then usar.ListIndex = i: Exit For
    Next i
End Sub

Sub CrearExe_Click ()
    CrearExe.Refresh
    A_CrearExe
End Sub

Sub CrearExe_MouseDown (Button As Integer, Shift As Integer, X As Single, Y As Single)
    CrearExe.Picture = CrearExeDown.Picture
End Sub

Sub CrearExe_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
    lbllisto.Caption = "Crea un archivo ejecutable a partir de uno comprimido"
    Select Case Button
    Case 1
	If X <= 0 Or X > CrearExe.Width Or Y < 0 Or Y > CrearExe.Height Then
	    CrearExe.Picture = CrearExeUp.Picture
	Else
	    CrearExe.Picture = CrearExeDown.Picture
	End If
    End Select
End Sub

Sub CrearExe_MouseUp (Button As Integer, Shift As Integer, X As Single, Y As Single)
    CrearExe.Picture = CrearExeUp.Picture
End Sub

Function Cual () As String
    For i = 0 To List1.ListCount - 1
	If List1.Selected(i) Then
	    Ruta = Path(Dir1) + List1.List(i)
	    ACTUAL = Ruta
	    Cual = Ruta: Exit Function
	End If
    Next i
End Function

Sub Dir1_Change ()
    Dim i As Single
    Dim nBytes As Long
    Dim cPath As String
    file1.Path = Dir1.Path
    cPath = Path(Dir1)
    ChDir Dir1.Path
    List1.Clear
    screen.MousePointer = 11
    lbllisto.Caption = "Espere un momento .... "
    nTotal = file1.ListCount - 1
    For i = 0 To nTotal
	List1.AddItem file1.List(i)
	nBytes = nBytes + FileLen(cPath + file1.List(i))
    Next i
    screen.MousePointer = 0
    lbllisto.Caption = "Listo"
    lbltot.Caption = List1.ListCount
    lblbytes.Caption = nBytes
    List1.ListIndex = -1
    If List1.ListCount > 0 Then List1.Selected(0) = -1
End Sub

Sub Dir1_KeyPress (KeyAscii As Integer)
    If KeyAscii = 13 Then Dir1.Path = Dir1.List(Dir1.ListIndex): KeyAscii = 0
End Sub

Sub E_AgregarTodos ()
    Dim i As Single
    screen.MousePointer = 11
    If MULTIPLE = 0 Then
	For i = 0 To List1.ListCount - 1: List1.Selected(i) = -1: Next i
	List1.ListIndex = -1
    Else
	For i = 0 To list3.ListCount - 1: list3.Selected(i) = -1: Next i
	list3.ListIndex = -1
    End If
    screen.MousePointer = 0
End Sub

Sub E_QuitarTodos ()
    Dim i As Single
    screen.MousePointer = 11
    If MULTIPLE = 0 Then
	For i = 0 To List1.ListCount - 1: List1.Selected(i) = 0: Next i
	List1.ListIndex = -1
    Else
	For i = 0 To list3.ListCount - 1: list3.Selected(i) = 0: Next i
	list3.ListIndex = -1
    End If
    screen.MousePointer = 0
End Sub

Sub E_Seleccion ()
    fselmul.Show 1
End Sub

Sub Editar_Click ()
    Editar.Refresh
    A_Editar
End Sub

Sub Editar_MouseDown (Button As Integer, Shift As Integer, X As Single, Y As Single)
    Editar.Picture = EditarDown.Picture
End Sub

Sub Editar_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
    lbllisto.Caption = "Permite editar un archivo"
    Select Case Button
    Case 1
	If X <= 0 Or X > Editar.Width Or Y < 0 Or Y > Editar.Height Then
	    Editar.Picture = EditarUp.Picture
	Else
	    Editar.Picture = EditarDown.Picture
	End If
    End Select
End Sub

Sub Editar_MouseUp (Button As Integer, Shift As Integer, X As Single, Y As Single)
    Editar.Picture = EditarUp.Picture
End Sub

Sub EliminarArchivosDeTrabajo ()
    On Error Resume Next
    Kill PATH_EXC
    Kill PATH_ERROR
    Kill PATH_LIST
    Kill PATH_FILES
    Kill PATH_BAT
    Kill PATH_DIRS
End Sub

Sub Form_Activate ()
    If MULTIPLE = -1 Then list3.ListIndex = -1: Exit Sub Else file1.Refresh : Dir1_Change
End Sub

Sub Form_Resize ()
    fready.WindowState = 2
End Sub

Sub Form_Unload (Cancel As Integer)
    GrabarInformacionAIniFile
    EliminarArchivosDeTrabajo
    End
End Sub

Sub Imprimir_Click ()
    Imprimir.Refresh
    A_imprimir
End Sub

Sub Imprimir_MouseDown (Button As Integer, Shift As Integer, X As Single, Y As Single)
    Imprimir.Picture = ImprimirDown.Picture
End Sub

Sub Imprimir_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
    lbllisto.Caption = "Imprime archivos que están en el archivo comprimido"
    Select Case Button
    Case 1
	If X <= 0 Or X > Imprimir.Width Or Y < 0 Or Y > Imprimir.Height Then
	    Imprimir.Picture = ImprimirUp.Picture
	Else
	    Imprimir.Picture = ImprimirDown.Picture
	End If
    End Select
End Sub

Sub Imprimir_MouseUp (Button As Integer, Shift As Integer, X As Single, Y As Single)
    Imprimir.Picture = ImprimirUp.Picture
End Sub

Sub List1_DblClick ()
    A_Descomprimir
End Sub

Sub mArchivo_opcion_Click (Index As Integer)
End Sub

Sub mUtilitarios_Opcion_Click (Index As Integer)
    Select Case Index
	Case UTILITARIOS_ARJSORT: U_Arjsort
	Case UTILITARIOS_RCVT: U_Rcvt
	Case UTILITARIOS_REARJ: U_Rearj
    End Select
End Sub

Sub Nuevo_Click ()
    Nuevo.Refresh
    mArchivo_opcion_Click 0
End Sub

Sub Nuevo_MouseDown (Button As Integer, Shift As Integer, X As Single, Y As Single)
    Nuevo.Picture = NuevoDown.Picture
End Sub

Sub Nuevo_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
    lbllisto.Caption = "Comprime archivos hacia la unidad especificada"
    Select Case Button
    Case 1
	If X <= 0 Or X > Nuevo.Width Or Y < 0 Or Y > Nuevo.Height Then
	    Nuevo.Picture = NuevoUp.Picture
	Else
	    Nuevo.Picture = NuevoDown.Picture
	End If
    End Select
End Sub

Sub Nuevo_MouseUp (Button As Integer, Shift As Integer, X As Single, Y As Single)
    Nuevo.Picture = NuevoUp.Picture
End Sub

Sub O_ActualizarPosibles ()
    If mOpciones_Opciones(OPCIONES_LISTA_DIR).Checked = 0 Then
	mOpciones_Opciones(OPCIONES_LISTA_DIR).Checked = -1
	flistadir.Show 1
    Else
	mOpciones_Opciones(OPCIONES_LISTA_DIR).Checked = 0
    End If
    LISTA_DIR = Str(mOpciones_Opciones(OPCIONES_LISTA_DIR).Checked)
End Sub

Sub O_Configuracion ()
    fsetup.Show 1
End Sub

Sub Panel_MouseMove (Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lbllisto.Caption = "Listo"
End Sub

Sub RutinaError (MsgShow As String, Accion As String)
    Msg = MsgShow + " no esta instalado." + CRLF
    Msg = Msg + CRLF
    Msg = Msg + "En el menú opciones seleccione" + CRLF
    Msg = Msg + """Configuración"",después escriba el path" + CRLF
    Msg = Msg + "en el cuadro de texto " + MsgShow + " ó " + CRLF
    Msg = Msg + "haga click en el botón ""Buscar"" para" + CRLF
    Msg = Msg + "encontrar en su computador el compresor " + MsgShow + CRLF
    Msg = Msg + "Esto es necesario para realizar con" + CRLF
    Msg = Msg + "éxito la acción de " + Accion + "." + CRLF
    MsgBox Msg, 48, "Win Ready"
End Sub

Sub Salir ()
    On Error Resume Next
    GrabarInformacionAIniFile
    Ret = WinHelp(hWnd, PATH_HELP, HELP_QUIT, 0&)
    EliminarArchivosDeTrabajo
    End
End Sub

Sub ShowErr (MsgShow As String)
    Msg = "Extensión no válida." + CRLF
    Msg = Msg + CRLF
    Msg = Msg + "Para" + MsgShow + "archivos la extensión" + CRLF
    Msg = Msg + "debe ser arj, lzh, pak, rar o zip." + CRLF
    MsgBox Msg, 48, "Win Ready"
End Sub

Sub SugeHelp (Msg As String)
    lblsuge.Caption = Msg
End Sub

Sub U_Arjsort ()
    Dim Copiar, Archivo, Destino As String
    If List1.ListIndex = -1 Then Exit Sub
    Ruta = Cual()
    Select Case Right(Ruta, 3)
	Case "arj": If VerCompresor("Arj") Then farjsort.Show 1 Else RutinaError "Arj", "Ordenar": Exit Sub
	Case Else
	    Msg = "Extensión no válida" + CRLF
	    Msg = Msg + CRLF
	    Msg = Msg + "Para sortear archivos la extensión" + CRLF
	    Msg = Msg + "debe ser arj." + CRLF
	    MsgBox Msg, 48, "Win Ready"
	    Exit Sub
    End Select
End Sub

Sub U_Rcvt ()
    Dim Archivo, Destino, JobPath As String
    If List1.ListIndex = -1 Then Exit Sub
    Ruta = Cual()
    MsgError = "Error en la conversión > " + PATH_ERROR
    MsgOk = "Exito en la conversión >" + PATH_ERROR
    Select Case Right(Ruta, 3)
	Case "arj": JobPath = DIR_DE_ARJ: If Not VerCompresor("Arj") Then RutinaError "Arj", "Convertir": Exit Sub
	Case "lzh": JobPath = DIR_DE_LHA: If Not VerCompresor("Lha") Then RutinaError "Lha", "Convertir": Exit Sub
	Case "rar": JobPath = DIR_DE_RAR: If Not VerCompresor("Rar") Then RutinaError "Rar", "Convertir": Exit Sub
	Case "zip": JobPath = DIR_DE_PKUNZIP: If Not VerCompresor("Pkunzip") Then RutinaError "Pkunzip", "Convertir": Exit Sub
	Case Else
	    Msg = "Extensión no válida" + CRLF
	    Msg = Msg + CRLF
	    Msg = Msg + "Para convertir archivos la extensión" + CRLF
	    Msg = Msg + "debe ser arj,lzh,rar o zip." + CRLF
	    MsgBox Msg, 48, "Win Ready"
	    Exit Sub
    End Select
    Origen = QuitSlash(DIR_DE_RCVT)
    Archivo = GetFileName(JobPath)
    Destino = Origen + Archivo
    
    Start = DIR_DE_RCVT + " " + Ruta
    Open Destino For Binary As #1
    If LOF(1) = 0 Then
	FileCopy JobPath, Destino
	Go Start, MsgError, MsgOk
	Close #1
	Kill Destino
    Else
	Close #1
	Go Start, MsgError, MsgOk
    End If
End Sub

Sub U_Rearj ()
    Dim Archivo, Destino, JobPath As String
    If List1.ListIndex = -1 Then Exit Sub
    Ruta = Cual()
    MsgError = "Error en la conversión > " + PATH_ERROR
    MsgOk = "Exito en la conversión >" + PATH_ERROR
    Select Case Right(Ruta, 3)
	Case "arj": JobPath = DIR_DE_ARJ: If Not VerCompresor("Arj") Then RutinaError "Arj", "Convertir": Exit Sub
	Case "lzh": JobPath = DIR_DE_LHA: If Not VerCompresor("Lha") Then RutinaError "Lha", "Convertir": Exit Sub
	Case "pak": JobPath = DIR_DE_PAK: If Not VerCompresor("Pak") Then RutinaError "Pak", "Convertir": Exit Sub
	Case "rar": JobPath = DIR_DE_RAR: If Not VerCompresor("Rar") Then RutinaError "Rar", "Convertir": Exit Sub
	Case "zip": JobPath = DIR_DE_PKUNZIP: If Not VerCompresor("Pkunzip") Then RutinaError "Pkunzip", "Convertir": Exit Sub
	Case Else
	    Msg = "Extensión no válida" + CRLF
	    Msg = Msg + CRLF
	    Msg = Msg + "Para convertir archivos la extensión" + CRLF
	    Msg = Msg + "debe ser arj,lzh,pak,rar o zip." + CRLF
	    MsgBox Msg, 48, "Win Ready"
	    Exit Sub
    End Select
    Origen = QuitSlash(DIR_DE_REARJ)
    Archivo = GetFileName(JobPath)
    Destino = Origen + Archivo
    Start = DIR_DE_REARJ + " " + Ruta
    Open Destino For Binary As #1
    If LOF(1) = 0 Then
	Close #1
	FileCopy JobPath, Destino
	Go Start, MsgError, MsgOk
	Kill Destino
    Else
	Close #1
	Go Start, MsgError, MsgOk
    End If
End Sub

Sub Usar_Change ()
    
End Sub

Function VerCompresor (Cual As String)
    Dim i As Integer

    For i = 0 To usar.ListCount - 1
	If usar.List(i) = Cual Then
	    VerCompresor = True
	    COMPRESS_ACTIVE = usar.List(usar.ListIndex)
	    Exit Function
	End If
    Next i
    VerCompresor = False
End Function

Sub Verificar_Click ()
    Verificar.Refresh
    A_Verificar
End Sub

Sub Verificar_MouseDown (Button As Integer, Shift As Integer, X As Single, Y As Single)
    Verificar.Picture = VerificarDown.Picture
End Sub

Sub Verificar_MouseMove (Button As Integer, Shift As Integer, X As Single, Y As Single)
    lbllisto.Caption = "Verifica la integridad del archivo comprimido"
    Select Case Button
    Case 1
	If X <= 0 Or X > Verificar.Width Or Y < 0 Or Y > Verificar.Height Then
	    Verificar.Picture = VerificarUp.Picture
	Else
	    Verificar.Picture = VerificarDown.Picture
	End If
    End Select
End Sub

Sub Verificar_MouseUp (Button As Integer, Shift As Integer, X As Single, Y As Single)
    Verificar.Picture = VerificarUp.Picture
End Sub

