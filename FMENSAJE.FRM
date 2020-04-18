VERSION 5.00
Begin VB.Form fmensaje 
   Appearance      =   0  'Flat
   BorderStyle     =   0  'None
   ClientHeight    =   615
   ClientLeft      =   2340
   ClientTop       =   3030
   ClientWidth     =   5235
   ControlBox      =   0   'False
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
   ScaleHeight     =   615
   ScaleWidth      =   5235
   Begin VB.PictureBox Panel3D1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   5205
      TabIndex        =   0
      Top             =   0
      Width           =   5235
   End
End
Attribute VB_Name = "fmensaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    CentrarForma fmensaje
End Sub

