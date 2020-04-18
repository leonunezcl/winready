VERSION 2.00
Begin Form fexpandir 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Expandir "
   ClientHeight    =   3015
   ClientLeft      =   1410
   ClientTop       =   2145
   ClientWidth     =   6720
   Height          =   3420
   Left            =   1350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   6720
   Top             =   1800
   Width           =   6840
   Begin ComboBox Posibles 
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   300
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   2580
      Width           =   4815
   End
   Begin SSCheck Posdes 
      Caption         =   "&Posibles destinos"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2280
      Value           =   -1  'True
      Width           =   1575
   End
   Begin SSCommand btn 
      Caption         =   "A&yuda"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   2
      Left            =   5160
      Outline         =   0   'False
      TabIndex        =   7
      Top             =   1560
      Width           =   1455
   End
   Begin SSCommand btn 
      Caption         =   "&Cancelar"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   1
      Left            =   5160
      Outline         =   0   'False
      TabIndex        =   6
      Top             =   1080
      Width           =   1455
   End
   Begin SSCommand btn 
      Caption         =   "&Aceptar"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   0
      Left            =   5160
      Outline         =   0   'False
      TabIndex        =   5
      Top             =   600
      Width           =   1455
   End
   Begin DriveListBox Drive1 
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   315
      Left            =   2880
      TabIndex        =   4
      Top             =   1870
      Width           =   2175
   End
   Begin DirListBox Dir1 
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   1155
      Left            =   2880
      TabIndex        =   3
      Top             =   600
      Width           =   2175
   End
   Begin MListBox List1 
      AddItemHeight   =   0
      Alignment       =   1  'Left
      AllowFocusRect  =   -1  'True
      AutoCheck       =   -1  'True
      BorderStyle     =   0  'Normal
      CheckColor      =   &H00000000&
      CheckStyle      =   0  'Cross Style
      DefPicture      =   FEXPANDI.FRX:0000
      DrawRegions     =   1
      EnableVirtualMsgs=   0   'False
      ExtendedSelect  =   0   'False
      FallColor       =   &H00808080&
      FindDirection   =   0  'Forward
      FindResult      =   0  'Update ListIndex
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      GridColor       =   &H00000000&
      GridStyle       =   0  'Solid
      Height          =   1590
      HiliteBackColor =   &H00000000&
      HiliteForeColor =   &H00000000&
      HorizontalGrids =   0   'False
      ImageRegion     =   1
      ImageType       =   3  'Aligned Bitmap
      ItemHeight      =   195
      ItemWidth       =   1560
      Left            =   240
      ListBoxStyle    =   0  'Fixed
      MaskingColor    =   &H00FFFFFF&
      MultiColumn     =   0   'False
      MultiSelect     =   0   'False
      NoIntegralHeight=   0   'False
      OwnerCompare    =   0   'False
      OwnerDraw       =   0   'False
      RiseColor       =   &H00FFFFFF&
      SearchCompare   =   1  'Case Insensitive
      SelectMode      =   0  'Normal
      SortColumn      =   0
      Sorted          =   0   'False
      SortOrder       =   0  'Ascending
      StringCompare   =   0  'Case Sensitive
      TabIndex        =   1
      Top             =   600
      Version         =   "04.56"
      VerticalGrids   =   0   'False
      VirtualMsgZone  =   0
      Width           =   2535
   End
   Begin Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Expandir &hacia el Directorio/unidad"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   390
      Left            =   2880
      TabIndex        =   2
      Top             =   120
      Width           =   1335
      WordWrap        =   -1  'True
   End
   Begin Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "A&rchivos seleccionados a expandir"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   390
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1935
      WordWrap        =   -1  'True
   End
End
Dim N, SW
Dim ExtOri(0 To 13), ExtDes(0 To 13) As String

Sub btn_Click (Index As Integer)
    Select Case Index
        Case 0: Expandir
        Case 1: Unload fexparc: Unload fexpandir
        Case 2
    End Select
End Sub

Sub btn_KeyPress (Index As Integer, keyascii As Integer)
    If keyascii = 13 Then
        keyascii = 0
        Select Case Index
            Case 0: Expandir
            Case 1: Unload fexparc: Unload fexpandir
            Case 2
        End Select
    End Select
End Sub

Sub CargarTablas ()
    ExtOri(0) = "cp_"
    ExtOri(1) = "ex_"
    ExtOri(2) = "sy_"
    ExtOri(3) = "co_"
    ExtOri(4) = "tx_"
    ExtOri(5) = "hl_"
    ExtOri(6) = "bi_"
    ExtOri(7) = "in_"
    ExtOri(8) = "pr_"
    ExtOri(9) = "38_"
    ExtOri(10) = "dl_"
    ExtOri(11) = "gr_"
    ExtOri(12) = "ov_"
    ExtOri(13) = "ls_"

    ExtDes(0) = "cpi"
    ExtDes(1) = "exe"
    ExtDes(2) = "sys"
    ExtDes(3) = "com"
    ExtDes(4) = "txt"
    ExtDes(5) = "hlp"
    ExtDes(6) = "bin"
    ExtDes(7) = "inf"
    ExtDes(8) = "pro"
    ExtDes(9) = "386"
    ExtDes(10) = "dll"
    ExtDes(11) = "grp"
    ExtDes(12) = "ovl"
    ExtDes(13) = "lst"

End Sub

Sub Drive1_Change ()
    On Error GoTo ErrorUnidad
    Dir1.Path = drive1.Drive
    Exit Sub
ErrorUnidad:
    Msg = "La unidad " + drive1.Drive + " no está preparada." + CRLF
    Msg = Msg + CRLF
    Msg = Msg + "Si expande hacia unidades de disquete" + CRLF
    Msg = Msg + "inserte un disquete en ella, asegurese" + CRLF
    Msg = Msg + "que el disquete este formateado y que no" + CRLF
    Msg = Msg + "este protegido contra escritura." + CRLF
    MsgBox Msg, 48, "WinReady"
    drive1.Drive = app.Path
    Exit Sub
End Sub

Sub Expandir ()
    Dim i As Single
    CargarTablas


    For i = 0 To list1.listcount - 1
        If list1.selected(i) Then
            For v = 0 To 14
                Select Case Right(list1.list(i), 3)
                    Case ExtOri(v)
                End Select
            Next v
        End If
    Next i
    Do
        If Not CopyFile(SourcePath$, destPath$, Origen(Incre), Destino(Incre)) Then ErrorSetup Else Progreso Incre
        Incre = Incre + 1
        DoEvents
    Loop Until Incre > 9

End Sub

Sub Form_Load ()
    CentrarForma fexpandir
    Load fexparc
    Select Case winjob.list1.ListIndex
        Case -1: AddXlist winjob.list1, fexpandir.list1, 0: N = True
        Case Else: AddXlist winjob.list1, fexpandir.list1, 1: N = False
    End Select
    
    PosDes.Enabled = IIf(Val(LISTA_DIR) = -1, -1, 0)
    PosDes.Value = IIf(Val(LISTA_DIR) = -1, -1, 0)
    Posibles.Enabled = IIf(Val(LISTA_DIR) = -1, -1, 0)
    If Val(LISTA_DIR) = -1 Then LLenarDirs fexpandir.Posibles
    If Val(LISTA_DIR) = -1 Then Posibles.ListIndex = 0
End Sub

Sub Form_Paint ()
    Call FormIn3d(Me, 1, False)
End Sub

Sub Posibles_Click ()
    Dir1.Path = Posibles.List(Posibles.ListIndex)
End Sub

Sub Posibles_Change ()
    Dir1.Path = Posibles.list(Posibles.ListIndex)
End Sub

Sub Progreso ()

End Sub

