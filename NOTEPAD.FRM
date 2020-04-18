VERSION 2.00
Begin Form frmNotePad 
   Caption         =   "PC Plus-[Sin titulo]"
   ClientHeight    =   3975
   ClientLeft      =   1155
   ClientTop       =   2460
   ClientWidth     =   5610
   Height          =   4665
   Icon            =   NOTEPAD.FRX:0000
   Left            =   1095
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3975
   ScaleWidth      =   5610
   Top             =   1830
   Visible         =   0   'False
   Width           =   5730
   Begin TextBox Text1 
      Height          =   3855
      HideSelection   =   0   'False
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   5655
   End
   Begin Menu mnuFile 
      Caption         =   "&Archivo"
      Begin Menu mnuFNew 
         Caption         =   "&Nuevo"
      End
      Begin Menu mnuFOpen 
         Caption         =   "A&brir"
      End
      Begin Menu mnuFClose 
         Caption         =   "&Cerrar"
      End
      Begin Menu mnuFSave 
         Caption         =   "&Guardar"
      End
      Begin Menu mnuFSaveAs 
         Caption         =   "Guardar &como..."
      End
      Begin Menu mnuFSep 
         Caption         =   "-"
      End
      Begin Menu mnuFExit 
         Caption         =   "&Salir"
      End
      Begin Menu mnuRecentFile 
         Caption         =   "-"
         Index           =   0
         Visible         =   0   'False
      End
      Begin Menu mnuRecentFile 
         Caption         =   "RecentFile1"
         Index           =   1
         Visible         =   0   'False
      End
      Begin Menu mnuRecentFile 
         Caption         =   "RecentFile2"
         Index           =   2
         Visible         =   0   'False
      End
      Begin Menu mnuRecentFile 
         Caption         =   "RecentFile3"
         Index           =   3
         Visible         =   0   'False
      End
      Begin Menu mnuRecentFile 
         Caption         =   "RecentFile4"
         Index           =   4
         Visible         =   0   'False
      End
   End
   Begin Menu mnuEdit 
      Caption         =   "&Edición"
      Begin Menu mnuECut 
         Caption         =   "Cor&tar"
         Shortcut        =   ^X
      End
      Begin Menu mnuECopy 
         Caption         =   "&Copiar"
         Shortcut        =   ^C
      End
      Begin Menu mnuEPaste 
         Caption         =   "&Pegar"
         Shortcut        =   ^V
      End
      Begin Menu mnuEDelete 
         Caption         =   "Borrar"
         Shortcut        =   {DEL}
      End
      Begin Menu mnuESep1 
         Caption         =   "-"
      End
      Begin Menu mnuESelectAll 
         Caption         =   "Selecionar &Todo"
      End
      Begin Menu mnuETime 
         Caption         =   "Hora/&Fecha"
      End
      Begin Menu mnuESep2 
         Caption         =   "-"
      End
      Begin Menu mnuSFind 
         Caption         =   "&Buscar"
      End
      Begin Menu mnuSFindNext 
         Caption         =   "Buscar &siguiente palabra"
         Shortcut        =   {F3}
      End
   End
   Begin Menu mnuOptions 
      Caption         =   "&Opciones"
      Begin Menu mnuOToolbar 
         Caption         =   "&Barra de herramientas"
      End
      Begin Menu mnuFont 
         Caption         =   "&Fuentes"
         Begin Menu mnuFontName 
            Caption         =   "FontName"
            Index           =   0
         End
      End
   End
   Begin Menu mnuWindow 
      Caption         =   "&Ventana"
      WindowList      =   -1  'True
      Begin Menu mnuWCascade 
         Caption         =   "&Cascada"
      End
      Begin Menu mnuWTile 
         Caption         =   "&Mosaico"
      End
      Begin Menu mnuWArrange 
         Caption         =   "&Organizar ventanas"
      End
   End
End

Sub Form_Load ()
    Dim i As Integer
    
    mnuFontName(0).Caption = screen.Fonts(0)
    For i = 1 To screen.FontCount - 1
	Load mnuFontName(i)
	mnuFontName(0).Caption = screen.Fonts(i)
    Next
End Sub

Sub Form_QueryUnload (Cancel As Integer, UnloadMode As Integer)
    Dim Msg, Filename, NL
    Dim Response As Integer

    If FState(Me.Tag).Dirty Then
	Filename = Me.Caption
	NL = Chr$(10) & Chr$(13)
	Msg = "El archivo [" & Filename & "] ha cambiado."
	Msg = Msg & NL
	Msg = Msg & "Grabar los cambios"
	Response = MsgBox(Msg, 51, feditor.Caption)
	Select Case Response
	' User selects Yes
	Case 6
	    'Get the filename to save the file
	    Filename = GetFileName()
	    'If the user did notspecify a file name,
	    'cancel the unload; otherwise, save it.
	    If Filename = "" Then
		Cancel = True
	    Else
		SaveFileAs (Filename)
	    End If

	' User selects No
	' Ok to unload
	Case 7
	    Cancel = False
	' User selects Cancel
	' Cancel the unload
	Case 2
	    Cancel = True
	End Select
    End If
End Sub

Sub Form_Resize ()
    If windowstate <> 1 And ScaleHeight <> 0 Then
	Text1.Visible = False
	Text1.Height = ScaleHeight
	Text1.Width = ScaleWidth
	Text1.Visible = True
    End If
End Sub

Sub Form_Unload (Cancel As Integer)
    FState(Me.Tag).Deleted = True
    
'Hide toolbar edit buttons if no notepad windows
    If Not AnyPadsLeft() Then
	feditor!imgCutButton.Visible = False
	feditor!imgCopyButton.Visible = False
	feditor!imgPasteButton.Visible = False
    End If
End Sub

Sub mnuECopy_Click ()
    EditCopyProc
End Sub

Sub mnuECut_Click ()
    EditCutProc
End Sub

Sub mnuEDelete_Click ()
  ' If cursor is not at the end of the notepad.
  If screen.ActiveControl.SelStart <> Len(screen.ActiveControl.Text) Then
    ' If nothing is selected, extend selection by one.
    If screen.ActiveControl.SelLength = 0 Then
      screen.ActiveControl.SelLength = 1
      ' If cursor is on a blank line, extend selection by two.
      If Asc(screen.ActiveControl.SelText) = 13 Then
	screen.ActiveControl.SelLength = 2
      End If
    End If
    ' Delete selected text.
    screen.ActiveControl.SelText = ""
  End If
End Sub

Sub mnuEPaste_Click ()
    EditPasteProc
End Sub

Sub mnuESelectAll_Click ()
    feditor.ActiveForm.Text1.SelStart = 0
    feditor.ActiveForm.Text1.SelLength = Len(feditor.ActiveForm.Text1.Text)
End Sub

Sub mnuETime_Click ()
    Dim TimeStr As String, DateStr As String
    
    Text1.SelText = Now
End Sub

Sub mnuFClose_Click ()
    Unload Me
End Sub

Sub mnuFExit_Click ()
    ' Unloading the MDI form invokes the QueryUnload event
    ' for each child form, then the MDI form - before unloading
    ' the MDI form. Setting the Cancel argument to True in any of the
    ' QueryUnload events aborts the unload.

    Unload feditor
End Sub

Sub mnuFNew_Click ()
    FileNew
End Sub

Sub mnuFontName_Click (index As Integer)
    Text1.FontName = mnuFontName(index).Caption
End Sub

Sub mnuFOpen_Click ()
    FOpenProc
End Sub

Sub mnuFSave_Click ()
    Dim Filename As String

    If Left(Me.Caption, 8) = "Untitled" Then
	' The file hasn't been saved yet,
	' get the filename, then call the
	' save procedure
	Filename = GetFileName()
    Else
	' The caption contains the name of the open file
	Filename = Me.Caption
    End If
    ' call the save procedure, if Filename = Empty then
    ' the user selected Cancel in the Save As dialog, otherwise
    ' save the file
    If Filename <> "" Then
	SaveFileAs Filename
    End If
End Sub

Sub mnuFSaveAs_Click ()
    Dim SaveFileName As String

    SaveFileName = GetFileName()
    If SaveFileName <> "" Then SaveFileAs (SaveFileName)
    ' Update the recent files menu
    UpdateFileMenu (SaveFileName)
End Sub

Sub mnuOptions_Click ()
    mnuOToolbar.Checked = feditor!picToolbar.Visible
End Sub

Sub mnuOToolbar_Click ()
    OptionsToolbarProc Me
End Sub

Sub mnuRecentFile_Click (index As Integer)
  OpenFile (mnuRecentFile(index).Caption)
  ' Update recent files list for new notepad.
  GetRecentFiles
End Sub

Sub mnuSFind_Click ()
    If Me!Text1.SelText <> "" Then
	fbuscar!Text1.Text = Me!Text1.SelText
    Else
	fbuscar!Text1.Text = FindString
    End If
    gFirstTime = True
    fbuscar.Show
End Sub

Sub mnuSFindNext_Click ()
    If Len(gFindString) > 0 Then
	FindIt
    Else
	mnuSFind_Click
    End If
End Sub

Sub mnuWArrange_Click ()
    feditor.Arrange ARRANGE_ICONS
End Sub

Sub mnuWCascade_Click ()
    feditor.Arrange CASCADE
End Sub

Sub mnuWTile_Click ()
    feditor.Arrange TILE_HORIZONTAL
End Sub

Sub Text1_Change ()
    FState(Me.Tag).Dirty = True
End Sub

Sub Text1_GotFocus ()
    If fbuscar.Visible Then
	fbuscar.ZOrder 0
    End If
End Sub

