Option Explicit

Type POINTAPI
    x As Integer
    y As Integer
End Type

Declare Function SendMessage Lib "User" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long
Global Const WM_USER = &H400
Global Const EM_LIMITTEXT = WM_USER + 21
Global Const EM_CANUNDO = WM_USER + 22
Global Const EM_UNDO = WM_USER + 23

Declare Sub SetWindowPos Lib "User" (ByVal h1%, ByVal h2%, ByVal x%, ByVal y%, ByVal cx%, ByVal cy%, ByVal f%)
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2
Global Const SWP_NOACTIVATE = &H10
Global Const SWP_SHOWWINDOW = &H40

Global gbFound As Integer
Global gbMatchCase As Integer
Global gbWholeWord As Integer
Global gsFindText As String

Global gbMainActive As Integer

Global gcForm As frmDocument

Sub CheckEditMenu ()
Dim bSel As Integer
Dim bUndo As Integer

    If frmMain.ActiveForm Is Nothing Then
	frmMain!mnuEditUndo.Enabled = False
	frmMain!mnuEditCut.Enabled = False
	frmMain!mnuEditCopy.Enabled = False
	frmMain!mnuEditPaste.Enabled = False
	frmMain!mnuEditDelete.Enabled = False
	frmMain!mnuEditSelectAll.Enabled = False
	frmMain!mnuEditDateTime.Enabled = False
	frmMain!mnuFind.Enabled = False
	frmMain!mnuFileSave.Enabled = False
	frmMain!mnuFileSaveAs.Enabled = False
	frmMain!mnuFilePrint.Enabled = False
    Else
	bSel = frmMain.ActiveForm!txtDocument.SelLength > 0
	bUndo = SendMessage(frmMain.ActiveForm!txtDocument.hWnd, EM_CANUNDO, 0, 0)
	frmMain!mnuEditUndo.Enabled = bUndo
	frmMain!mnuEditCut.Enabled = bSel
	frmMain!mnuEditCopy.Enabled = bSel
	frmMain!mnuEditPaste.Enabled = (ClipBoard.GetText() <> "")
	frmMain!mnuFindFindNext.Enabled = gbFound
	frmMain!mnuEditDelete.Enabled = bSel
	frmMain!mnuEditSelectAll.Enabled = True
	frmMain!mnuEditDateTime.Enabled = True
	frmMain!mnuFind.Enabled = True
	frmMain!mnuFileSave.Enabled = True
	frmMain!mnuFileSaveAs.Enabled = True
	frmMain!mnuFilePrint.Enabled = True
	frmMain!ButtonBar1.ButtonState(7) = frmMain.ActiveForm!txtDocument.FontBold
	frmMain!ButtonBar1.ButtonState(8) = frmMain.ActiveForm!txtDocument.FontItalic
	frmMain!ButtonBar1.ButtonState(9) = frmMain.ActiveForm!txtDocument.FontUnderLine
    End If
    frmMain!ButtonBar1.ButtonEnabled(3) = frmMain!mnuEditCut.Enabled
    frmMain!ButtonBar1.ButtonEnabled(4) = frmMain!mnuEditCopy.Enabled
    frmMain!ButtonBar1.ButtonEnabled(5) = frmMain!mnuEditPaste.Enabled
    frmMain!ButtonBar1.ButtonEnabled(2) = frmMain!mnuFileSave.Enabled
End Sub

Sub FindText ()
Dim iPos As Integer
Dim sText As String
Dim sFind As String
Dim I As Integer
Dim iComp As Integer
    If frmMain.ActiveForm Is Nothing Then
	frmMain!mnuFind.Enabled = False
	Beep
	Exit Sub
    End If
    frmMain.ActiveForm!lblAction.Caption = "get findposition"
    iPos = Val(frmMain.ActiveForm!lblAction.Caption)
    sText = frmMain.ActiveForm!txtDocument.Text
    sFind = gsFindText
    gbFound = True
    If gbMatchCase = 0 Then iComp = 1 Else iComp = 0
StartFind:
    I = InStr(iPos + 2, sText, sFind, iComp)
    If I > iPos Then
	If gbWholeWord Then
	    If InStr(";,. !?:", Mid$(sText, I - 1, 1)) > 0 And InStr(";., !?:", Mid$(sText, I + Len(sFind), 1)) > 0 Then
		frmMain.ActiveForm.txtDocument.SelStart = I - 1
		frmMain.ActiveForm.txtDocument.SelLength = Len(sFind)
	    Else
		iPos = I
		GoTo StartFind
	    End If
	Else
	    frmMain.ActiveForm.txtDocument.SelStart = I - 1
	    frmMain.ActiveForm.txtDocument.SelLength = Len(sFind)
	End If
    Else
	Beep
    End If
    CheckEditMenu
End Sub

Sub InitFind ()
    gbFound = False
End Sub

