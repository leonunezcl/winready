
'---------------------------------------------------------------
' Copies file SrcFilename from SourcePath to DestinationPath.
'
' Returns 0 if it could not find the file, or other runtime
' error occurs.  Otherwise, returns true.
'
' If the source file is older, the function returns success (-1)
' even though no file was copied, since no error occurred.
'---------------------------------------------------------------
Function CopyFile (ByVal SourcePath As String, ByVal DestinationPath As String, ByVal SrcFilename As String, ByVal DestFileName As String)
' ----- VerInstallFile() flags -----
    
    Const VIFF_FORCEINSTALL% = &H1, VIFF_DONTDELETEOLD% = &H2
    Const OF_DELETE% = &H200
    Const VIF_TEMPFILE& = &H1
    Const VIF_MISMATCH& = &H2
    Const VIF_SRCOLD& = &H4

    Const VIF_DIFFLANG& = &H8
    Const VIF_DIFFCODEPG& = &H10
    Const VIF_DIFFTYPE& = &H20
    Const VIF_WRITEPROT& = &H40
    Const VIF_FILEINUSE& = &H80
    Const VIF_OUTOFSPACE& = &H100
    Const VIF_ACCESSVIOLATION& = &H200
    Const VIF_SHARINGVIOLATION = &H400
    Const VIF_CANNOTCREATE = &H800
    Const VIF_CANNOTDELETE = &H1000
    Const VIF_CANNOTRENAME = &H2000
    Const VIF_CANNOTDELETECUR = &H4000
    Const VIF_OUTOFMEMORY = &H8000

    Const VIF_CANNOTREADSRC = &H10000
    Const VIF_CANNOTREADDST = &H20000

    Const VIF_BUFFTOOSMALL = &H40000
    Dim TmpOFStruct As OFStruct
    On Error GoTo ErrorCopy

    screen.MousePointer = 11

    '--------------------------------------
    ' Add ending \ symbols to path variables
    '--------------------------------------
    If Right$(SourcePath$, 1) <> "\" Then
    SourcePath$ = SourcePath$ + "\"
    End If
    If Right$(DestinationPath$, 1) <> "\" Then
    DestinationPath$ = DestinationPath$ + "\"
    End If
    
    '----------------------------
    ' Update status dialog info
    '----------------------------
    desde.Caption = UCase$(SourcePath$ + SrcFilename$)
    desde.Refresh
    hacia.Caption = UCase$(DestinationPath$ + DestFileName$)
    hacia.Refresh

    '-----------------------------------------
    ' Check the validity of the path and file
    '-----------------------------------------
CheckForExist:
    If Not FileExists(SourcePath$ + SrcFilename$) Then
    screen.MousePointer = 0
    x% = MsgBox("Ha ocurrido un error al intentar instalar el archivo. No puedo encontrar : """ + SourcePath$ + SrcFilename$ + """", 34, "WinReady")
    screen.MousePointer = 11
    If x% = 3 Then
    CopyFile = False
    ElseIf x% = 4 Then
    GoTo CheckForExist
    ElseIf x% = 5 Then
    GoTo SkipThisFile
    End If
    Else
    '-------------------------------------------------
    ' VerInstallFile installs the file. We need to initialize
    ' some arguments for the temp file that is created by the call
    '-------------------------------------------------
TryToCopyAgain:
    CurrDir$ = String$(255, 0)
    TmpFile$ = String$(255, 0)
    lpwTempFileLen% = 255
    InFileVer$ = GetFileVersion(SourcePath$ + SrcFilename$)
    OutFileVer$ = GetFileVersion(DestinationPath$ + DestFileName$)
    
    ' Install if no version info is available
    If Len(InFileVer$) <> 0 And Len(OutFileVer$) <> 0 Then
    ' Don't install older or same version of file
    If InFileVer$ <= OutFileVer$ Then
    'UpdateStatus GetFileSize(SourcePath$ + SrcFilename$)
    CopyFile = True
    Exit Function
    End If
    End If

    Result& = VerInstallFile&(0, SrcFilename$, DestFileName$, SourcePath$, DestinationPath$, CurrDir$, TmpFile$, lpwTempFileLen%)

    '--------------------------------------------
    ' After copying, update the installation meter
    '---------------------------------------------
    
    S$ = DestinationPath$
    If Right$(S$, 1) <> "\" Then S$ = S$ + "\"
    S$ = S$ + DestFileName$
    'If Not TryAgain% Then Progreso Incre

    '--------------------------------
    ' There are many return values that you can test for.
    ' The constants are listed above.
    ' The following lines of code return will set the Function to
    ' True if the VerInstallFile call was successful.
    '
    ' If the call was unsuccessful due to a different language on the
    ' users machine, VerInstallFile is called again to force installation.
    ' You can change this to not install if you choose.
    ' Be careful about using FORCEINSTALL.  Other flags could be
    ' set which indicate that this file should not be overridden.
    '
    ' Under any other circumstance, the tempfile created by VerInstallFile
    ' is removed using OpenFile and the CopyFile function returns false.
    '--------------------------------------------------------
    
    If Result& = 0 Or (Result& And VIF_SRCOLD&) = VIF_SRCOLD& Then
    CopyFile = True
    ElseIf (Result& And VIF_DIFFLANG&) = VIF_DIFFLANG& Then
    Result& = VerInstallFile&(VIFF_FORCEINSTALL%, SrcFilename$, DestFileName$, SourcePath$, DestinationPath$, CurrDir$, TmpFile$, lpwTempFileLen%)
    CopyFile = True
    ElseIf (Result& And VIF_WRITEPROT&) = VIF_WRITEPROT& Then
    Result& = VerInstallFile&(VIFF_FORCEINSTALL%, SrcFilename$, DestFileName$, SourcePath$, winSysDir$ + "\", CurrDir$, TmpFile$, lpwTempFileLen%)
    CopyFile = True
    ElseIf (Result& And VIF_CANNOTREADSRC) = VIF_CANNOTREADSRC Then
    ' VerInstallFile does will not handle compressed files that have been split.
    ' Use VB's FileCopy stmt
    FileCopy SourcePath$ + SrcFilename$, DestinationPath$ + DestFileName$
    CopyFile = True
    Else
    screen.MousePointer = 0
    If (Result& And VIF_FILEINUSE&) = VIF_FILEINUSE& Then
    x% = MsgBox(DestFileName$ & " está en uso. Cierre todas las aplicaciones y vuelva a instalar otra vez.", 34, "WinReady")
    If x% = 3 Then
	CopyFile = False
    ElseIf x% = 4 Then
	TryAgain% = True
	GoTo TryToCopyAgain
    ElseIf x% = 5 Then
	CopyFile = True
	GoTo SkipThisFile
    End If
    Else
    MsgBox DestFileName$ & " No pudo ser expandido.", 48, "WinReady"
    CopyFile = False
    End If
    screen.MousePointer = 11
    End If

    If (Result& And VIF_TEMPFILE&) = VIF_TEMPFILE& Then copyresult% = OpenFile(TmpFile$, TmpOFStruct, OF_DELETE%)
       screen.MousePointer = 0
       Exit Function
    End If

SkipThisFile:
       Exit Function
ErrorCopy:
    CopyFile = False
    screen.MousePointer = 0
    Exit Function

End Function

'----------------------------------------------------
' Get the disk Allocation unit for the current drive
'----------------------------------------------------
Function GetDrivesAllocUnit (drive As String) As Long
    ChDrive drive
    GetDrivesAllocUnit = AllocUnit()
End Function

'------------------------
' Get the size of the file
'------------------------
Function GetFileSize (source$) As Long
    x = FreeFile
    Open source$ For Binary Access Read As x
    GetFileSize = LOF(x)
    Close x
End Function

Function GetFileVersion (FileToCheck As String) As String
    On Error Resume Next
    VersionInfoSize& = GetFileVersionInfoSize(FileToCheck, lpdwHandle&)
    If VersionInfoSize& = 0 Then
    GetFileVersion = ""
    Exit Function
    End If
    lpvdata$ = String(VersionInfoSize&, Chr$(0))
    VersionInfo% = GetFileVersionInfo(FileToCheck, lpdwHandle&, VersionInfoSize&, lpvdata$)
    ptrFixed% = VerQueryValue(lpvdata$, "\FILEVERSION", lplpBuffer&, lpcb%)
    If ptrFixed% = 0 Then
    ' Take a shot with the hardcoded TransString
    TransString$ = "040904E4"
    ptrString% = VerQueryValue(lpvdata$, "\StringFileInfo\" & TransString$ & "\CompanyName", lplpBuffer&, lpcb%)
    If ptrString% <> 0 Then GoTo GetValues
    ptrFixed% = VerQueryValue(lpvdata$, "\", lplpBuffer&, lpcb%)
    If ptrFixed% = 0 Then
    GetFileVersion = ""
    Exit Function
    Else
    TransString$ = ""
    fixedstr$ = String(lpcb% + 1, Chr(0))
    stringcopy& = lstrcpyn(fixedstr$, lplpBuffer&, lpcb% + 1)
    For i = lpcb% To 1 Step -1
    char$ = Hex(Asc(Mid(fixedstr$, i, 1)))
    If Len(char$) = 1 Then
	char$ = "0" + char$
    End If
    TransString$ = TransString$ + char$
    If Len(TransString$ & nextchar$) Mod 8 = 0 Then
	TransString$ = "&H" & TransString$
	TransValue& = Val(TransString$)
	TransString$ = ""
    End If
    Next i
    End If
    End If
    TransTable$ = String(lpcb% + 1, Chr(0))
    TransString$ = String(0, Chr(0))
    stringcopy& = lstrcpyn(TransTable$, lplpBuffer&, lpcb% + 1)
    For i = 1 To lpcb%
    char$ = Hex(Asc(Mid(TransTable$, i, 1)))
    If Len(char$) = 1 Then
    char$ = "0" + char$
    End If
    If Len(TransString$ & nextchar$) Mod 4 = 0 Then
    nextchar$ = char$
    Else
    TransString$ = TransString$ + char$ + nextchar$
    nextchar$ = ""
    char$ = ""
    End If
    Next i
GetValues:
    ptrString% = VerQueryValue(lpvdata$, "\StringFileInfo\" & TransString$ & "\FileVersion", lplpBuffer&, lpcb%)
    If ptrString% = 1 Then
    TransTable$ = String(lpcb%, Chr(0))
    stringcopy& = lstrcpyn(TransTable$, lplpBuffer&, lpcb% + 1)
    GetFileVersion = TransTable$
    Else
    GetFileVersion = ""
    End If
End Function

'-----------------------------------------------------------------------------
' Set the Destination File's date and time to the Source file's date and time
'-----------------------------------------------------------------------------
Function SetFileDateTime (SourceFile As String, DestinationFile As String) As Integer
    x = SetTime(SourceFile, DestinationFile)
    SetFileDateTime = -1
End Function

