Attribute VB_Name = "modUpdate"
Option Explicit

' ---------------------------------------------------------------------------
' modUpdate - the Check for updates button.
'
' Fetches the tool's modules from the address in "Update source" on the
' Setup sheet, says whether anything has changed, and on a Yes installs them
' into this workbook, rebuilds the Setup sheet and saves.
'
' The address can be:
'
'   https://github.com/<owner>/<repo>               main branch
'   https://github.com/<owner>/<repo>/tree/<branch> that branch
'   https://raw.githubusercontent.com/<owner>/<repo>/<branch>
'   a folder, e.g. a synced SharePoint folder       holding the repo's files
'
' Either way the modules are read from its vba folder. A GitHub repository
' has to be public: a private one needs a login, and a token kept in a
' spreadsheet that gets copied to every project is not a secret for long.
'
' This module calls nothing in the other three, on purpose: they are the
' ones being replaced while it runs. It cannot replace itself while it is
' running either, so a new copy of it is written to a temporary file, its path
' noted in the hidden name ToolPendingUpdater, and AfterToolUpdate in modMain
' swaps it in once this has finished. Those two names, and the ToolSource_
' checksums, are the whole contract between the two modules.
'
' Whether anything changed is decided by the code itself, not the version
' number, so an update with the same TOOL_VERSION is still found.
' ---------------------------------------------------------------------------

Private Const TOOL_MODULES As String = "modUtil,modSchedule,modMain"
Private Const SELF_MODULE  As String = "modUpdate"
Private Const LBL_SOURCE   As String = "Update source"

' Workbook names (hidden) the updater keeps its notes in.
Private Const SUM_PREFIX   As String = "ToolSource_"        ' checksum of what was installed
Private Const PENDING_NAME As String = "ToolPendingUpdater" ' new modUpdate waiting to go in


Public Sub CheckForUpdates()
    Dim src As String, base As String, why As String
    Dim names As Variant
    Dim code() As String
    Dim i As Long, n As Long, k As Long
    Dim changed As String
    Dim selfCode As String, selfChanged As Boolean
    Dim localVer As String, remoteVer As String
    Dim backupPath As String, tempPath As String
    Const JOB As String = "Check for updates"

    If Not CanEditCode(why) Then
        MsgBox why, vbExclamation, JOB
        Exit Sub
    End If

    src = SourceSetting()
    If Len(src) = 0 Then
        MsgBox "Put the address to update from in '" & LBL_SOURCE & "' on the Setup " & _
               "sheet first: the GitHub address of the tool, or a folder holding a copy " & _
               "of it.", vbExclamation, JOB
        Exit Sub
    End If
    base = SourceBase(src)

    ' --- fetch everything first; nothing is touched unless all of it arrives
    names = Split(TOOL_MODULES, ",")
    n = UBound(names)
    ReDim code(0 To n)
    For i = 0 To n
        Application.StatusBar = "Checking for updates: " & names(i)
        code(i) = Fetch(base, CStr(names(i)), why)
        If Len(code(i)) = 0 Then
            Application.StatusBar = False
            MsgBox "Could not get " & names(i) & " from" & vbCrLf & vbCrLf & base & _
                   vbCrLf & vbCrLf & why & vbCrLf & vbCrLf & "Nothing was changed.", _
                   vbExclamation, JOB
            Exit Sub
        End If
        If Not IsDifferent(CStr(names(i)), code(i)) Then
            ' unchanged
        Else
            changed = changed & "    " & names(i) & vbCrLf
        End If
    Next i

    ' A source without the updater in it is fine; it just stays as it is.
    selfCode = Fetch(base, SELF_MODULE, why)
    If Len(selfCode) > 0 Then selfChanged = IsDifferent(SELF_MODULE, selfCode)
    Application.StatusBar = False

    localVer = VersionIn(ModuleText("modUtil"))
    remoteVer = VersionIn(code(IndexOf(names, "modUtil")))

    If Len(changed) = 0 And Not selfChanged Then
        MsgBox "Up to date." & vbCrLf & vbCrLf & "Version " & localVer & ", the same as " & _
               base, vbInformation, JOB
        Exit Sub
    End If

    If MsgBox("An update is available." & vbCrLf & vbCrLf & _
              "    Installed:   " & localVer & vbCrLf & _
              "    Available:   " & remoteVer & _
              IIf(localVer = remoteVer, "  (same number, different code)", "") & vbCrLf & vbCrLf & _
              "From " & base & vbCrLf & vbCrLf & _
              "Install it into this workbook?" & vbCrLf & vbCrLf & _
              "A copy of this workbook is saved first. Afterwards the Setup sheet and " & _
              "buttons are rebuilt, your values on it are kept, and the file is saved.", _
              vbQuestion + vbYesNo, JOB) = vbNo Then Exit Sub

    ' --- a way back --------------------------------------------------------
    backupPath = BackupCopy()
    If Len(backupPath) = 0 Then
        If MsgBox("Could not save a backup copy of this workbook first." & vbCrLf & vbCrLf & _
                  "Install anyway?", vbExclamation + vbYesNo + vbDefaultButton2, JOB) = vbNo Then Exit Sub
    End If

    ' --- everything that must happen after the swap is arranged BEFORE it,
    ' because changing the code of a running project can stop it outright.
    If selfChanged Then
        tempPath = WriteTemp(SELF_MODULE, selfCode)
        If Len(tempPath) > 0 Then SetNote PENDING_NAME, tempPath
    End If
    For i = 0 To n
        SetNote SUM_PREFIX & names(i), Checksum(code(i))
    Next i
    Application.OnTime Now, "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!AfterToolUpdate"

    ' --- the swap. All three go in together, never a mixed set -------------
    If Len(changed) = 0 Then Exit Sub       ' only the updater changed
    For i = 0 To n
        why = ReplaceModule(CStr(names(i)), code(i))
        If Len(why) > 0 Then
            For k = 0 To n
                ClearNote SUM_PREFIX & names(k)
            Next k
            MsgBox "The update stopped part way, at " & names(i) & ":" & vbCrLf & vbCrLf & _
                   why & vbCrLf & vbCrLf & _
                   "Close this workbook WITHOUT saving and use the copy saved before the " & _
                   "update:" & vbCrLf & vbCrLf & backupPath, vbCritical, JOB
            Exit Sub
        End If
    Next i
End Sub


' ===========================================================================
' Where from
' ===========================================================================

Private Function SourceSetting() As String
    Dim ws As Worksheet
    Dim r As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Setup")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    For r = 1 To 40
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), LBL_SOURCE, vbTextCompare) = 0 Then
            SourceSetting = Trim$(CStr(ws.Cells(r, 2).Value))
            Exit Function
        End If
    Next r
End Function


' What the files are read from: a raw.githubusercontent.com address with the
' branch on the end, or a folder. An ordinary github.com address, the one in
' the browser, is turned into the raw one.
Private Function SourceBase(ByVal src As String) As String
    Dim parts As Variant
    Dim rest As String, branch As String, repo As String

    src = Trim$(src)
    Do While Right$(src, 1) = "/" Or Right$(src, 1) = "\"
        src = Left$(src, Len(src) - 1)
    Loop

    If LCase$(Left$(src, 19)) = "https://github.com/" Then
        rest = Mid$(src, 20)
        parts = Split(rest, "/")
        If UBound(parts) >= 1 Then
            repo = CStr(parts(1))
            If LCase$(Right$(repo, 4)) = ".git" Then repo = Left$(repo, Len(repo) - 4)
            branch = "main"
            If UBound(parts) >= 3 Then
                If LCase$(CStr(parts(2))) = "tree" Then branch = CStr(parts(3))
            End If
            SourceBase = "https://raw.githubusercontent.com/" & parts(0) & "/" & repo & "/" & branch
            Exit Function
        End If
    End If

    SourceBase = src
End Function


Private Function IsWeb(ByVal base As String) As Boolean
    IsWeb = (LCase$(Left$(base, 8)) = "https://" Or LCase$(Left$(base, 7)) = "http://")
End Function


' One module's source text, or "" with `why` saying what went wrong. Checked
' to really be that module, so a "page not found" or a sign-in page is never
' installed as code.
Private Function Fetch(ByVal base As String, ByVal moduleName As String, _
                       ByRef why As String) As String
    Dim txt As String
    Dim http As Object
    Dim p As String

    why = ""
    On Error GoTo Failed

    If IsWeb(base) Then
        Set http = CreateObject("MSXML2.XMLHTTP.6.0")
        ' The time on the end stops Windows handing back a cached copy.
        http.Open "GET", base & "/vba/" & moduleName & ".bas?t=" & Format$(Now, "yyyymmddhhnnss"), False
        http.setRequestHeader "Cache-Control", "no-cache"
        http.Send
        If http.Status <> 200 Then
            why = "The server said " & http.Status & " " & http.statusText & "." & _
                  IIf(http.Status = 404, " Check the address, and that the repository " & _
                      "is public: a private one cannot be read without signing in.", "")
            Exit Function
        End If
        txt = http.responseText
    Else
        p = base & Application.PathSeparator & "vba" & Application.PathSeparator & moduleName & ".bas"
        If Len(Dir$(p)) = 0 Then p = base & Application.PathSeparator & moduleName & ".bas"
        If Len(Dir$(p)) = 0 Then
            why = "There is no " & moduleName & ".bas in that folder or its vba folder."
            Exit Function
        End If
        txt = ReadFile(p)
    End If

    txt = Replace(txt, vbCr, "")
    If AscW(Left$(txt & " ", 1)) = &HFEFF Then txt = Mid$(txt, 2)      ' byte order mark

    If InStr(1, txt, "Attribute VB_Name = """ & moduleName & """", vbTextCompare) <> 1 _
       Or InStr(1, txt, "Option Explicit", vbTextCompare) = 0 Then
        why = "What came back is not the " & moduleName & " module (a sign-in page or an " & _
              "error page, most likely)."
        Exit Function
    End If

    Fetch = txt
    Exit Function

Failed:
    why = "Error " & Err.Number & " - " & Err.Description & _
          IIf(IsWeb(base), " The network may be blocking that address.", "")
End Function


' ===========================================================================
' Has it changed?
' ===========================================================================

' True when `code` is not what is installed. Compared with the checksum noted
' at the last update when there is one. The first time, there is none, so it
' is compared with the code in the module, ignoring case, because the VBA
' editor re-cases names as it pleases.
Private Function IsDifferent(ByVal moduleName As String, ByVal code As String) As Boolean
    Dim noted As String
    Dim installed As String

    noted = GetNote(SUM_PREFIX & moduleName)
    If Len(noted) > 0 Then
        IsDifferent = (noted <> Checksum(code))
        Exit Function
    End If

    installed = ModuleText(moduleName)
    If Len(installed) = 0 Then
        IsDifferent = True
    Else
        IsDifferent = (StrComp(Normalised(installed), Normalised(code), vbTextCompare) <> 0)
        If Not IsDifferent Then SetNote SUM_PREFIX & moduleName, Checksum(code)
    End If
End Function


' The code as it is meant to read: no Attribute lines (the editor hides them),
' no trailing spaces, no blank lines at the end, LF line ends.
Private Function Normalised(ByVal txt As String) As String
    Dim lines As Variant
    Dim keep() As String
    Dim i As Long, n As Long, last As Long
    Dim ln As String

    lines = Split(Replace(txt, vbCr, ""), vbLf)
    ReDim keep(0 To UBound(lines) + 1)
    last = -1
    For i = LBound(lines) To UBound(lines)
        ln = RTrim$(CStr(lines(i)))
        If LCase$(Left$(ln, 13)) <> "attribute vb_" Then
            keep(n) = ln
            If Len(ln) > 0 Then last = n
            n = n + 1
        End If
    Next i

    If last < 0 Then Exit Function
    ReDim Preserve keep(0 To last)
    Normalised = Join(keep, vbLf)
End Function


' Adler-32 of the normalised code, as "length:b:a". Not for security, just a
' short fingerprint that changes when the code does.
Private Function Checksum(ByVal txt As String) As String
    Dim a As Long, b As Long, i As Long
    Dim s As String

    s = Normalised(txt)
    a = 1
    For i = 1 To Len(s)
        a = (a + (AscW(Mid$(s, i, 1)) And &HFFFF&)) Mod 65521
        b = (b + a) Mod 65521
    Next i
    Checksum = Len(s) & ":" & b & ":" & a
End Function


Private Function ModuleText(ByVal moduleName As String) As String
    Dim cm As Object
    On Error Resume Next
    Set cm = ThisWorkbook.VBProject.VBComponents(moduleName).CodeModule
    If cm Is Nothing Then Exit Function
    If cm.CountOfLines > 0 Then ModuleText = cm.Lines(1, cm.CountOfLines)
    On Error GoTo 0
End Function


Private Function VersionIn(ByVal code As String) As String
    Dim p As Long, q As Long
    p = InStr(1, code, "TOOL_VERSION As String = """, vbTextCompare)
    If p = 0 Then
        VersionIn = "(unknown)"
        Exit Function
    End If
    p = p + Len("TOOL_VERSION As String = """)
    q = InStr(p, code, """")
    If q > p Then VersionIn = Mid$(code, p, q - p) Else VersionIn = "(unknown)"
End Function


Private Function IndexOf(ByVal arr As Variant, ByVal wanted As String) As Long
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If StrComp(CStr(arr(i)), wanted, vbTextCompare) = 0 Then
            IndexOf = i
            Exit Function
        End If
    Next i
End Function


' ===========================================================================
' Putting it in
' ===========================================================================

' Excel only lets code change code when "Trust access to the VBA project
' object model" is ticked, per person. It cannot be ticked from code.
Private Function CanEditCode(ByRef why As String) As Boolean
    Dim n As Long

    If ThisWorkbook.ReadOnly Then
        why = "This workbook is open read only, so it cannot be updated. Open it normally."
        Exit Function
    End If

    On Error Resume Next
    n = ThisWorkbook.VBProject.VBComponents.Count
    If Err.Number <> 0 Or n = 0 Then
        Err.Clear
        why = "Excel is not letting the tool update its own code." & vbCrLf & vbCrLf & _
              "To allow it, once on this PC: File > Options > Trust Center > Trust " & _
              "Center Settings > Macro Settings, tick 'Trust access to the VBA project " & _
              "object model', OK, and press this button again." & vbCrLf & vbCrLf & _
              "If that box is greyed out, IT have locked it. Update by hand instead: " & _
              "remove modUtil, modSchedule and modMain in the VBA editor (Alt+F11), " & _
              "import the new ones, and run InstallTool."
        Exit Function
    End If
    On Error GoTo 0
    CanEditCode = True
End Function


' Swaps one module for new code. The old one is renamed before it is removed,
' because Excel only finishes removing a module after the macro ends, and
' importing a module of the same name before then gets "modMain1".
Private Function ReplaceModule(ByVal moduleName As String, ByVal code As String) As String
    Dim comps As Object, old As Object
    Dim path As String

    path = WriteTemp(moduleName, code)
    If Len(path) = 0 Then
        ReplaceModule = "Could not write a temporary copy of " & moduleName & "."
        Exit Function
    End If

    On Error GoTo Failed
    Set comps = ThisWorkbook.VBProject.VBComponents

    On Error Resume Next
    Set old = comps(moduleName)
    On Error GoTo Failed
    If Not old Is Nothing Then
        old.Name = moduleName & "_old" & Format$(Now, "hhnnss")
        comps.Remove old
    End If

    comps.Import path
    On Error Resume Next
    Kill path
    Exit Function

Failed:
    ReplaceModule = "Error " & Err.Number & " - " & Err.Description
End Function


' The module written out with Windows line ends, which is what the VBA
' editor's Import expects. "" if it could not be written.
Private Function WriteTemp(ByVal moduleName As String, ByVal code As String) As String
    Dim path As String
    Dim f As Integer

    path = Environ$("TEMP") & Application.PathSeparator & moduleName & "_" & _
           Format$(Now, "yyyymmddhhnnss") & ".bas"
    On Error GoTo Failed
    f = FreeFile
    Open path For Output As #f
    Print #f, Replace(Replace(code, vbCr, ""), vbLf, vbCrLf);
    Close #f
    WriteTemp = path
    Exit Function

Failed:
    On Error Resume Next
    Close #f
End Function


Private Function ReadFile(ByVal path As String) As String
    Dim f As Integer
    On Error GoTo Failed
    f = FreeFile
    Open path For Binary Access Read As #f
    ReadFile = Space$(LOF(f))
    Get #f, , ReadFile
    Close #f
    Exit Function
Failed:
    On Error Resume Next
    Close #f
    ReadFile = ""
End Function


' A copy of this workbook as it is now, next to it in "_backup tool <date>",
' or in the temp folder when the workbook is open from a web address.
' Returns the copy's path, or "".
Private Function BackupCopy() As String
    Dim folder As String, path As String

    folder = ThisWorkbook.Path
    If Len(folder) = 0 Or LCase$(Left$(folder, 4)) = "http" Then folder = Environ$("TEMP")
    folder = folder & Application.PathSeparator & "_backup tool " & Format$(Now, "yyyy-mm-dd hh-nn")

    On Error Resume Next
    MkDir folder
    path = folder & Application.PathSeparator & ThisWorkbook.Name
    ThisWorkbook.SaveCopyAs path
    If Err.Number = 0 And Len(Dir$(path)) > 0 Then BackupCopy = path
    On Error GoTo 0
End Function


' ===========================================================================
' Notes kept in hidden workbook names
' ===========================================================================

Private Sub SetNote(ByVal noteName As String, ByVal value As String)
    ClearNote noteName
    On Error Resume Next
    ThisWorkbook.Names.Add Name:=noteName, RefersTo:="=""" & Replace(value, """", """""") & """", _
                           Visible:=False
    On Error GoTo 0
End Sub


Private Function GetNote(ByVal noteName As String) As String
    Dim v As String
    On Error Resume Next
    v = ThisWorkbook.Names(noteName).RefersTo
    On Error GoTo 0
    If Left$(v, 2) = "=""" And Right$(v, 1) = """" Then
        GetNote = Replace(Mid$(v, 3, Len(v) - 3), """""", """")
    End If
End Function


Private Sub ClearNote(ByVal noteName As String)
    On Error Resume Next
    ThisWorkbook.Names(noteName).Delete
    On Error GoTo 0
End Sub
