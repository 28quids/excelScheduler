Attribute VB_Name = "modMain"
Option Explicit

' ---------------------------------------------------------------------------
' modMain - the buttons.
'
'   InstallTool           run once, after importing the modules
'   SetupProject          link/repair every schedule in the folder
'   RefreshScheduleList   read every schedule back and QA it
'   AddRevisionToTicked   append one revision line to the ticked schedules
' ---------------------------------------------------------------------------

' Setup sheet layout. Rows 1-4 match the original sheet so existing links to
' $B$1 / $B$3 / $B$4 keep working.
Private Const R_OPT_FOLDER  As Long = 7
Private Const R_OPT_BACKUP  As Long = 8
Private Const R_OPT_AUTO    As Long = 9
Private Const R_OPT_SETUP   As Long = 10
Private Const R_OPT_LIST    As Long = 11
Private Const R_REV_FIRST   As Long = 14   ' Revision .. Description = 14..20

' ScheduleList columns
Private Const C_PICK As Long = 1
Private Const C_FILE As Long = 2
Private Const C_LAST As Long = 16          ' P = Checks


' ===========================================================================
' One-time install
' ===========================================================================
Public Sub InstallTool()
    Dim wsSetup As Worksheet
    Dim wsList As Worksheet

    Set wsSetup = EnsureSheet(SH_SETUP)
    Set wsList = EnsureSheet(SH_LIST)

    BuildSetupSheet wsSetup
    BuildListHeaders wsList
    BuildButtons wsSetup

    wsSetup.Activate
    MsgBox "Ready." & vbCrLf & vbCrLf & _
           "1. Fill in Client, Project Name and Project Number on this sheet." & vbCrLf & _
           "2. Save this file into the project folder with the schedules." & vbCrLf & _
           "3. Press 'Set up / repair schedules'.", vbInformation, "Schedule tool"
End Sub


' ===========================================================================
' Button 1 - point every schedule in the folder at this MAINPROJECTINFO
' ===========================================================================
Public Sub SetupProject()
    Dim folderPath As String
    Dim fileName As String, fullPath As String
    Dim wsSetup As Worksheet
    Dim wbTgt As Workbook
    Dim statuses As Variant
    Dim projNameRef As String, projNoRef As String, clientRef As String
    Dim setupSheetName As String
    Dim backupDir As String
    Dim log As String, oneLog As String
    Dim done As Long, skipped As Long, failed As Long
    Dim files As Collection, i As Long

    Set wsSetup = GetSheet(ThisWorkbook, SH_SETUP)
    If wsSetup Is Nothing Then
        MsgBox "No Setup sheet. Run InstallTool first.", vbExclamation
        Exit Sub
    End If
    setupSheetName = wsSetup.Name

    If Len(ThisWorkbook.Path) = 0 Then
        MsgBox "Save this file into the project folder first.", vbExclamation
        Exit Sub
    End If

    If Not SetupRefs(wsSetup, projNameRef, projNoRef, clientRef) Then Exit Sub

    If Len(Trim$(wsSetup.Range("B3").Value)) = 0 _
       Or Len(Trim$(wsSetup.Range("B4").Value)) = 0 Then
        If MsgBox("Project Name or Project Number is blank on the Setup sheet." & vbCrLf & _
                  "Carry on anyway?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If

    folderPath = SchedulesFolder()
    If Len(folderPath) = 0 Then Exit Sub

    Set files = ScheduleFiles(folderPath)
    If files.Count = 0 Then
        MsgBox "No other Excel files found in:" & vbCrLf & folderPath, vbInformation
        Exit Sub
    End If

    If MsgBox(PathFormWarning(folderPath) & _
              files.Count & " workbook(s) will be opened, relinked to this file and saved." & vbCrLf & vbCrLf & _
              folderPath & vbCrLf & vbCrLf & "Continue?", _
              vbQuestion + vbYesNo, "Set up / repair schedules") = vbNo Then Exit Sub

    statuses = GatherStatuses(wsSetup, files, folderPath)

    If UCase$(Trim$(CStr(wsSetup.Cells(R_OPT_BACKUP, 2).Value))) <> "NO" Then
        backupDir = EndSep(folderPath) & "_backup " & Format$(Now, "yyyy-mm-dd hh-nn")
        On Error Resume Next
        MkDir backupDir
        On Error GoTo 0
    End If

    BeginQuiet

    For i = 1 To files.Count
        fileName = files(i)
        fullPath = EndSep(folderPath) & fileName

        If Len(backupDir) > 0 Then
            On Error Resume Next
            FileCopy fullPath, EndSep(backupDir) & fileName
            On Error GoTo 0
        End If

        Set wbTgt = Nothing
        On Error Resume Next
        Set wbTgt = Workbooks.Open(fileName:=fullPath, ReadOnly:=False, UpdateLinks:=0)
        On Error GoTo 0

        If wbTgt Is Nothing Then
            failed = failed + 1
            log = log & vbCrLf & "FAILED to open: " & fileName
        ElseIf Not LooksLikeSchedule(wbTgt) Then
            wbTgt.Close SaveChanges:=False
            skipped = skipped + 1
        Else
            oneLog = ""
            On Error Resume Next
            oneLog = RepairWorkbook(wbTgt, ThisWorkbook.FullName, setupSheetName, _
                                    projNameRef, projNoRef, clientRef, statuses)
            If Err.Number <> 0 Then
                oneLog = "ERROR " & Err.Number & " - " & Err.Description
                Err.Clear
                failed = failed + 1
                wbTgt.Close SaveChanges:=False
            Else
                wbTgt.Close SaveChanges:=True
                done = done + 1
            End If
            On Error GoTo 0
            If Len(oneLog) > 0 Then log = log & vbCrLf & fileName & ": " & oneLog
        End If
    Next i

    wsSetup.Cells(R_OPT_SETUP, 2).Value = Format$(Now, "dd/mm/yyyy hh:nn")
    EndQuiet

    MsgBox "Set up " & done & " schedule(s)." & vbCrLf & _
           "Skipped (not a schedule): " & skipped & vbCrLf & _
           "Failed: " & failed & vbCrLf & _
           IIf(Len(backupDir) > 0, "Backup: " & backupDir & vbCrLf, "") & _
           IIf(Len(log) > 0, vbCrLf & "Notes:" & log, ""), _
           vbInformation, "Set up / repair schedules"

    RefreshScheduleList
End Sub


' ===========================================================================
' Button 2 - read every schedule back and check it
' ===========================================================================
Public Sub RefreshScheduleList()
    Dim wsSetup As Worksheet, wsList As Worksheet
    Dim folderPath As String, fileName As String
    Dim wbTgt As Workbook
    Dim files As Collection, i As Long
    Dim r As Long
    Dim mpiName As String, mpiNo As String, mpiClient As String
    Dim issues As String

    Set wsSetup = GetSheet(ThisWorkbook, SH_SETUP)
    Set wsList = EnsureSheet(SH_LIST)
    If wsSetup Is Nothing Then Exit Sub

    folderPath = SchedulesFolder()
    If Len(folderPath) = 0 Then Exit Sub

    mpiName = Trim$(CStr(wsSetup.Range("B3").Value))
    mpiNo = Trim$(CStr(wsSetup.Range("B4").Value))
    mpiClient = Trim$(CStr(wsSetup.Range("B1").Value))

    Set files = ScheduleFiles(folderPath)

    BeginQuiet
    BuildListHeaders wsList
    wsList.Range(wsList.Cells(2, 1), wsList.Cells(wsList.Rows.Count, C_LAST)).Clear

    r = 2
    For i = 1 To files.Count
        fileName = files(i)
        Set wbTgt = Nothing
        On Error Resume Next
        Set wbTgt = Workbooks.Open(fileName:=EndSep(folderPath) & fileName, _
                                   ReadOnly:=True, UpdateLinks:=0)
        On Error GoTo 0

        If Not wbTgt Is Nothing Then
            If LooksLikeSchedule(wbTgt) Then
                wsList.Cells(r, C_FILE).Value = fileName
                wsList.Cells(r, 3).Value = ReadMeta(wbTgt, "ScheduleName")
                wsList.Cells(r, 4).Value = ReadMeta(wbTgt, "Project Name")
                wsList.Cells(r, 5).Value = ReadMeta(wbTgt, "Project Number")
                wsList.Cells(r, 6).Value = ReadMeta(wbTgt, "Client")
                wsList.Cells(r, 7).Value = ReadMeta(wbTgt, "DocumentType")
                wsList.Cells(r, 8).Value = ReadMeta(wbTgt, "Revision")
                wsList.Cells(r, 9).Value = ReadMeta(wbTgt, "Date")
                wsList.Cells(r, 10).Value = ReadMeta(wbTgt, "Prepared by")
                wsList.Cells(r, 11).Value = ReadMeta(wbTgt, "Checked by")
                wsList.Cells(r, 12).Value = ReadMeta(wbTgt, "Approved by")
                wsList.Cells(r, 13).Value = ReadMeta(wbTgt, "DocumentNumber")
                wsList.Cells(r, 14).Value = ReadMeta(wbTgt, "Suitability Status")
                wsList.Cells(r, 15).Value = ReadMeta(wbTgt, "Suitability Description")

                issues = CheckWorkbook(wbTgt, fileName, mpiName, mpiNo, mpiClient)
                wsList.Cells(r, C_LAST).Value = issues
                r = r + 1
            End If
            wbTgt.Close SaveChanges:=False
        End If
    Next i

    FlagOddOnesOut wsList, r - 1
    ColourIssues wsList, r - 1

    wsList.Columns("A:P").AutoFit
    If wsList.Columns(C_LAST).ColumnWidth > 60 Then wsList.Columns(C_LAST).ColumnWidth = 60
    wsSetup.Cells(R_OPT_LIST, 2).Value = Format$(Now, "dd/mm/yyyy hh:nn")

    EndQuiet
    wsList.Activate
End Sub


' ===========================================================================
' Button 3 - append one revision line to every ticked schedule
' ===========================================================================
Public Sub AddRevisionToTicked()
    Dim wsSetup As Worksheet, wsList As Worksheet
    Dim rev As String, status As String, prep As String, chk As String
    Dim app As String, descr As String
    Dim issueDate As Variant
    Dim r As Long, lastRow As Long, n As Long
    Dim folderPath As String, fileName As String
    Dim wbTgt As Workbook
    Dim res As String, log As String

    Set wsSetup = GetSheet(ThisWorkbook, SH_SETUP)
    Set wsList = GetSheet(ThisWorkbook, SH_LIST)
    If wsSetup Is Nothing Or wsList Is Nothing Then Exit Sub

    rev = Trim$(CStr(wsSetup.Cells(R_REV_FIRST + 0, 2).Value))
    status = Trim$(CStr(wsSetup.Cells(R_REV_FIRST + 1, 2).Value))
    issueDate = wsSetup.Cells(R_REV_FIRST + 2, 2).Value
    prep = Trim$(CStr(wsSetup.Cells(R_REV_FIRST + 3, 2).Value))
    chk = Trim$(CStr(wsSetup.Cells(R_REV_FIRST + 4, 2).Value))
    app = Trim$(CStr(wsSetup.Cells(R_REV_FIRST + 5, 2).Value))
    descr = Trim$(CStr(wsSetup.Cells(R_REV_FIRST + 6, 2).Value))

    If Len(rev) = 0 Or Len(status) = 0 Or Not IsDate(issueDate) Then
        MsgBox "Fill in at least Revision, Status and a valid Date in the " & _
               "'New revision' block on the Setup sheet.", vbExclamation
        Exit Sub
    End If

    lastRow = wsList.Cells(wsList.Rows.Count, C_FILE).End(xlUp).Row
    For r = 2 To lastRow
        If Len(Trim$(CStr(wsList.Cells(r, C_PICK).Value))) > 0 Then n = n + 1
    Next r

    If n = 0 Then
        MsgBox "Nothing ticked. Put an x in the 'Add?' column on ScheduleList " & _
               "next to the schedules that are being reissued.", vbExclamation
        Exit Sub
    End If

    If MsgBox("Add revision " & rev & " (" & Format$(issueDate, "dd/mm/yyyy") & ") to " & _
              n & " schedule(s)?", vbQuestion + vbYesNo) = vbNo Then Exit Sub

    folderPath = SchedulesFolder()
    If Len(folderPath) = 0 Then Exit Sub
    BeginQuiet

    For r = 2 To lastRow
        If Len(Trim$(CStr(wsList.Cells(r, C_PICK).Value))) > 0 Then
            fileName = CStr(wsList.Cells(r, C_FILE).Value)
            Set wbTgt = Nothing
            On Error Resume Next
            Set wbTgt = Workbooks.Open(fileName:=EndSep(folderPath) & fileName, _
                                       ReadOnly:=False, UpdateLinks:=0)
            On Error GoTo 0
            If wbTgt Is Nothing Then
                log = log & vbCrLf & fileName & ": could not open"
            Else
                res = AppendRevision(wbTgt, rev, status, CDate(issueDate), prep, chk, app, descr)
                If Len(res) = 0 Then
                    wbTgt.Close SaveChanges:=True
                Else
                    wbTgt.Close SaveChanges:=False
                    log = log & vbCrLf & fileName & ": " & res
                End If
            End If
        End If
    Next r

    EndQuiet
    MsgBox "Done." & IIf(Len(log) > 0, vbCrLf & vbCrLf & "Notes:" & log, ""), vbInformation
    RefreshScheduleList
End Sub


Public Sub Auto_Open()
    Dim wsSetup As Worksheet
    Set wsSetup = GetSheet(ThisWorkbook, SH_SETUP)
    If wsSetup Is Nothing Then Exit Sub
    If UCase$(Trim$(CStr(wsSetup.Cells(R_OPT_AUTO, 2).Value))) = "YES" Then
        ' Never nag on startup - if the folder is not known, just skip.
        If Len(SchedulesFolder(False)) > 0 Then RefreshScheduleList
    End If
End Sub


' ===========================================================================
' Checks
' ===========================================================================
Private Function CheckWorkbook(ByVal wb As Workbook, ByVal fileName As String, _
                               ByVal mpiName As String, ByVal mpiNo As String, _
                               ByVal mpiClient As String) As String
    Dim out As String
    Dim v As String
    Dim links As Variant, i As Long
    Dim tail As String, schedName As String

    If GetSheet(wb, SH_META) Is Nothing Then
        CheckWorkbook = "NOT SET UP - no Metadata sheet"
        Exit Function
    End If

    v = ReadMeta(wb, "Project Name")
    If Len(mpiName) > 0 And StrComp(v, mpiName, vbTextCompare) <> 0 Then _
        out = out & "Project Name is '" & v & "', not '" & mpiName & "'. "

    v = ReadMeta(wb, "Project Number")
    If Len(mpiNo) > 0 And StrComp(v, mpiNo, vbTextCompare) <> 0 Then _
        out = out & "Project Number is '" & v & "', not '" & mpiNo & "'. "

    v = ReadMeta(wb, "Client")
    If Len(mpiClient) > 0 And StrComp(v, mpiClient, vbTextCompare) <> 0 Then _
        out = out & "Client is '" & v & "', not '" & mpiClient & "'. "

    v = ReadMeta(wb, "Revision")
    If Len(v) = 0 Or InStr(v, "#") > 0 Then out = out & "Revision not resolving. "

    schedName = ReadMeta(wb, "ScheduleName")
    If Len(schedName) = 0 Then
        out = out & "Schedule name blank. "
    Else
        ' Filename convention: "<doc number> - <schedule name>.xlsx"
        tail = fileName
        If InStrRev(tail, ".") > 0 Then tail = Left$(tail, InStrRev(tail, ".") - 1)
        If InStr(tail, "-") > 0 Then tail = Mid$(tail, InStrRev(tail, "-") + 1)
        If StrComp(Squash(tail), Squash(schedName), vbTextCompare) <> 0 Then
            out = out & "File name says '" & Trim$(tail) & "', schedule says '" & schedName & "'. "
        End If
    End If

    On Error Resume Next
    links = wb.LinkSources(xlExcelLinks)
    On Error GoTo 0
    If Not IsEmpty(links) Then
        For i = LBound(links) To UBound(links)
            If StrComp(BaseName(CStr(links(i))), ThisWorkbook.Name, vbTextCompare) <> 0 Then
                out = out & "Links to " & BaseName(CStr(links(i))) & ". "
            End If
        Next i
    End If

    If Len(out) = 0 Then out = "OK"
    CheckWorkbook = Trim$(out)
End Function


' Flags any schedule whose revision or date differs from most of the others.
' This is the "title block says P03 while everything else says P04" check.
Private Sub FlagOddOnesOut(ByVal wsList As Worksheet, ByVal lastRow As Long)
    Dim modeRev As String, modeDate As String
    Dim r As Long

    If lastRow < 3 Then Exit Sub

    modeRev = MostCommon(wsList, 8, lastRow)
    modeDate = MostCommon(wsList, 9, lastRow)

    For r = 2 To lastRow
        If Len(modeRev) > 0 Then
            If StrComp(Trim$(CStr(wsList.Cells(r, 8).Value)), modeRev, vbTextCompare) <> 0 Then
                Append wsList.Cells(r, C_LAST), "Revision differs from most schedules (" & modeRev & "). "
            End If
        End If
        If Len(modeDate) > 0 Then
            If StrComp(Trim$(CStr(wsList.Cells(r, 9).Value)), modeDate, vbTextCompare) <> 0 Then
                Append wsList.Cells(r, C_LAST), "Date differs from most schedules (" & modeDate & "). "
            End If
        End If
    Next r
End Sub


Private Function MostCommon(ByVal ws As Worksheet, ByVal col As Long, ByVal lastRow As Long) As String
    Dim d As Object, r As Long, v As String
    Dim best As String, bestN As Long
    Dim k As Variant

    Set d = CreateObject("Scripting.Dictionary")
    For r = 2 To lastRow
        v = Trim$(CStr(ws.Cells(r, col).Value))
        If Len(v) > 0 Then d(v) = d(v) + 1
    Next r
    For Each k In d.Keys
        If d(k) > bestN Then
            bestN = d(k)
            best = CStr(k)
        End If
    Next k
    ' Only meaningful if it really is the majority.
    If bestN * 2 > (lastRow - 1) Then MostCommon = best
End Function


Private Sub Append(ByVal cell As Range, ByVal txt As String)
    Dim cur As String
    cur = Trim$(CStr(cell.Value))
    If cur = "OK" Then cur = ""
    cell.Value = Trim$(cur & " " & txt)
End Sub


Private Sub ColourIssues(ByVal wsList As Worksheet, ByVal lastRow As Long)
    Dim r As Long
    For r = 2 To lastRow
        If Trim$(CStr(wsList.Cells(r, C_LAST).Value)) = "OK" Then
            wsList.Cells(r, C_LAST).Interior.Color = RGB(226, 244, 226)
        Else
            wsList.Range(wsList.Cells(r, 1), wsList.Cells(r, C_LAST)).Interior.Color = RGB(255, 235, 200)
            wsList.Cells(r, C_LAST).Interior.Color = RGB(255, 205, 205)
        End If
    Next r
End Sub


' ===========================================================================
' Plumbing
' ===========================================================================
' Where the schedules live, as a path the file system can actually read.
'
' Order: the Setup sheet override, then this workbook's own folder. If that
' is a SharePoint / OneDrive / Filery URL it is resolved to the local synced
' folder; if it cannot be resolved the user picks it once and the answer is
' written back to the Setup sheet. Returns "" if the user cancels.
Private Function SchedulesFolder(Optional ByVal askIfUnknown As Boolean = True) As String
    Dim wsSetup As Worksheet
    Dim v As String, resolved As String

    Set wsSetup = GetSheet(ThisWorkbook, SH_SETUP)

    If Not wsSetup Is Nothing Then
        v = Trim$(CStr(wsSetup.Cells(R_OPT_FOLDER, 2).Value))
        If Len(v) > 0 Then
            If FolderExists(v) Then
                SchedulesFolder = EndSep(v)
                Exit Function
            End If
            resolved = ResolveLocalFolder(v)
            If Len(resolved) > 0 Then
                SchedulesFolder = EndSep(resolved)
                Exit Function
            End If
        End If
    End If

    v = ThisWorkbook.Path
    If FolderExists(v) Then
        SchedulesFolder = EndSep(v)
        Exit Function
    End If

    resolved = ResolveLocalFolder(v)
    If Len(resolved) > 0 Then
        If Not wsSetup Is Nothing Then wsSetup.Cells(R_OPT_FOLDER, 2).Value = resolved
        SchedulesFolder = EndSep(resolved)
        Exit Function
    End If

    If Not askIfUnknown Then Exit Function

    MsgBox "This file is open from a location Excel cannot browse:" & vbCrLf & vbCrLf & _
           IIf(Len(v) > 0, v, "(not saved yet)") & vbCrLf & vbCrLf & _
           "That happens when it is opened straight from Filery, SharePoint or a " & _
           "browser rather than from the synced folder on this PC." & vbCrLf & vbCrLf & _
           "Pick the folder holding the schedules. It will be remembered on the " & _
           "Setup sheet.", vbInformation, "Where are the schedules?"

    resolved = PickFolder("Folder containing the schedules")
    If Len(resolved) = 0 Then Exit Function

    If Not wsSetup Is Nothing Then wsSetup.Cells(R_OPT_FOLDER, 2).Value = resolved
    SchedulesFolder = EndSep(resolved)
End Function


' Every workbook in the folder except this one.
Private Function ScheduleFiles(ByVal folderPath As String) As Collection
    Dim all As Collection, c As New Collection
    Dim i As Long

    Set all = FolderWorkbooks(folderPath)
    For i = 1 To all.Count
        If StrComp(all(i), ThisWorkbook.Name, vbTextCompare) <> 0 Then c.Add all(i)
    Next i

    Set ScheduleFiles = c
End Function


' Warns once when this workbook is open from a URL but the schedules are being
' written from a local folder, because Excel then stores the link to this file
' as a full URL instead of just its name.
Private Function PathFormWarning(ByVal folderPath As String) As String
    If Not IsUrlPath(ThisWorkbook.Path) Then Exit Function
    PathFormWarning = _
        "Note: this file is open from a URL (" & ThisWorkbook.Path & ") while the " & _
        "schedules are in " & folderPath & "." & vbCrLf & _
        "The links will be written as full URLs rather than as a plain file name. " & _
        "They work, but to keep them relative, open MAINPROJECTINFO from the synced " & _
        "folder in File Explorer and run this again." & vbCrLf & vbCrLf
End Function


' Locates the value cells next to Client / Project Name / Project Number on
' the Setup sheet, so inserting a row up here cannot silently repoint the
' whole project at the wrong cell.
Private Function SetupRefs(ByVal wsSetup As Worksheet, ByRef projNameRef As String, _
                           ByRef projNoRef As String, ByRef clientRef As String) As Boolean
    Dim a As Range, b As Range, c As Range

    Set a = FindLabel(wsSetup, "Project Name", 1, 6)
    Set b = FindLabel(wsSetup, "Project Number", 1, 6)
    Set c = FindLabel(wsSetup, "Client", 1, 6)

    If a Is Nothing Or b Is Nothing Or c Is Nothing Then
        MsgBox "The Setup sheet must have 'Client', 'Project Name' and " & _
               "'Project Number' in column A, in the first 6 rows." & vbCrLf & _
               "Run InstallTool to rebuild it.", vbExclamation
        Exit Function
    End If

    projNameRef = AbsRef(a.Offset(0, 1))
    projNoRef = AbsRef(b.Offset(0, 1))
    clientRef = AbsRef(c.Offset(0, 1))
    SetupRefs = True
End Function


' The suitability list pushed into every schedule: whatever is typed in
' column F of the Setup sheet, or if that is empty, the union of what the
' schedules already use (so nothing is invented).
Private Function GatherStatuses(ByVal wsSetup As Worksheet, ByVal files As Collection, _
                                ByVal folderPath As String) As Variant
    Dim bag As Object
    Dim r As Long, lastRow As Long
    Dim v As String
    Dim i As Long
    Dim wbTgt As Workbook
    Dim out() As String
    Dim k As Variant, n As Long

    Set bag = CreateObject("Scripting.Dictionary")

    lastRow = wsSetup.Cells(wsSetup.Rows.Count, 6).End(xlUp).Row
    For r = 2 To lastRow
        v = Trim$(CStr(wsSetup.Cells(r, 6).Value))
        If Len(v) > 0 Then bag(v) = True
    Next r

    If bag.Count = 0 Then
        BeginQuiet
        For i = 1 To files.Count
            Set wbTgt = Nothing
            On Error Resume Next
            Set wbTgt = Workbooks.Open(fileName:=EndSep(folderPath) & files(i), _
                                       ReadOnly:=True, UpdateLinks:=0)
            On Error GoTo 0
            If Not wbTgt Is Nothing Then
                CollectStatuses wbTgt, bag
                wbTgt.Close SaveChanges:=False
            End If
        Next i
        EndQuiet

        ' Write them back so the list becomes editable in one place.
        wsSetup.Range("F1").Value = "Suitability Codes"
        r = 2
        For Each k In bag.Keys
            wsSetup.Cells(r, 6).Value = k
            r = r + 1
        Next k
        wsSetup.Columns("F").AutoFit
    End If

    If bag.Count = 0 Then
        GatherStatuses = Empty
        Exit Function
    End If

    ReDim out(0 To bag.Count - 1)
    For Each k In bag.Keys
        out(n) = CStr(k)
        n = n + 1
    Next k
    GatherStatuses = out
End Function


Private Function EnsureSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    Set ws = GetSheet(ThisWorkbook, sheetName)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    End If
    Set EnsureSheet = ws
End Function


Private Sub BuildSetupSheet(ByVal ws As Worksheet)
    ws.Range("A1").Value = "Client"
    ws.Range("A3").Value = "Project Name"
    ws.Range("A4").Value = "Project Number"

    ws.Cells(6, 1).Value = "Options"
    ws.Cells(R_OPT_FOLDER, 1).Value = "Schedules folder"
    ws.Cells(R_OPT_BACKUP, 1).Value = "Backup before changes"
    ws.Cells(R_OPT_AUTO, 1).Value = "Refresh list on open"
    ws.Cells(R_OPT_SETUP, 1).Value = "Last setup run"
    ws.Cells(R_OPT_LIST, 1).Value = "Last list refresh"

    If Len(Trim$(CStr(ws.Cells(R_OPT_BACKUP, 2).Value))) = 0 Then ws.Cells(R_OPT_BACKUP, 2).Value = "Yes"
    If Len(Trim$(CStr(ws.Cells(R_OPT_AUTO, 2).Value))) = 0 Then ws.Cells(R_OPT_AUTO, 2).Value = "No"
    ws.Cells(R_OPT_FOLDER, 3).Value = "(blank = the folder this file is saved in)"

    ws.Cells(13, 1).Value = "New revision (used by 'Add revision to ticked')"
    ws.Cells(R_REV_FIRST + 0, 1).Value = "Revision"
    ws.Cells(R_REV_FIRST + 1, 1).Value = "Status"
    ws.Cells(R_REV_FIRST + 2, 1).Value = "Date"
    ws.Cells(R_REV_FIRST + 3, 1).Value = "Prepared by"
    ws.Cells(R_REV_FIRST + 4, 1).Value = "Checked by"
    ws.Cells(R_REV_FIRST + 5, 1).Value = "Approved by"
    ws.Cells(R_REV_FIRST + 6, 1).Value = "Description"
    ws.Cells(R_REV_FIRST + 2, 2).NumberFormat = "dd/mm/yyyy"

    If Len(Trim$(CStr(ws.Range("F1").Value))) = 0 Then ws.Range("F1").Value = "Suitability Codes"

    ws.Range("A1,A3,A4,A6,A13").Font.Bold = True
    ws.Range("F1").Font.Bold = True
    ws.Columns("A").AutoFit
    If ws.Columns("B").ColumnWidth < 30 Then ws.Columns("B").ColumnWidth = 30
End Sub


Private Sub BuildListHeaders(ByVal ws As Worksheet)
    Dim h As Variant
    Dim i As Long

    ' Keep it a plain range - simpler to clear and rewrite than a table.
    For i = ws.ListObjects.Count To 1 Step -1
        ws.ListObjects(i).Unlist
    Next i

    h = Array("Add?", "FileName", "ScheduleName", "ProjectName", "ProjectNo", _
              "Client", "DocType", "Revision", "Date", "PrBy", "ChBy", "ApBy", _
              "DocumentNo", "SuitabilitySt", "SuitabilityDs", "Checks")

    For i = 0 To UBound(h)
        ws.Cells(1, i + 1).Value = h(i)
    Next i
    ws.Rows(1).Font.Bold = True
    ws.Rows(1).Interior.Color = RGB(230, 230, 230)

    On Error Resume Next
    If Not ws.AutoFilterMode Then ws.Range("A1:P1").AutoFilter
    On Error GoTo 0
End Sub


Private Sub BuildButtons(ByVal ws As Worksheet)
    Dim b As Object
    Dim i As Long

    For i = ws.Buttons.Count To 1 Step -1
        ws.Buttons(i).Delete
    Next i

    Set b = ws.Buttons.Add(ws.Range("H2").Left, ws.Range("H2").Top, 200, 30)
    b.OnAction = "SetupProject"
    b.Caption = "Set up / repair schedules"

    Set b = ws.Buttons.Add(ws.Range("H2").Left, ws.Range("H2").Top + 36, 200, 30)
    b.OnAction = "RefreshScheduleList"
    b.Caption = "Refresh schedule list"

    Set b = ws.Buttons.Add(ws.Range("H2").Left, ws.Range("H2").Top + 72, 200, 30)
    b.OnAction = "AddRevisionToTicked"
    b.Caption = "Add revision to ticked"
End Sub


Private Sub BeginQuiet()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.AskToUpdateLinks = False
    Application.Calculation = xlCalculationAutomatic
End Sub


Private Sub EndQuiet()
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.AskToUpdateLinks = True
End Sub


' Compares text ignoring spaces, underscores and case.
Private Function Squash(ByVal s As String) As String
    s = Replace$(s, " ", "")
    s = Replace$(s, "_", "")
    Squash = LCase$(s)
End Function
