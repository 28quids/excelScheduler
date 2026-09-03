Attribute VB_Name = "modMain"
Option Explicit

' ---------------------------------------------------------------------------
' modMain - the buttons.
'
'   InstallTool           run once, after importing the modules
'   SetupProject          link/repair every schedule in the folder
'   RefreshScheduleList   read every schedule back and QA it
'   AddRevisionToTicked   append a revision line to the ticked schedules
'
' Every run reports to the Log sheet and finishes with a summary box.
' ---------------------------------------------------------------------------

' Setup sheet layout. Rows 1-4 match the original sheet so existing links to
' $B$1 / $B$3 / $B$4 keep working.
Private Const R_OPT_FOLDER  As Long = 7
Private Const R_OPT_BACKUP  As Long = 8
Private Const R_OPT_AUTO    As Long = 9
Private Const R_OPT_FULL    As Long = 10
Private Const R_OPT_HFSRC   As Long = 11
Private Const R_OPT_SETUP   As Long = 12   ' read only
Private Const R_OPT_LIST    As Long = 13   ' read only
Private Const R_REV_FIRST   As Long = 16   ' Revision..Description = 16..22

' Setup sheet palette. Yellow means "you fill this in".
Private Const CLR_INPUT     As Long = 14810111   ' RGB(255, 251, 225)
Private Const CLR_READONLY  As Long = 15921906   ' RGB(242, 242, 242)
Private Const CLR_NOTE      As Long = 8421504    ' RGB(128, 128, 128)
Private Const CLR_SECTION   As Long = 6575172    ' RGB(68, 84, 100)
Private Const CLR_LABEL     As Long = 4210752    ' RGB(64, 64, 64)

' ScheduleList columns.
Private Const C_PICK      As Long = 1
Private Const C_FILE      As Long = 2
Private Const C_DATA_1    As Long = 3      ' ScheduleName
Private Const C_CHECKS    As Long = 16
Private Const C_NEW_FIRST As Long = 17     ' New Rev..New Description = 17..23
Private Const C_STAMP     As Long = 24     ' hidden, file modified time
Private Const C_FILECHK   As Long = 25     ' hidden, checks that come from the file itself
Private Const REV_FIELDS  As Long = 7      ' Rev, Status, Date, Pr, Ch, Ap, Descr

Private mLogRow As Long


' ===========================================================================
' One-time install
' ===========================================================================
Public Sub InstallTool()
    Dim wsSetup As Worksheet
    Dim hadCodes As Long

    Set wsSetup = EnsureSheet(SH_SETUP)
    hadCodes = SuitabilityCount(wsSetup)
    BuildSetupSheet wsSetup
    BuildListHeaders EnsureSheet(SH_LIST)
    EnsureSheet(SH_LOG).Cells.Clear
    BuildButtons wsSetup

    If hadCodes > 0 Then
        If MsgBox("Replace the suitability codes in column F with the ISO 19650 " & _
                  "defaults?" & vbCrLf & vbCrLf & _
                  "Existing revision lines keep whatever they already say. This only " & _
                  "changes what the dropdown offers from now on.", _
                  vbQuestion + vbYesNo, "Suitability codes") = vbYes Then
            SeedSuitabilityCodes wsSetup
        End If
    End If

    wsSetup.Activate
    On Error Resume Next
    ActiveWindow.DisplayGridlines = False
    On Error GoTo 0

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
    Dim backupDir As String
    Dim oneLog As String
    Dim done As Long, skipped As Long, failed As Long
    Dim files As Collection, i As Long
    Dim started As Double

    Set wsSetup = GetSheet(ThisWorkbook, SH_SETUP)
    If wsSetup Is Nothing Then
        MsgBox "No Setup sheet. Run InstallTool first.", vbExclamation
        Exit Sub
    End If

    If Not SetupRefs(wsSetup, projNameRef, projNoRef, clientRef) Then Exit Sub

    If Len(Trim$(CStr(wsSetup.Range("B3").Value))) = 0 _
       Or Len(Trim$(CStr(wsSetup.Range("B4").Value))) = 0 Then
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

    statuses = GatherStatuses(wsSetup)

    If UCase$(Trim$(CStr(wsSetup.Cells(R_OPT_BACKUP, 2).Value))) <> "NO" Then
        backupDir = EndSep(folderPath) & "_backup " & Format$(Now, "yyyy-mm-dd hh-nn")
        On Error Resume Next
        MkDir backupDir
        On Error GoTo 0
    End If

    On Error GoTo Fail
    LogStart "Set up / repair schedules"
    BeginQuiet xlCalculationAutomatic
    ProgressStart files.Count, "Setting up schedules"
    started = Timer

    For i = 1 To files.Count
        fileName = files(i)
        fullPath = EndSep(folderPath) & fileName
        ProgressStep i - 1, fileName

        If Len(backupDir) > 0 Then
            On Error Resume Next
            FileCopy fullPath, EndSep(backupDir) & fileName
            On Error GoTo 0
        End If

        Set wbTgt = OpenQuiet(fullPath, False)

        If wbTgt Is Nothing Then
            failed = failed + 1
            LogLine fileName, "FAILED", "Could not open the file."
        ElseIf Not LooksLikeSchedule(wbTgt) Then
            wbTgt.Close SaveChanges:=False
            skipped = skipped + 1
            LogLine fileName, "Skipped", "No Revision Page - not a schedule."
        Else
            oneLog = ""
            On Error Resume Next
            oneLog = RepairWorkbook(wbTgt, ThisWorkbook.FullName, SH_SETUP, _
                                    projNameRef, projNoRef, clientRef, statuses)
            If Err.Number <> 0 Then
                LogLine fileName, "FAILED", "Error " & Err.Number & " - " & Err.Description
                Err.Clear
                failed = failed + 1
                wbTgt.Close SaveChanges:=False
            Else
                wbTgt.Close SaveChanges:=True
                done = done + 1
                LogLine fileName, "OK", oneLog
            End If
            On Error GoTo 0
        End If
    Next i

    ProgressDone
    wsSetup.Cells(R_OPT_SETUP, 2).Value = Format$(Now, "dd/mm/yyyy hh:nn")
    EndQuiet

    ShowSummary "Set up / repair schedules", done, skipped, failed, _
                Timer - started, IIf(Len(backupDir) > 0, "Backup: " & backupDir, "")

    RefreshScheduleList
    Exit Sub

Fail:
    Recover "Set up / repair schedules", wbTgt
End Sub


' ===========================================================================
' Button 2 - read every schedule back and check it
' ===========================================================================
Public Sub RefreshScheduleList()
    Dim wsSetup As Worksheet, wsList As Worksheet
    Dim folderPath As String, fileName As String, fullPath As String
    Dim wbTgt As Workbook
    Dim files As Collection, i As Long
    Dim r As Long
    Dim mpiName As String, mpiNo As String, mpiClient As String
    Dim keep As Object
    Dim fullRefresh As Boolean
    Dim stamp As Double
    Dim reused As Long, readCount As Long, failed As Long
    Dim started As Double

    Set wsSetup = GetSheet(ThisWorkbook, SH_SETUP)
    Set wsList = EnsureSheet(SH_LIST)
    If wsSetup Is Nothing Then Exit Sub

    folderPath = SchedulesFolder()
    If Len(folderPath) = 0 Then Exit Sub

    mpiName = Trim$(CStr(wsSetup.Range("B3").Value))
    mpiNo = Trim$(CStr(wsSetup.Range("B4").Value))
    mpiClient = Trim$(CStr(wsSetup.Range("B1").Value))
    fullRefresh = (UCase$(Trim$(CStr(wsSetup.Cells(R_OPT_FULL, 2).Value))) = "YES")

    Set files = ScheduleFiles(folderPath)
    Set keep = SnapshotList(wsList)

    On Error GoTo Fail
    LogStart "Refresh schedule list"
    ' Manual calculation: the point is to read what is SAVED in each file,
    ' which is what a recipient sees. It is also far quicker.
    BeginQuiet xlCalculationManual
    ProgressStart files.Count, "Reading schedules"
    started = Timer

    BuildListHeaders wsList
    wsList.Range(wsList.Cells(2, 1), wsList.Cells(wsList.Rows.Count, C_FILECHK)).Clear

    r = 2
    For i = 1 To files.Count
        fileName = files(i)
        fullPath = EndSep(folderPath) & fileName
        stamp = FileStamp(fullPath)
        ProgressStep i - 1, fileName

        If Not fullRefresh And CanReuse(keep, fileName, stamp) Then
            RestoreRow wsList, r, keep(LCase$(fileName))
            LogLine fileName, "Unchanged", "Not reopened - same as the last refresh."
            reused = reused + 1
            r = r + 1
        Else
            Set wbTgt = OpenQuiet(fullPath, True)
            If wbTgt Is Nothing Then
                failed = failed + 1
                LogLine fileName, "FAILED", "Could not open the file."
            Else
                If LooksLikeSchedule(wbTgt) Then
                    FillRow wsList, r, fileName, wbTgt
                    wsList.Cells(r, C_STAMP).Value = stamp
                    LogLine fileName, "Read", AsText(wsList.Cells(r, C_FILECHK).Value)
                    readCount = readCount + 1
                    r = r + 1
                Else
                    LogLine fileName, "Skipped", "No Revision Page - not a schedule."
                End If
                wbTgt.Close SaveChanges:=False
            End If
        End If
    Next i

    RestoreTypedEntries wsList, keep, r - 1
    ComposeChecks wsList, r - 1, mpiName, mpiNo, mpiClient
    FlagOddOnesOut wsList, r - 1
    ColourIssues wsList, r - 1
    FormatList wsList, r - 1

    ProgressDone
    wsSetup.Cells(R_OPT_LIST, 2).Value = Format$(Now, "dd/mm/yyyy hh:nn")
    EndQuiet

    ShowSummary "Refresh schedule list", readCount, reused, failed, Timer - started, _
                IIf(reused > 0, reused & " file(s) were unchanged since the last refresh " & _
                                "and were not reopened.", "")
    wsList.Activate
    Exit Sub

Fail:
    Recover "Refresh schedule list", wbTgt
End Sub


' ===========================================================================
' Button 3 - append a revision line to every ticked schedule
'
' Each row can carry its own new revision in the "New ..." columns. Anything
' left blank there falls back to the block on the Setup sheet, so reissuing
' all 24 on one revision and reissuing 6 of them on different ones are the
' same operation.
' ===========================================================================
Public Sub AddRevisionToTicked()
    Dim wsSetup As Worksheet, wsList As Worksheet
    Dim r As Long, lastRow As Long, n As Long, k As Long, j As Long
    Dim folderPath As String, fileName As String
    Dim wbTgt As Workbook
    Dim res As String
    Dim fld(0 To 6) As Variant
    Dim done As Long, failed As Long, skipped As Long
    Dim started As Double

    Set wsSetup = GetSheet(ThisWorkbook, SH_SETUP)
    Set wsList = GetSheet(ThisWorkbook, SH_LIST)
    If wsSetup Is Nothing Or wsList Is Nothing Then Exit Sub

    lastRow = wsList.Cells(wsList.Rows.Count, C_FILE).End(xlUp).Row
    For r = 2 To lastRow
        If IsTicked(wsList, r) Then n = n + 1
    Next r

    If n = 0 Then
        MsgBox "Nothing ticked." & vbCrLf & vbCrLf & _
               "Put an x in the 'Add?' column next to each schedule being reissued, " & _
               "and type the new revision in the 'New ...' columns on that row." & vbCrLf & vbCrLf & _
               "Anything you leave blank is taken from the 'New revision' block on " & _
               "the Setup sheet, so you can fill in the common bits once.", _
               vbExclamation, "Add revision"
        Exit Sub
    End If

    If MsgBox("Add a revision line to " & n & " schedule(s)?", _
              vbQuestion + vbYesNo, "Add revision") = vbNo Then Exit Sub

    folderPath = SchedulesFolder()
    If Len(folderPath) = 0 Then Exit Sub

    On Error GoTo Fail
    LogStart "Add revision"
    BeginQuiet xlCalculationAutomatic
    ProgressStart n, "Adding revisions"
    started = Timer

    For r = 2 To lastRow
        If IsTicked(wsList, r) Then
            fileName = AsText(wsList.Cells(r, C_FILE).Value)
            k = k + 1
            ProgressStep k - 1, fileName

            For j = 0 To REV_FIELDS - 1
                fld(j) = RowOrSetup(wsList, wsSetup, r, j)
            Next j

            If Len(AsText(fld(0))) = 0 Or Not IsDate(fld(2)) Then
                skipped = skipped + 1
                LogLine fileName, "Skipped", _
                    "Needs at least a revision and a valid date, on the row or on Setup."
            Else
                Set wbTgt = OpenQuiet(EndSep(folderPath) & fileName, False)
                If wbTgt Is Nothing Then
                    failed = failed + 1
                    LogLine fileName, "FAILED", "Could not open the file."
                Else
                    res = AppendRevision(wbTgt, AsText(fld(0)), AsText(fld(1)), CDate(fld(2)), _
                                         AsText(fld(3)), AsText(fld(4)), AsText(fld(5)), AsText(fld(6)))
                    If Len(res) = 0 Then
                        wbTgt.Close SaveChanges:=True
                        done = done + 1
                        LogLine fileName, "OK", "Added " & AsText(fld(0)) & " " & _
                                Format$(CDate(fld(2)), "dd/mm/yyyy")
                        ClearRowEntry wsList, r
                    Else
                        wbTgt.Close SaveChanges:=False
                        skipped = skipped + 1
                        LogLine fileName, "Skipped", res
                    End If
                End If
            End If
        End If
    Next r

    ProgressDone
    EndQuiet
    ShowSummary "Add revision", done, skipped, failed, Timer - started, ""
    RefreshScheduleList
    Exit Sub

Fail:
    Recover "Add revision", wbTgt
End Sub


' ===========================================================================
' Button 4 - copy headers and footers out of one workbook
'
' This is how the security classification banner (OFFICIAL, OFFICIAL-SENSITIVE,
' CONFIDENTIAL, or none) gets applied consistently: set one schedule up by hand
' under Page Layout, then push it to the rest.
' ===========================================================================
Public Sub CopyHeadersFooters()
    Dim wsSetup As Worksheet
    Dim folderPath As String, srcPath As String, fileName As String
    Dim wbSrc As Workbook, wbTgt As Workbook
    Dim files As Collection, i As Long
    Dim backupDir As String
    Dim done As Long, skipped As Long, failed As Long
    Dim started As Double
    Dim oneLog As String, gWarn As String

    Set wsSetup = GetSheet(ThisWorkbook, SH_SETUP)
    If wsSetup Is Nothing Then
        MsgBox "No Setup sheet. Run InstallTool first.", vbExclamation
        Exit Sub
    End If

    folderPath = SchedulesFolder()
    If Len(folderPath) = 0 Then Exit Sub

    srcPath = HeaderSourcePath(wsSetup, folderPath)
    If Len(srcPath) = 0 Then Exit Sub

    Set files = ScheduleFiles(folderPath)
    If files.Count = 0 Then
        MsgBox "No other Excel files found in:" & vbCrLf & folderPath, vbInformation
        Exit Sub
    End If

    If MsgBox("Copy the headers and footers from" & vbCrLf & vbCrLf & _
              BaseName(srcPath) & vbCrLf & vbCrLf & _
              "into the other workbooks in this folder?" & vbCrLf & vbCrLf & _
              "Front Cover, Revision Page and the schedule sheets are matched up " & _
              "separately. Whatever those sheets currently have is replaced.", _
              vbQuestion + vbYesNo, "Copy headers & footers") = vbNo Then Exit Sub

    If UCase$(Trim$(CStr(wsSetup.Cells(R_OPT_BACKUP, 2).Value))) <> "NO" Then
        backupDir = EndSep(folderPath) & "_backup " & Format$(Now, "yyyy-mm-dd hh-nn")
        On Error Resume Next
        MkDir backupDir
        On Error GoTo 0
    End If

    On Error GoTo Fail
    LogStart "Copy headers & footers"
    BeginQuiet xlCalculationManual
    ProgressStart files.Count, "Copying headers and footers"
    started = Timer

    Set wbSrc = OpenQuiet(srcPath, True)
    If wbSrc Is Nothing Then
        EndQuiet
        MsgBox "Could not open the source workbook:" & vbCrLf & srcPath, vbExclamation
        Exit Sub
    End If

    LogLine BaseName(srcPath), "Source", SourceSummary(wbSrc)
    gWarn = GraphicWarning(wbSrc)

    For i = 1 To files.Count
        fileName = files(i)
        ProgressStep i - 1, fileName

        If StrComp(EndSep(folderPath) & fileName, srcPath, vbTextCompare) = 0 Then
            skipped = skipped + 1
            LogLine fileName, "Skipped", "This is the source workbook."
        Else
            If Len(backupDir) > 0 Then
                On Error Resume Next
                FileCopy EndSep(folderPath) & fileName, EndSep(backupDir) & fileName
                On Error GoTo 0
            End If

            Set wbTgt = OpenQuiet(EndSep(folderPath) & fileName, False)
            If wbTgt Is Nothing Then
                failed = failed + 1
                LogLine fileName, "FAILED", "Could not open the file."
            Else
                oneLog = OldHeaderFooter(wbTgt) & " -> " & CopyHeadersFootersTo(wbSrc, wbTgt)
                wbTgt.Close SaveChanges:=True
                done = done + 1
                LogLine fileName, "OK", oneLog
            End If
        End If
    Next i

    wbSrc.Close SaveChanges:=False
    ProgressDone
    EndQuiet

    ShowSummary "Copy headers & footers", done, skipped, failed, Timer - started, _
                gWarn & IIf(Len(backupDir) > 0, "Backup: " & backupDir, "")
    Exit Sub

Fail:
    On Error Resume Next
    If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False
    On Error GoTo 0
    Recover "Copy headers & footers", wbTgt
End Sub


' The workbook to copy from. Remembered on the Setup sheet, so re-running
' after a tweak is one click.
Private Function HeaderSourcePath(ByVal wsSetup As Worksheet, ByVal folderPath As String) As String
    Dim v As String, candidate As String

    v = Trim$(CStr(wsSetup.Cells(R_OPT_HFSRC, 2).Value))

    If Len(v) > 0 Then
        If FileExists(v) Then
            HeaderSourcePath = v
            Exit Function
        End If
        candidate = EndSep(folderPath) & v
        If FileExists(candidate) Then
            HeaderSourcePath = candidate
            Exit Function
        End If
    End If

    candidate = PickWorkbook("Pick the workbook whose headers and footers are correct", folderPath)
    If Len(candidate) = 0 Then Exit Function

    ' Store just the name when it lives in the schedules folder, so the setting
    ' survives the folder moving between Filery and local.
    If StrComp(EndSep(folderPath), EndSep(Left$(candidate, InStrRev(candidate, Application.PathSeparator))), vbTextCompare) = 0 Then
        wsSetup.Cells(R_OPT_HFSRC, 2).Value = BaseName(candidate)
    Else
        wsSetup.Cells(R_OPT_HFSRC, 2).Value = candidate
    End If

    HeaderSourcePath = candidate
End Function


Private Function SourceSummary(ByVal wbSrc As Workbook) As String
    Dim ws As Worksheet
    Dim h As HFSet
    Dim out As String

    Set ws = GetSheet(wbSrc, SH_FRONT)
    If Not ws Is Nothing Then
        h = CaptureHF(ws)
        out = "Front Cover: " & DescribeHF(h) & ". "
    End If

    Set ws = GetSheet(wbSrc, "Schedule")
    If ws Is Nothing Then Set ws = FirstScheduleSheet(wbSrc)
    If Not ws Is Nothing Then
        h = CaptureHF(ws)
        out = out & ws.Name & ": " & DescribeHF(h) & "."
    End If

    SourceSummary = out
End Function


' What the target had before, so the log is a record you can undo from.
Private Function OldHeaderFooter(ByVal wb As Workbook) As String
    Dim ws As Worksheet
    Dim h As HFSet

    Set ws = GetSheet(wb, SH_FRONT)
    If ws Is Nothing Then Set ws = FirstScheduleSheet(wb)
    If ws Is Nothing Then Exit Function

    h = CaptureHF(ws)
    OldHeaderFooter = "was " & DescribeHF(h)
End Function


' &G is a picture placeholder. The image lives inside each file and cannot be
' carried across, so say so rather than let someone find out at print time.
Private Function GraphicWarning(ByVal wbSrc As Workbook) As String
    Dim ws As Worksheet
    Dim h As HFSet

    For Each ws In wbSrc.Worksheets
        h = CaptureHF(ws)
        If UsesGraphic(h) Then
            GraphicWarning = "The source uses a header/footer image (&G). The placeholder " & _
                             "was copied, but each file keeps its own image - check one " & _
                             "print preview." & vbCrLf & vbCrLf
            Exit Function
        End If
    Next ws
End Function


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
' Reading one schedule into a row
' ===========================================================================
Private Sub FillRow(ByVal wsList As Worksheet, ByVal r As Long, ByVal fileName As String, _
                    ByVal wb As Workbook)
    wsList.Cells(r, C_FILE).Value = fileName
    wsList.Cells(r, 3).Value = ReadMeta(wb, "ScheduleName")
    wsList.Cells(r, 4).Value = ReadMeta(wb, "Project Name")
    wsList.Cells(r, 5).Value = ReadMeta(wb, "Project Number")
    wsList.Cells(r, 6).Value = ReadMeta(wb, "Client")
    wsList.Cells(r, 7).Value = ReadMeta(wb, "DocumentType")
    wsList.Cells(r, 8).Value = ReadMeta(wb, "Revision")
    wsList.Cells(r, 9).Value = ReadMeta(wb, "Date")
    wsList.Cells(r, 10).Value = ReadMeta(wb, "Prepared by")
    wsList.Cells(r, 11).Value = ReadMeta(wb, "Checked by")
    wsList.Cells(r, 12).Value = ReadMeta(wb, "Approved by")
    wsList.Cells(r, 13).Value = ReadMeta(wb, "DocumentNumber")
    wsList.Cells(r, 14).Value = ReadMeta(wb, "Suitability Status")
    wsList.Cells(r, 15).Value = ReadMeta(wb, "Suitability Description")
    wsList.Cells(r, C_FILECHK).Value = CheckFile(wb)
End Sub


' Checks are about the project being consistent. They never stop a value being
' read, and there is deliberately no check on the file name: document numbers
' and schedule names do not have to agree with what the file is called.
'
' Split in two so a row reused from the last refresh is still compared against
' the CURRENT Setup values. Only this half needs the file open.
Private Function CheckFile(ByVal wb As Workbook) As String
    Dim out As String
    Dim links As Variant, i As Long

    If GetSheet(wb, SH_META) Is Nothing Then
        CheckFile = "NOT SET UP - no Metadata sheet. "
        Exit Function
    End If

    If Len(AsText(ReadMeta(wb, "Revision"))) = 0 Then out = out & "Revision is blank. "
    If Len(AsText(ReadMeta(wb, "ScheduleName"))) = 0 Then out = out & "Schedule name is blank. "

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

    CheckFile = out
End Function


' Builds the visible Checks column for every row, from the values on the row
' plus the file-derived half. Recomputed every refresh, so changing the
' project name on Setup updates the report even for files that were reused.
Private Sub ComposeChecks(ByVal wsList As Worksheet, ByVal lastRow As Long, _
                          ByVal mpiName As String, ByVal mpiNo As String, _
                          ByVal mpiClient As String)
    Dim r As Long
    Dim out As String, v As String

    For r = 2 To lastRow
        out = AsText(wsList.Cells(r, C_FILECHK).Value)

        v = AsText(wsList.Cells(r, 4).Value)
        If Len(mpiName) > 0 And StrComp(v, mpiName, vbTextCompare) <> 0 Then _
            out = out & "Project Name is '" & v & "', not '" & mpiName & "'. "

        v = AsText(wsList.Cells(r, 5).Value)
        If Len(mpiNo) > 0 And StrComp(v, mpiNo, vbTextCompare) <> 0 Then _
            out = out & "Project Number is '" & v & "', not '" & mpiNo & "'. "

        v = AsText(wsList.Cells(r, 6).Value)
        If Len(mpiClient) > 0 And StrComp(v, mpiClient, vbTextCompare) <> 0 Then _
            out = out & "Client is '" & v & "', not '" & mpiClient & "'. "

        If Len(Trim$(out)) = 0 Then out = "OK"
        wsList.Cells(r, C_CHECKS).Value = Trim$(out)
    Next r
End Sub


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
            If StrComp(AsText(wsList.Cells(r, 8).Value), modeRev, vbTextCompare) <> 0 Then
                AppendCheck wsList.Cells(r, C_CHECKS), _
                            "Revision differs from most schedules (" & modeRev & "). "
            End If
        End If
        If Len(modeDate) > 0 Then
            If StrComp(AsText(wsList.Cells(r, 9).Value), modeDate, vbTextCompare) <> 0 Then
                AppendCheck wsList.Cells(r, C_CHECKS), _
                            "Date differs from most schedules (" & modeDate & "). "
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
        v = AsText(ws.Cells(r, col).Value)
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


Private Sub AppendCheck(ByVal cell As Range, ByVal txt As String)
    Dim cur As String
    cur = AsText(cell.Value)
    If cur = "OK" Then cur = ""
    cell.Value = Trim$(cur & " " & txt)
End Sub


Private Sub ColourIssues(ByVal wsList As Worksheet, ByVal lastRow As Long)
    Dim r As Long
    For r = 2 To lastRow
        If AsText(wsList.Cells(r, C_CHECKS).Value) = "OK" Then
            wsList.Cells(r, C_CHECKS).Interior.Color = RGB(226, 244, 226)
        Else
            wsList.Range(wsList.Cells(r, 1), wsList.Cells(r, C_CHECKS)).Interior.Color = RGB(255, 235, 200)
            wsList.Cells(r, C_CHECKS).Interior.Color = RGB(255, 205, 205)
        End If
    Next r
End Sub


' ===========================================================================
' Keeping what the user typed across a refresh
' ===========================================================================

' Everything currently on the list, keyed by lower-case file name.
Private Function SnapshotList(ByVal wsList As Worksheet) As Object
    Dim d As Object
    Dim lastRow As Long, r As Long, c As Long
    Dim row() As Variant
    Dim f As String

    Set d = CreateObject("Scripting.Dictionary")
    Set SnapshotList = d

    lastRow = wsList.Cells(wsList.Rows.Count, C_FILE).End(xlUp).Row
    If lastRow < 2 Then Exit Function

    For r = 2 To lastRow
        f = LCase$(AsText(wsList.Cells(r, C_FILE).Value))
        If Len(f) > 0 Then
            ReDim row(1 To C_FILECHK)
            For c = 1 To C_FILECHK
                row(c) = wsList.Cells(r, c).Value
            Next c
            d(f) = row
        End If
    Next r
End Function


' A cached row is reusable only when the file has not been touched since.
Private Function CanReuse(ByVal keep As Object, ByVal fileName As String, _
                          ByVal stamp As Double) As Boolean
    Dim row As Variant
    Dim old As Double

    If stamp = 0 Then Exit Function
    If Not keep.Exists(LCase$(fileName)) Then Exit Function

    row = keep(LCase$(fileName))
    If Len(AsText(row(C_DATA_1))) = 0 And Len(AsText(row(C_CHECKS))) = 0 Then Exit Function

    On Error Resume Next
    old = CDbl(row(C_STAMP))
    On Error GoTo 0

    CanReuse = (old <> 0) And (Abs(old - stamp) < 0.000001)
End Function


Private Sub RestoreRow(ByVal wsList As Worksheet, ByVal r As Long, ByVal row As Variant)
    Dim c As Long
    For c = C_FILE To C_CHECKS
        wsList.Cells(r, c).Value = row(c)
    Next c
    wsList.Cells(r, C_STAMP).Value = row(C_STAMP)
    wsList.Cells(r, C_FILECHK).Value = row(C_FILECHK)
End Sub


' Puts back the tick and any new-revision text the user had typed, matched by
' file name so it follows the row even if the order changed.
Private Sub RestoreTypedEntries(ByVal wsList As Worksheet, ByVal keep As Object, _
                                ByVal lastRow As Long)
    Dim r As Long, c As Long
    Dim f As String
    Dim row As Variant

    For r = 2 To lastRow
        f = LCase$(AsText(wsList.Cells(r, C_FILE).Value))
        If keep.Exists(f) Then
            row = keep(f)
            If Len(AsText(wsList.Cells(r, C_PICK).Value)) = 0 Then _
                wsList.Cells(r, C_PICK).Value = row(C_PICK)
            For c = C_NEW_FIRST To C_NEW_FIRST + REV_FIELDS - 1
                If Len(AsText(wsList.Cells(r, c).Value)) = 0 Then wsList.Cells(r, c).Value = row(c)
            Next c
        End If
    Next r
End Sub


Private Function IsTicked(ByVal wsList As Worksheet, ByVal r As Long) As Boolean
    IsTicked = (Len(AsText(wsList.Cells(r, C_PICK).Value)) > 0)
End Function


' Field n of the new revision: the row first, then the Setup block.
Private Function RowOrSetup(ByVal wsList As Worksheet, ByVal wsSetup As Worksheet, _
                            ByVal r As Long, ByVal n As Long) As Variant
    Dim v As Variant
    v = wsList.Cells(r, C_NEW_FIRST + n).Value
    If Len(AsText(v)) > 0 Then
        RowOrSetup = v
    Else
        RowOrSetup = wsSetup.Cells(R_REV_FIRST + n, 2).Value
    End If
End Function


Private Sub ClearRowEntry(ByVal wsList As Worksheet, ByVal r As Long)
    wsList.Cells(r, C_PICK).ClearContents
    wsList.Range(wsList.Cells(r, C_NEW_FIRST), _
                 wsList.Cells(r, C_NEW_FIRST + REV_FIELDS - 1)).ClearContents
End Sub


' ===========================================================================
' Log sheet and summary
' ===========================================================================
Private Sub LogStart(ByVal what As String)
    Dim ws As Worksheet
    Set ws = EnsureSheet(SH_LOG)
    ws.Cells.Clear
    ws.Range("A1").Value = what & " - " & Format$(Now, "dd/mm/yyyy hh:nn:ss")
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = "File"
    ws.Range("B2").Value = "Result"
    ws.Range("C2").Value = "Notes"
    ws.Range("A2:C2").Font.Bold = True
    mLogRow = 3
End Sub


Private Sub LogLine(ByVal fileName As String, ByVal result As String, ByVal notes As String)
    Dim ws As Worksheet
    Set ws = GetSheet(ThisWorkbook, SH_LOG)
    If ws Is Nothing Then Exit Sub
    If mLogRow < 3 Then mLogRow = 3

    ws.Cells(mLogRow, 1).Value = fileName
    ws.Cells(mLogRow, 2).Value = result
    ws.Cells(mLogRow, 3).Value = notes
    Select Case result
        Case "FAILED": ws.Cells(mLogRow, 2).Interior.Color = RGB(255, 205, 205)
        Case "Skipped": ws.Cells(mLogRow, 2).Interior.Color = RGB(255, 235, 200)
        Case Else: ws.Cells(mLogRow, 2).Interior.Color = RGB(226, 244, 226)
    End Select
    mLogRow = mLogRow + 1
End Sub


Private Sub ShowSummary(ByVal what As String, ByVal okCount As Long, ByVal otherCount As Long, _
                        ByVal failCount As Long, ByVal seconds As Double, ByVal extra As String)
    Dim ws As Worksheet
    Dim msg As String
    Dim icon As Long

    Set ws = GetSheet(ThisWorkbook, SH_LOG)
    If Not ws Is Nothing Then ws.Columns("A:C").AutoFit

    msg = okCount & " succeeded" & vbCrLf & _
          otherCount & " skipped / unchanged" & vbCrLf & _
          failCount & " failed" & vbCrLf & vbCrLf & _
          "Took " & Duration(seconds) & "."

    If Len(extra) > 0 Then msg = msg & vbCrLf & vbCrLf & extra
    msg = msg & vbCrLf & vbCrLf & "Line by line detail is on the Log sheet."

    icon = IIf(failCount > 0, vbExclamation, vbInformation)
    MsgBox msg, icon, what
End Sub


' ===========================================================================
' Plumbing
' ===========================================================================

' Opens a workbook without link prompts, and re-asserts the quiet settings
' afterwards because opening a file can turn screen updating back on.
Private Function OpenQuiet(ByVal fullPath As String, ByVal readOnlyMode As Boolean) As Workbook
    Dim wb As Workbook
    On Error Resume Next
    Set wb = Workbooks.Open(fileName:=fullPath, ReadOnly:=readOnlyMode, UpdateLinks:=0)
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
    Application.ScreenUpdating = False
    Set OpenQuiet = wb
End Function


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
' schedules already use (so nothing is invented). Only the first run pays for
' the extra pass; after that column F is filled in.
Private Function GatherStatuses(ByVal wsSetup As Worksheet) As Variant
    Dim bag As Object
    Dim r As Long, lastRow As Long
    Dim v As String
    Dim out() As String
    Dim k As Variant, n As Long

    Set bag = CreateObject("Scripting.Dictionary")

    lastRow = wsSetup.Cells(wsSetup.Rows.Count, 6).End(xlUp).Row
    For r = 2 To lastRow
        v = Trim$(CStr(wsSetup.Cells(r, 6).Value))
        If Len(v) > 0 Then bag(v) = True
    Next r

    If bag.Count = 0 Then
        SeedSuitabilityCodes wsSetup
        lastRow = wsSetup.Cells(wsSetup.Rows.Count, 6).End(xlUp).Row
        For r = 2 To lastRow
            v = Trim$(CStr(wsSetup.Cells(r, 6).Value))
            If Len(v) > 0 Then bag(v) = True
        Next r
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
    ' Start from a clean slate so re-running does not stack formatting up.
    ws.Range("A1:F40").ClearFormats
    ws.Range("H1:H22").ClearFormats
    ws.Range("A1:A" & R_REV_FIRST + 6).Font.Color = CLR_LABEL

    ' Thin spacer rows, so each block reads as a block instead of a gap.
    ws.Rows(2).RowHeight = 6
    ws.Rows(5).RowHeight = 6
    ws.Rows(14).RowHeight = 6

    ' --- Project -----------------------------------------------------------
    ' Rows 1, 3 and 4 are fixed. Schedules set up before now link to $B$1,
    ' $B$3 and $B$4, so moving them would silently repoint every one of them
    ' until the repair was run again. Not worth it for tidiness.
    ws.Range("A1").Value = "Client"
    ws.Range("A3").Value = "Project Name"
    ws.Range("A4").Value = "Project Number"
    InputCell ws.Range("B1")
    InputCell ws.Range("B3")
    InputCell ws.Range("B4")

    ' --- Options -----------------------------------------------------------
    SectionHeader ws, 6, "OPTIONS"
    ws.Cells(R_OPT_FOLDER, 1).Value = "Schedules folder"
    ws.Cells(R_OPT_BACKUP, 1).Value = "Backup before changes"
    ws.Cells(R_OPT_AUTO, 1).Value = "Refresh list on open"
    ws.Cells(R_OPT_FULL, 1).Value = "Full refresh every time"
    ws.Cells(R_OPT_HFSRC, 1).Value = "Header/footer source"
    ws.Cells(R_OPT_SETUP, 1).Value = "Last setup run"
    ws.Cells(R_OPT_LIST, 1).Value = "Last list refresh"

    InputCell ws.Cells(R_OPT_FOLDER, 2)
    InputCell ws.Cells(R_OPT_BACKUP, 2)
    InputCell ws.Cells(R_OPT_AUTO, 2)
    InputCell ws.Cells(R_OPT_FULL, 2)
    InputCell ws.Cells(R_OPT_HFSRC, 2)
    ReadOnlyCell ws.Cells(R_OPT_SETUP, 2)
    ReadOnlyCell ws.Cells(R_OPT_LIST, 2)

    If Len(Trim$(CStr(ws.Cells(R_OPT_BACKUP, 2).Value))) = 0 Then ws.Cells(R_OPT_BACKUP, 2).Value = "Yes"
    If Len(Trim$(CStr(ws.Cells(R_OPT_AUTO, 2).Value))) = 0 Then ws.Cells(R_OPT_AUTO, 2).Value = "No"
    If Len(Trim$(CStr(ws.Cells(R_OPT_FULL, 2).Value))) = 0 Then ws.Cells(R_OPT_FULL, 2).Value = "No"

    YesNoList ws.Cells(R_OPT_BACKUP, 2)
    YesNoList ws.Cells(R_OPT_AUTO, 2)
    YesNoList ws.Cells(R_OPT_FULL, 2)

    Note ws.Cells(R_OPT_FOLDER, 3), "blank = the folder this file is saved in"
    Note ws.Cells(R_OPT_BACKUP, 3), "copies every file into a timestamped folder first"
    Note ws.Cells(R_OPT_AUTO, 3), "reads every schedule when this file is opened"
    Note ws.Cells(R_OPT_FULL, 3), "No = only reopen files that changed since last time"
    Note ws.Cells(R_OPT_HFSRC, 3), "the workbook to copy headers/footers from; blank = ask"

    ' --- New revision ------------------------------------------------------
    SectionHeader ws, R_REV_FIRST - 1, "NEW REVISION"
    Note ws.Cells(R_REV_FIRST - 1, 3), "used for blank 'New ...' cells on ScheduleList"
    ws.Cells(R_REV_FIRST + 0, 1).Value = "Revision"
    ws.Cells(R_REV_FIRST + 1, 1).Value = "Status"
    ws.Cells(R_REV_FIRST + 2, 1).Value = "Date"
    ws.Cells(R_REV_FIRST + 3, 1).Value = "Prepared by"
    ws.Cells(R_REV_FIRST + 4, 1).Value = "Checked by"
    ws.Cells(R_REV_FIRST + 5, 1).Value = "Approved by"
    ws.Cells(R_REV_FIRST + 6, 1).Value = "Description"
    InputCell ws.Range(ws.Cells(R_REV_FIRST, 2), ws.Cells(R_REV_FIRST + 6, 2))
    ws.Cells(R_REV_FIRST + 2, 2).NumberFormat = "dd/mm/yyyy"

    ' --- Suitability codes -------------------------------------------------
    SectionHeader ws, 1, "SUITABILITY CODES", 6
    If SuitabilityCount(ws) = 0 Then SeedSuitabilityCodes ws
    If SuitabilityCount(ws) > 0 Then _
        InputCell ws.Range(ws.Cells(2, 6), ws.Cells(SuitabilityCount(ws) + 1, 6))
    SuitabilityList ws, ws.Cells(R_REV_FIRST + 1, 2)

    ' --- Notes beside the buttons -----------------------------------------
    Note ws.Range("H11"), "Cells shaded yellow are the ones you fill in."
    Note ws.Range("H12"), "Progress is shown in the status bar, bottom-left of the Excel window."
    Note ws.Range("H13"), "Every run writes a line per file to the Log sheet, then a summary."
    Note ws.Range("H15"), "Added a schedule? Press 'Set up / repair schedules' again - it is safe to re-run."
    Note ws.Range("H16"), "To reissue: on ScheduleList put an x in 'Add?', fill the blue 'New ...' columns,"
    Note ws.Range("H17"), "then press 'Add revision to ticked'. Blanks fall back to the block above."
    Note ws.Range("H19"), "Security classification: set the header/footer on one workbook by hand under"
    Note ws.Range("H20"), "Page Layout, then press 'Copy headers && footers' to push it to the rest."

    ' --- Layout ------------------------------------------------------------
    ws.Columns("A").ColumnWidth = 24
    ws.Columns("B").ColumnWidth = 34
    ws.Columns("C").ColumnWidth = 46
    ws.Columns("D:E").ColumnWidth = 3
    ws.Columns("F").ColumnWidth = 34
    ws.Columns("G").ColumnWidth = 3
    ws.Range("A1:F40").VerticalAlignment = xlCenter
End Sub


' A section title: bold, coloured, with a rule underneath.
Private Sub SectionHeader(ByVal ws As Worksheet, ByVal r As Long, ByVal txt As String, _
                          Optional ByVal firstCol As Long = 1, Optional ByVal lastCol As Long = 3)
    Dim rng As Range

    If firstCol > 1 Then lastCol = firstCol
    ws.Cells(r, firstCol).Value = txt
    Set rng = ws.Range(ws.Cells(r, firstCol), ws.Cells(r, lastCol))

    With rng.Font
        .Bold = True
        .Color = CLR_SECTION
        .Size = 10
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(180, 190, 205)
    End With
End Sub


' A cell the user is meant to type in.
Private Sub InputCell(ByVal rng As Range)
    rng.Interior.Color = CLR_INPUT
    With rng.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(214, 206, 160)
    End With
End Sub


' A cell the tool writes and the user should leave alone.
Private Sub ReadOnlyCell(ByVal rng As Range)
    rng.Interior.Color = CLR_READONLY
    rng.Font.Italic = True
    rng.Font.Color = CLR_NOTE
End Sub


Private Sub Note(ByVal rng As Range, ByVal txt As String)
    rng.Value = txt
    rng.Font.Italic = True
    rng.Font.Color = CLR_NOTE
    rng.Font.Size = 9
End Sub


Private Sub YesNoList(ByVal rng As Range)
    On Error Resume Next
    rng.Validation.Delete
    rng.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:="Yes,No"
    rng.Validation.InCellDropdown = True
    On Error GoTo 0
End Sub


Private Sub SuitabilityList(ByVal wsSetup As Worksheet, ByVal rng As Range)
    Dim n As Long
    n = SuitabilityCount(wsSetup)
    If n < 1 Then Exit Sub

    On Error Resume Next
    rng.Validation.Delete
    rng.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:="=$F$2:$F$" & (n + 1)
    rng.Validation.IgnoreBlank = True
    rng.Validation.InCellDropdown = True
    On Error GoTo 0
End Sub


Private Function SuitabilityCount(ByVal wsSetup As Worksheet) As Long
    Dim lastRow As Long, r As Long
    lastRow = wsSetup.Cells(wsSetup.Rows.Count, 6).End(xlUp).Row
    For r = 2 To lastRow
        If Len(Trim$(CStr(wsSetup.Cells(r, 6).Value))) > 0 Then _
            SuitabilityCount = SuitabilityCount + 1
    Next r
End Function


Private Sub SeedSuitabilityCodes(ByVal wsSetup As Worksheet)
    Dim codes As Variant
    Dim i As Long

    codes = DefaultSuitabilityCodes()
    wsSetup.Range("F2:F200").ClearContents
    For i = LBound(codes) To UBound(codes)
        wsSetup.Cells(2 + i - LBound(codes), 6).Value = codes(i)
    Next i
    wsSetup.Columns("F").AutoFit
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
              "DocumentNo", "SuitabilitySt", "SuitabilityDs", "Checks", _
              "New Rev", "New Status", "New Date", "New PrBy", "New ChBy", _
              "New ApBy", "New Description", "_Stamp", "_FileChecks")

    For i = 0 To UBound(h)
        ws.Cells(1, i + 1).Value = h(i)
    Next i

    ws.Rows(1).Font.Bold = True
    ws.Range(ws.Cells(1, 1), ws.Cells(1, C_CHECKS)).Interior.Color = RGB(230, 230, 230)
    ws.Range(ws.Cells(1, C_NEW_FIRST), ws.Cells(1, C_NEW_FIRST + REV_FIELDS - 1)). _
        Interior.Color = RGB(214, 232, 255)
    ws.Columns(C_STAMP).Hidden = True
    ws.Columns(C_FILECHK).Hidden = True

    HeaderNote ws.Cells(1, C_PICK), _
        "Put an x here on every schedule you are reissuing, then press " & _
        "'Add revision to ticked' on the Setup sheet." & vbCrLf & vbCrLf & _
        "The revision itself comes from the blue 'New ...' columns on the same row. " & _
        "Anything you leave blank there is taken from the 'New revision' block on " & _
        "the Setup sheet, so common values only get typed once."
    HeaderNote ws.Cells(1, C_NEW_FIRST), _
        "The revision line to add to this schedule." & vbCrLf & vbCrLf & _
        "Fill in as much or as little as you like: blank cells fall back to the " & _
        "'New revision' block on the Setup sheet. Rows without an x in 'Add?' are " & _
        "ignored." & vbCrLf & vbCrLf & _
        "The new line goes in directly under the last revision in the table."

    On Error Resume Next
    If Not ws.AutoFilterMode Then ws.Range(ws.Cells(1, 1), ws.Cells(1, C_CHECKS)).AutoFilter
    On Error GoTo 0
End Sub


Private Sub FormatList(ByVal ws As Worksheet, ByVal lastRow As Long)
    Dim wsSetup As Worksheet
    Dim codes As Long
    Dim rng As Range

    ws.Columns(9).NumberFormat = "dd/mm/yyyy"
    ws.Columns(C_NEW_FIRST + 2).NumberFormat = "dd/mm/yyyy"
    ' Clearing the rows resets these to General every refresh, so set it here.
    ws.Columns(C_STAMP).NumberFormat = "dd/mm/yyyy hh:mm"

    ws.Range(ws.Cells(1, 1), ws.Cells(1, C_CHECKS)).EntireColumn.AutoFit
    If ws.Columns(C_CHECKS).ColumnWidth > 60 Then ws.Columns(C_CHECKS).ColumnWidth = 60
    ws.Columns(C_STAMP).Hidden = True
    ws.Columns(C_FILECHK).Hidden = True

    If lastRow < 2 Then Exit Sub

    ' Dropdown of suitability codes on the New Status column.
    Set wsSetup = GetSheet(ThisWorkbook, SH_SETUP)
    If wsSetup Is Nothing Then Exit Sub
    codes = wsSetup.Cells(wsSetup.Rows.Count, 6).End(xlUp).Row
    If codes < 2 Then Exit Sub

    Set rng = ws.Range(ws.Cells(2, C_NEW_FIRST + 1), ws.Cells(lastRow, C_NEW_FIRST + 1))
    On Error Resume Next
    rng.Validation.Delete
    rng.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:="=" & SheetRef(SH_SETUP) & "!$F$2:$F$" & codes
    rng.Validation.IgnoreBlank = True
    rng.Validation.InCellDropdown = True
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

    Set b = ws.Buttons.Add(ws.Range("H2").Left, ws.Range("H2").Top + 108, 200, 30)
    b.OnAction = "CopyHeadersFooters"
    b.Caption = "Copy headers && footers"
End Sub


' Always puts Excel back the way it was. Without this an unexpected error
' would leave screen updating off and the whole application looking frozen.
Private Sub Recover(ByVal what As String, ByRef wbTgt As Workbook)
    Dim n As Long, d As String

    n = Err.Number
    d = Err.Description

    On Error Resume Next
    If Not wbTgt Is Nothing Then wbTgt.Close SaveChanges:=False
    ProgressDone
    EndQuiet
    On Error GoTo 0

    MsgBox what & " stopped." & vbCrLf & vbCrLf & _
           "Error " & n & ": " & d & vbCrLf & vbCrLf & _
           "Nothing else was changed. What had been done up to that point is " & _
           "on the Log sheet.", vbExclamation, what
End Sub


Private Sub BeginQuiet(ByVal calcMode As XlCalculation)
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.AskToUpdateLinks = False
    Application.Calculation = calcMode
End Sub


Private Sub EndQuiet()
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.AskToUpdateLinks = True
    Application.StatusBar = False
End Sub
