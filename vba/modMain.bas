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
Private Const R_OPT_HFIMG   As Long = 12
Private Const R_OPT_HFSCALE As Long = 13
Private Const R_OPT_SETUP   As Long = 14   ' read only
Private Const R_OPT_LIST    As Long = 15   ' read only
Private Const R_REV_FIRST   As Long = 18   ' Revision..Description = 18..24
Private Const R_FLD_FIRST   As Long = 27   ' extra project fields
Private Const R_FLD_COUNT   As Long = 12

' Option rows are WRITTEN at the constants above but READ by these labels.
' The rows have moved every time an option was added, and each move left the
' previous value sitting under a new meaning. Reading by label means an older
' sheet still reads correctly, and InstallTool tidies the layout.
Private Const LBL_FOLDER  As String = "Schedules folder"
Private Const LBL_BACKUP  As String = "Backup before changes"
Private Const LBL_AUTO    As String = "Refresh list on open"
Private Const LBL_FULL    As String = "Full refresh every time"
Private Const LBL_HFSRC   As String = "Reference schedule"
Private Const LBL_HFIMG   As String = "Header image"
Private Const LBL_HFSCALE As String = "Header image scale %"
Private Const LBL_SETUP   As String = "Last setup run"
Private Const LBL_LIST    As String = "Last list refresh"

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
Private Const C_NEWNAME   As Long = 24     ' proposed file name, for the rename
Private Const C_STAMP     As Long = 25     ' hidden, file modified time
Private Const C_FILECHK   As Long = 26     ' hidden, checks that come from the file itself
Private Const REV_FIELDS  As Long = 7      ' Rev, Status, Date, Pr, Ch, Ap, Descr

' How many runs the Log sheet keeps. Every button finishes by refreshing the
' list, so a log that held one run only ever showed the refresh.
Private Const LOG_RUNS As Long = 5

Private mLogRow As Long


' ===========================================================================
' One-time install
' ===========================================================================
Public Sub InstallTool()
    Dim wsSetup As Worksheet
    Dim wsList As Worksheet
    Dim hadCodes As Long

    Set wsSetup = EnsureSheet(SH_SETUP)
    Set wsList = EnsureSheet(SH_LIST)
    hadCodes = SuitabilityCount(wsSetup)
    BuildSetupSheet wsSetup

    ' The columns have moved as inputs were added, so anything sitting on the
    ' list belongs to the old layout and is now under the wrong heading.
    ' Emptying it costs nothing: the next refresh rebuilds it from the files.
    BuildListHeaders wsList
    wsList.Range(wsList.Cells(2, 1), wsList.Cells(wsList.Rows.Count, C_FILECHK)).Clear
    SetColumnState wsList
    EnsureSheet SH_LOG
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

    MsgBox "Ready. Schedule tool version " & TOOL_VERSION & "." & vbCrLf & vbCrLf & _
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

    If UCase$(Trim$(CStr(Opt(wsSetup, LBL_BACKUP, R_OPT_BACKUP).Value))) <> "NO" Then
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
                                    projNameRef, projNoRef, clientRef, statuses, _
                                    ProjectFields(wsSetup))
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
    Opt(wsSetup, LBL_SETUP, R_OPT_SETUP).Value = Format$(Now, "dd/mm/yyyy hh:nn")
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
    fullRefresh = (UCase$(Trim$(CStr(Opt(wsSetup, LBL_FULL, R_OPT_FULL).Value))) = "YES")

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
    Opt(wsSetup, LBL_LIST, R_OPT_LIST).Value = Format$(Now, "dd/mm/yyyy hh:nn")
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
    Dim oneLog As String
    Dim imgPath As String
    Dim imgW As Double, imgH As Double, scalePct As Double

    Set wsSetup = GetSheet(ThisWorkbook, SH_SETUP)
    If wsSetup Is Nothing Then
        MsgBox "No Setup sheet. Run InstallTool first.", vbExclamation
        Exit Sub
    End If

    folderPath = SchedulesFolder()
    If Len(folderPath) = 0 Then Exit Sub

    srcPath = HeaderSourcePath(wsSetup, folderPath)
    If Len(srcPath) = 0 Then Exit Sub

    imgPath = HeaderImagePath(wsSetup, folderPath)
    If Len(imgPath) > 0 Then
        If Not MeasureImage(imgPath, imgW, imgH) Then
            MsgBox "Could not read that image:" & vbCrLf & imgPath, vbExclamation
            Exit Sub
        End If
        scalePct = CDbl(Opt(wsSetup, LBL_HFSCALE, R_OPT_HFSCALE).Value)
        If scalePct <= 0 Then scalePct = 20
        imgW = imgW * scalePct / 100
        imgH = imgH * scalePct / 100
    End If

    Set files = ScheduleFiles(folderPath)
    If files.Count = 0 Then
        MsgBox "No other Excel files found in:" & vbCrLf & folderPath, vbInformation
        Exit Sub
    End If

    If MsgBox("Copy the headers and footers from" & vbCrLf & vbCrLf & _
              BaseName(srcPath) & vbCrLf & vbCrLf & _
              "into the other workbooks in this folder?" & vbCrLf & vbCrLf & _
              "Front Cover, Revision Page and the schedule sheets are matched up " & _
              "separately. Whatever those sheets currently have is replaced." & vbCrLf & vbCrLf & _
              IIf(Len(imgPath) > 0, "Header image: " & BaseName(imgPath) & " at " & _
                  CStr(scalePct) & "%, top right.", "No header image set."), _
              vbQuestion + vbYesNo, "Copy headers & footers") = vbNo Then Exit Sub

    If UCase$(Trim$(CStr(Opt(wsSetup, LBL_BACKUP, R_OPT_BACKUP).Value))) <> "NO" Then
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
                oneLog = CopyHeadersFootersTo(wbSrc, wbTgt, imgPath, imgW, imgH)
                wbTgt.Close SaveChanges:=True
                If InStr(1, oneLog, "PROBLEM:", vbTextCompare) > 0 Then
                    failed = failed + 1
                    LogLine fileName, "FAILED", oneLog
                Else
                    done = done + 1
                    LogLine fileName, "OK", oneLog
                End If
            End If
        End If
    Next i

    wbSrc.Close SaveChanges:=False
    ProgressDone
    EndQuiet

    ShowSummary "Copy headers & footers", done, skipped, failed, Timer - started, _
                IIf(Len(imgPath) > 0, "Header image set from " & BaseName(imgPath) & _
                    ". Check one print preview." & vbCrLf & vbCrLf, "") & _
                IIf(Len(backupDir) > 0, "Backup: " & backupDir, "")
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

    v = Trim$(CStr(Opt(wsSetup, LBL_HFSRC, R_OPT_HFSRC).Value))

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

    candidate = PickWorkbook("Pick the schedule that is set up correctly", folderPath)
    If Len(candidate) = 0 Then Exit Function

    ' Store just the name when it lives in the schedules folder, so the setting
    ' survives the folder moving between Filery and local.
    If StrComp(EndSep(folderPath), EndSep(Left$(candidate, InStrRev(candidate, Application.PathSeparator))), vbTextCompare) = 0 Then
        Opt(wsSetup, LBL_HFSRC, R_OPT_HFSRC).Value = BaseName(candidate)
    Else
        Opt(wsSetup, LBL_HFSRC, R_OPT_HFSRC).Value = candidate
    End If

    HeaderSourcePath = candidate
End Function


' The logo for the top-right of the header. Blank means the user is asked
' once; cancelling the picker means "no image", which is a valid answer.
Private Function HeaderImagePath(ByVal wsSetup As Worksheet, ByVal folderPath As String) As String
    Dim v As String, candidate As String

    v = Trim$(CStr(Opt(wsSetup, LBL_HFIMG, R_OPT_HFIMG).Value))

    If Len(v) > 0 Then
        If FileExists(v) Then
            HeaderImagePath = v
            Exit Function
        End If
        candidate = EndSep(folderPath) & v
        If FileExists(candidate) Then
            HeaderImagePath = candidate
            Exit Function
        End If
        MsgBox "The header image on the Setup sheet is not where it says:" & vbCrLf & vbCrLf & _
               v & vbCrLf & vbCrLf & "Pick it again, or cancel to carry on without one.", _
               vbExclamation, "Header image"
    Else
        If MsgBox("Put a logo in the top-right of the header?" & vbCrLf & vbCrLf & _
                  "Yes to pick the image file. No to leave headers as text only." & vbCrLf & _
                  "Whatever you choose is remembered on the Setup sheet.", _
                  vbQuestion + vbYesNo, "Header image") = vbNo Then Exit Function
    End If

    candidate = PickImage("Pick the header logo", folderPath)
    If Len(candidate) = 0 Then Exit Function

    ' Store just the name when it sits with the schedules, so the setting
    ' survives the folder moving between Filery and local.
    If StrComp(EndSep(folderPath), _
               EndSep(Left$(candidate, InStrRev(candidate, Application.PathSeparator))), _
               vbTextCompare) = 0 Then
        Opt(wsSetup, LBL_HFIMG, R_OPT_HFIMG).Value = BaseName(candidate)
    Else
        Opt(wsSetup, LBL_HFIMG, R_OPT_HFIMG).Value = candidate
    End If

    HeaderImagePath = candidate
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






' ===========================================================================
' Button 5 - copy the Front Cover and Revision Page from a reference schedule
'
' For a change to the common pages themselves: dropping a security
' classification, moving a logo, rewording the cover. Set one schedule up by
' hand, then push those two sheets to the others.
'
' Each target keeps its own schedule title, its whole revision history, and
' its document type, Delref, BSUID and trigger events. The links are rebuilt
' locally afterwards, so nothing points back at the reference.
' ===========================================================================
Public Sub CopyCommonSheets()
    Dim wsSetup As Worksheet
    Dim folderPath As String, srcPath As String, fileName As String
    Dim wbSrc As Workbook, wbTgt As Workbook
    Dim files As Collection, i As Long
    Dim backupDir As String
    Dim done As Long, skipped As Long, failed As Long
    Dim started As Double
    Dim oneLog As String
    Dim projNameRef As String, projNoRef As String, clientRef As String
    Dim statuses As Variant

    Set wsSetup = GetSheet(ThisWorkbook, SH_SETUP)
    If wsSetup Is Nothing Then
        MsgBox "No Setup sheet. Run InstallTool first.", vbExclamation
        Exit Sub
    End If

    If Not SetupRefs(wsSetup, projNameRef, projNoRef, clientRef) Then Exit Sub

    folderPath = SchedulesFolder()
    If Len(folderPath) = 0 Then Exit Sub

    srcPath = HeaderSourcePath(wsSetup, folderPath)
    If Len(srcPath) = 0 Then Exit Sub

    Set files = ChosenTargets(folderPath, srcPath)
    If files Is Nothing Then Exit Sub
    If files.Count = 0 Then
        MsgBox "Nothing to copy to.", vbInformation
        Exit Sub
    End If

    If MsgBox("Replace the Front Cover and Revision Page in " & files.Count & _
              " schedule(s) with the ones from" & vbCrLf & vbCrLf & _
              BaseName(srcPath) & vbCrLf & vbCrLf & _
              "Each schedule keeps its own:" & vbCrLf & _
              "  - schedule title" & vbCrLf & _
              "  - full revision history" & vbCrLf & _
              "  - document type, Delref, BSUID, trigger events" & vbCrLf & vbCrLf & _
              "Everything else on those two sheets is replaced, and the links " & _
              "are rebuilt to point at each schedule's own data." & vbCrLf & vbCrLf & _
              IIf(UCase$(Trim$(CStr(Opt(wsSetup, LBL_BACKUP, R_OPT_BACKUP).Value))) = "NO", _
                  "Backups are switched OFF on the Setup sheet.", _
                  "A backup is taken first."), _
              vbExclamation + vbYesNo + vbDefaultButton2, "Copy cover & revision page") = vbNo Then Exit Sub

    statuses = GatherStatuses(wsSetup)

    If UCase$(Trim$(CStr(Opt(wsSetup, LBL_BACKUP, R_OPT_BACKUP).Value))) <> "NO" Then
        backupDir = EndSep(folderPath) & "_backup " & Format$(Now, "yyyy-mm-dd hh-nn")
        On Error Resume Next
        MkDir backupDir
        On Error GoTo 0
    End If

    On Error GoTo Fail
    LogStart "Copy cover & revision page"
    BeginQuiet xlCalculationAutomatic
    ProgressStart files.Count, "Copying cover and revision page"
    started = Timer

    Set wbSrc = OpenQuiet(srcPath, True)
    If wbSrc Is Nothing Then
        EndQuiet
        MsgBox "Could not open the reference workbook:" & vbCrLf & srcPath, vbExclamation
        Exit Sub
    End If

    For i = 1 To files.Count
        fileName = files(i)
        ProgressStep i - 1, fileName

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
            ' Unlike the other buttons this one does not require a Revision
            ' Page: giving a bare schedule its common sheets is the job.
            oneLog = CopyCommonSheetsTo(wbSrc, wbTgt, fileName)

            If InStr(1, oneLog, "PROBLEM:", vbTextCompare) > 0 Then
                wbTgt.Close SaveChanges:=False
                failed = failed + 1
                LogLine fileName, "FAILED", oneLog & "Nothing was saved."
            Else
                ' Rebuild every link locally, so none point at the reference.
                oneLog = oneLog & RepairWorkbook(wbTgt, ThisWorkbook.FullName, SH_SETUP, _
                                                 projNameRef, projNoRef, clientRef, statuses, _
                                                 ProjectFields(wsSetup))
                wbTgt.Close SaveChanges:=True
                done = done + 1
                LogLine fileName, "OK", oneLog
            End If
        End If
    Next i

    wbSrc.Close SaveChanges:=False
    ProgressDone
    EndQuiet

    ShowSummary "Copy cover & revision page", done, skipped, failed, Timer - started, _
                IIf(Len(backupDir) > 0, "Backup: " & backupDir, "")
    RefreshScheduleList
    Exit Sub

Fail:
    On Error Resume Next
    If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False
    On Error GoTo 0
    Recover "Copy cover & revision page", wbTgt
End Sub


' Which schedules to act on: the ones ticked on ScheduleList, or all of them.
' Returns Nothing if the user backs out.
Private Function ChosenTargets(ByVal folderPath As String, ByVal srcPath As String) As Collection
    Dim all As Collection, picked As New Collection
    Dim wsList As Worksheet
    Dim r As Long, lastRow As Long, ticked As Long
    Dim answer As VbMsgBoxResult
    Dim f As String
    Dim i As Long

    Set all = ScheduleFiles(folderPath)

    Set wsList = GetSheet(ThisWorkbook, SH_LIST)
    If Not wsList Is Nothing Then
        lastRow = wsList.Cells(wsList.Rows.Count, C_FILE).End(xlUp).Row
        For r = 2 To lastRow
            If IsTicked(wsList, r) Then ticked = ticked + 1
        Next r
    End If

    If ticked > 0 Then
        answer = MsgBox(ticked & " schedule(s) are ticked on ScheduleList." & vbCrLf & vbCrLf & _
                        "Yes  - only those " & ticked & vbCrLf & _
                        "No   - all " & all.Count & " in the folder" & vbCrLf & _
                        "Cancel - stop", _
                        vbQuestion + vbYesNoCancel, "Which schedules?")
        If answer = vbCancel Then Exit Function
        If answer = vbYes Then
            For r = 2 To lastRow
                If IsTicked(wsList, r) Then
                    f = AsText(wsList.Cells(r, C_FILE).Value)
                    If Len(f) > 0 And StrComp(EndSep(folderPath) & f, srcPath, vbTextCompare) <> 0 Then
                        picked.Add f
                    End If
                End If
            Next r
            Set ChosenTargets = picked
            Exit Function
        End If
    End If

    For i = 1 To all.Count
        If StrComp(EndSep(folderPath) & all(i), srcPath, vbTextCompare) <> 0 Then picked.Add all(i)
    Next i
    Set ChosenTargets = picked
End Function


' ===========================================================================
' Button 6 - rename schedule files
'
' Two presses. The first fills in the New FileName column by find and replace
' so you can read what it intends to do; the second does it. Typing a name in
' that column by hand skips straight to the second.
'
' Renaming is safe for the links - schedules link to the MPI, never to each
' other, and the MPI finds them by scanning the folder. It is NOT cosmetic
' though: the document number is derived from the file name, so every renamed
' file is reopened and saved to bake the new number into its title block.
' ===========================================================================
Public Sub RenameFiles()
    Dim wsList As Worksheet
    Dim folderPath As String
    Dim proposed As Long

    Set wsList = GetSheet(ThisWorkbook, SH_LIST)
    If wsList Is Nothing Then
        MsgBox "No ScheduleList sheet. Run InstallTool first.", vbExclamation
        Exit Sub
    End If

    folderPath = SchedulesFolder()
    If Len(folderPath) = 0 Then Exit Sub

    proposed = CountProposedNames(wsList)

    If proposed = 0 Then
        ProposeNames wsList, folderPath
    Else
        ApplyRenames wsList, folderPath
    End If
End Sub


Private Function CountProposedNames(ByVal wsList As Worksheet) As Long
    Dim r As Long, lastRow As Long
    lastRow = wsList.Cells(wsList.Rows.Count, C_FILE).End(xlUp).Row
    For r = 2 To lastRow
        If Len(AsText(wsList.Cells(r, C_NEWNAME).Value)) > 0 Then _
            CountProposedNames = CountProposedNames + 1
    Next r
End Function


' First press: work out the new names and show them, change nothing.
Private Sub ProposeNames(ByVal wsList As Worksheet, ByVal folderPath As String)
    Dim findText As Variant, replaceText As Variant
    Dim r As Long, lastRow As Long, n As Long
    Dim old As String, proposed As String
    Dim onlyTicked As Boolean

    findText = Application.InputBox( _
        "Text to find in the file names." & vbCrLf & vbCrLf & _
        "For example PROJECTNUMBER, to swap the placeholder for the real one.", _
        "Rename files", Type:=2)
    If VarType(findText) = vbBoolean Then Exit Sub
    If Len(CStr(findText)) = 0 Then
        MsgBox "Nothing to find.", vbExclamation
        Exit Sub
    End If

    replaceText = Application.InputBox( _
        "Replace """ & findText & """ with what?" & vbCrLf & vbCrLf & _
        "Leave it empty to delete that text from the names.", _
        "Rename files", Type:=2)
    If VarType(replaceText) = vbBoolean Then Exit Sub

    onlyTicked = (TickedCount(wsList) > 0)
    If onlyTicked Then
        If MsgBox(TickedCount(wsList) & " schedule(s) are ticked." & vbCrLf & vbCrLf & _
                  "Yes - only those." & vbCrLf & "No  - every schedule on the list.", _
                  vbQuestion + vbYesNo, "Which schedules?") = vbNo Then onlyTicked = False
    End If

    lastRow = wsList.Cells(wsList.Rows.Count, C_FILE).End(xlUp).Row
    For r = 2 To lastRow
        If (Not onlyTicked) Or IsTicked(wsList, r) Then
            old = AsText(wsList.Cells(r, C_FILE).Value)
            If Len(old) > 0 Then
                proposed = Replace(old, CStr(findText), CStr(replaceText), 1, -1, vbTextCompare)
                If StrComp(proposed, old, vbBinaryCompare) <> 0 Then
                    wsList.Cells(r, C_NEWNAME).Value = proposed
                    n = n + 1
                End If
            End If
        End If
    Next r

    wsList.Columns(C_NEWNAME).AutoFit
    wsList.Activate

    If n = 0 Then
        MsgBox """" & findText & """ is not in any of those file names. Nothing to do.", _
               vbInformation, "Rename files"
    Else
        MsgBox n & " file name(s) proposed in the 'New FileName' column." & vbCrLf & vbCrLf & _
               "Read them, edit any you want to change, clear any you do not want, " & _
               "then press 'Rename files' again to do it.", _
               vbInformation, "Rename files - review"
    End If
End Sub


Private Function TickedCount(ByVal wsList As Worksheet) As Long
    Dim r As Long, lastRow As Long
    lastRow = wsList.Cells(wsList.Rows.Count, C_FILE).End(xlUp).Row
    For r = 2 To lastRow
        If IsTicked(wsList, r) Then TickedCount = TickedCount + 1
    Next r
End Function


' Second press: check the whole batch, then do it.
'
' Every problem is found before anything is renamed, and any one of them
' refuses the lot. A target that already exists is the exception: when the
' file holding it is itself being renamed, that is a chain or a swap and it
' is worked out rather than refused.
Private Sub ApplyRenames(ByVal wsList As Worksheet, ByVal folderPath As String)
    Dim r As Long, lastRow As Long
    Dim old As String, proposed As String
    Dim rowIdx() As Long, olds() As String, news() As String
    Dim placed() As Boolean
    Dim n As Long, i As Long, j As Long
    Dim problems As String
    Dim done As Long, failed As Long
    Dim started As Double
    Dim wb As Workbook
    Dim progress As Boolean
    Dim temp As String

    lastRow = wsList.Cells(wsList.Rows.Count, C_FILE).End(xlUp).Row
    ReDim rowIdx(1 To lastRow)
    ReDim olds(1 To lastRow)
    ReDim news(1 To lastRow)
    ReDim placed(1 To lastRow)

    For r = 2 To lastRow
        proposed = Trim$(AsText(wsList.Cells(r, C_NEWNAME).Value))
        If Len(proposed) > 0 Then
            old = AsText(wsList.Cells(r, C_FILE).Value)

            ' No extension typed means keep the one it has.
            If Len(FileExtension(proposed)) = 0 Then proposed = proposed & FileExtension(old)

            If StrComp(old, proposed, vbBinaryCompare) <> 0 Then
                n = n + 1
                rowIdx(n) = r
                olds(n) = old
                news(n) = proposed
                problems = problems & CheckRename(folderPath, old, proposed)
            End If
        End If
    Next r

    If n = 0 Then Exit Sub

    ' Two rows renaming to the same thing would destroy one of them.
    For i = 1 To n
        For j = i + 1 To n
            If StrComp(news(i), news(j), vbTextCompare) = 0 Then
                problems = problems & "Two schedules would both become '" & news(i) & "'. "
            End If
        Next j
    Next i

    ' A target that exists is only a problem when the file holding it is not
    ' itself moving. When it is, this is a chain or a swap and it is handled
    ' below rather than refused.
    For i = 1 To n
        If FileExists(EndSep(folderPath) & news(i)) Then
            If Not InBatch(olds, n, news(i)) Then
                problems = problems & "'" & news(i) & "' already exists and is not " & _
                           "being renamed itself. "
            End If
        End If
    Next i

    If Len(problems) > 0 Then
        MsgBox "Nothing was renamed." & vbCrLf & vbCrLf & problems & vbCrLf & vbCrLf & _
               "Fix the 'New FileName' column and try again.", vbExclamation, "Rename files"
        Exit Sub
    End If

    If MsgBox("Rename " & n & " file(s)?" & vbCrLf & vbCrLf & _
              olds(1) & vbCrLf & "    becomes" & vbCrLf & news(1) & _
              IIf(n > 1, vbCrLf & vbCrLf & "...and " & (n - 1) & " more.", "") & vbCrLf & vbCrLf & _
              "Each one is then reopened and saved, because the document number " & _
              "on the title block is taken from the file name.", _
              vbQuestion + vbYesNo, "Rename files") = vbNo Then Exit Sub

    On Error GoTo Fail
    LogStart "Rename files"
    BeginQuiet xlCalculationAutomatic
    ProgressStart n, "Renaming"
    started = Timer

    ' Pass 1: rename anything whose target is free, repeatedly. Each one that
    ' lands frees its old name, which may be another file's target, so a chain
    ' unwinds from the end.
    progress = True
    Do While progress
        progress = False
        For i = 1 To n
            If Not placed(i) Then
                If Not FileExists(EndSep(folderPath) & news(i)) Then
                    If RenameOne(folderPath, olds(i), news(i)) Then
                        placed(i) = True
                        progress = True
                        done = done + 1
                        LogLine olds(i), "OK", "Renamed to " & news(i)
                        ProgressStep done, news(i)
                    Else
                        placed(i) = True
                        failed = failed + 1
                        LogLine olds(i), "FAILED", "Rename failed - " & Err.Description
                    End If
                End If
            End If
        Next i
    Loop

    ' Pass 2: whatever is left is a cycle, where every target is held by
    ' another file in the batch. Park them under temporary names to break it,
    ' then pass 1's logic finishes the job.
    For i = 1 To n
        If Not placed(i) Then
            temp = FreeTempName(folderPath, olds(i))
            If RenameOne(folderPath, olds(i), temp) Then
                LogLine olds(i), "OK", "Held as " & temp & " to break a rename loop"
                olds(i) = temp
            Else
                placed(i) = True
                failed = failed + 1
                LogLine olds(i), "FAILED", "Could not park the file - " & Err.Description
            End If
        End If
    Next i

    For i = 1 To n
        If Not placed(i) Then
            If RenameOne(folderPath, olds(i), news(i)) Then
                placed(i) = True
                done = done + 1
                LogLine olds(i), "OK", "Renamed to " & news(i)
                ProgressStep done, news(i)
            Else
                placed(i) = True
                failed = failed + 1
                LogLine olds(i), "FAILED", "Left as " & olds(i) & " - " & Err.Description
            End If
        End If
    Next i

    ' The document number comes from the file name, so each renamed file is
    ' reopened and saved to put the new one in its title block.
    For i = 1 To n
        If FileExists(EndSep(folderPath) & news(i)) Then
            ProgressStep i, news(i)
            Set wb = OpenQuiet(EndSep(folderPath) & news(i), False)
            If wb Is Nothing Then
                LogLine news(i), "Skipped", "Renamed, but could not reopen it to " & _
                        "refresh the document number."
            Else
                Application.CalculateFull
                wb.Close SaveChanges:=True
            End If
            wsList.Cells(rowIdx(i), C_FILE).Value = news(i)
            wsList.Cells(rowIdx(i), C_NEWNAME).ClearContents
        End If
    Next i

    ProgressDone
    EndQuiet
    ShowSummary "Rename files", done, 0, failed, Timer - started, ""
    RefreshScheduleList
    Exit Sub

Fail:
    Recover "Rename files", wb
End Sub


' One rename. Returns False and leaves Err set when it will not go through.
Private Function RenameOne(ByVal folderPath As String, ByVal fromName As String, _
                           ByVal toName As String) As Boolean
    On Error Resume Next
    Err.Clear
    Name EndSep(folderPath) & fromName As EndSep(folderPath) & toName
    RenameOne = (Err.Number = 0)
    On Error GoTo 0
End Function


' A name nothing in the folder is using, for parking a file mid-swap.
Private Function FreeTempName(ByVal folderPath As String, ByVal fileName As String) As String
    Dim i As Long
    Dim candidate As String

    For i = 1 To 1000
        candidate = "_renaming" & i & "_" & fileName
        If Not FileExists(EndSep(folderPath) & candidate) Then
            FreeTempName = candidate
            Exit Function
        End If
    Next i

    FreeTempName = "_renaming" & Format$(Now, "hhnnss") & "_" & fileName
End Function


Private Function InBatch(ByRef olds() As String, ByVal n As Long, ByVal name As String) As Boolean
    Dim i As Long
    For i = 1 To n
        If StrComp(olds(i), name, vbTextCompare) = 0 Then
            InBatch = True
            Exit Function
        End If
    Next i
End Function


' Everything that would make one rename go wrong.
Private Function CheckRename(ByVal folderPath As String, ByVal old As String, _
                             ByVal proposed As String) As String
    Dim bad As String

    If Len(old) = 0 Then
        CheckRename = "A row has a new name but no current file name. "
        Exit Function
    End If

    If StrComp(old, proposed, vbBinaryCompare) = 0 Then Exit Function

    bad = BadNameChars(proposed)
    If Len(bad) > 0 Then _
        CheckRename = CheckRename & "'" & proposed & "' contains " & bad & ". "

    If LCase$(FileExtension(proposed)) <> LCase$(FileExtension(old)) Then _
        CheckRename = CheckRename & "'" & proposed & "' changes the file type. "

    If Not FileExists(EndSep(folderPath) & old) Then _
        CheckRename = CheckRename & "'" & old & "' is not in the folder. "

    If IsWorkbookOpen(old) Then _
        CheckRename = CheckRename & "'" & old & "' is open - close it first. "
End Function


' ===========================================================================
' Button 7 - tidy the sheets
'
' Every workbook the same way round: Front Cover, Revision Page, the schedule,
' then Metadata hidden at the end. A lone schedule sheet still called Sheet1
' gets a proper name.
'
' Only names, order and visibility change. Nothing on any sheet is touched.
' ===========================================================================
Public Sub TidySheets()
    Dim wsSetup As Worksheet
    Dim folderPath As String, fileName As String
    Dim wbTgt As Workbook
    Dim files As Collection, i As Long
    Dim backupDir As String
    Dim done As Long, skipped As Long, failed As Long
    Dim started As Double
    Dim oneLog As String

    Set wsSetup = GetSheet(ThisWorkbook, SH_SETUP)
    If wsSetup Is Nothing Then
        MsgBox "No Setup sheet. Run InstallTool first.", vbExclamation
        Exit Sub
    End If

    folderPath = SchedulesFolder()
    If Len(folderPath) = 0 Then Exit Sub

    Set files = ScheduleFiles(folderPath)
    If files.Count = 0 Then
        MsgBox "No Excel files found in:" & vbCrLf & folderPath, vbInformation
        Exit Sub
    End If

    If MsgBox("Tidy the sheets in " & files.Count & " workbook(s)?" & vbCrLf & vbCrLf & _
              "Order becomes Front Cover, Revision Page, Schedule, then Metadata " & _
              "hidden at the end." & vbCrLf & vbCrLf & _
              "A workbook with one schedule sheet has it renamed to 'Schedule'. " & _
              "One with several keeps its names, since they are the only thing " & _
              "telling those sheets apart." & vbCrLf & vbCrLf & _
              "Only names, order and visibility change. No cell is touched.", _
              vbQuestion + vbYesNo, "Tidy sheets") = vbNo Then Exit Sub

    If UCase$(Trim$(CStr(Opt(wsSetup, LBL_BACKUP, R_OPT_BACKUP).Value))) <> "NO" Then
        backupDir = EndSep(folderPath) & "_backup " & Format$(Now, "yyyy-mm-dd hh-nn")
        On Error Resume Next
        MkDir backupDir
        On Error GoTo 0
    End If

    On Error GoTo Fail
    LogStart "Tidy sheets"
    BeginQuiet xlCalculationManual
    ProgressStart files.Count, "Tidying sheets"
    started = Timer

    For i = 1 To files.Count
        fileName = files(i)
        ProgressStep i - 1, fileName

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
            oneLog = TidySheetsIn(wbTgt)

            If InStr(1, oneLog, "PROBLEM:", vbTextCompare) > 0 Then
                wbTgt.Close SaveChanges:=False
                failed = failed + 1
                LogLine fileName, "FAILED", oneLog & "Nothing was saved."
            ElseIf Len(oneLog) = 0 Then
                wbTgt.Close SaveChanges:=False
                skipped = skipped + 1
                LogLine fileName, "Unchanged", "Already tidy."
            Else
                wbTgt.Close SaveChanges:=True
                done = done + 1
                LogLine fileName, "OK", oneLog
            End If
        End If
    Next i

    ProgressDone
    EndQuiet
    ShowSummary "Tidy sheets", done, skipped, failed, Timer - started, _
                IIf(Len(backupDir) > 0, "Backup: " & backupDir, "")
    Exit Sub

Fail:
    Recover "Tidy sheets", wbTgt
End Sub


Public Sub Auto_Open()
    Dim wsSetup As Worksheet
    Set wsSetup = GetSheet(ThisWorkbook, SH_SETUP)
    If wsSetup Is Nothing Then Exit Sub
    If UCase$(Trim$(CStr(Opt(wsSetup, LBL_AUTO, R_OPT_AUTO).Value))) = "YES" Then
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
    Dim vals() As Variant
    Dim f As String

    Set d = CreateObject("Scripting.Dictionary")
    Set SnapshotList = d

    lastRow = wsList.Cells(wsList.Rows.Count, C_FILE).End(xlUp).Row
    If lastRow < 2 Then Exit Function

    For r = 2 To lastRow
        f = LCase$(AsText(wsList.Cells(r, C_FILE).Value))
        If Len(f) > 0 Then
            ReDim vals(1 To C_FILECHK)
            For c = 1 To C_FILECHK
                vals(c) = wsList.Cells(r, c).Value
            Next c
            d(f) = vals
        End If
    Next r
End Function


' A cached row is reusable only when the file has not been touched since.
Private Function CanReuse(ByVal keep As Object, ByVal fileName As String, _
                          ByVal stamp As Double) As Boolean
    Dim vals As Variant
    Dim old As Double

    If stamp = 0 Then Exit Function
    If Not keep.Exists(LCase$(fileName)) Then Exit Function

    vals = keep(LCase$(fileName))
    If Len(AsText(vals(C_DATA_1))) = 0 And Len(AsText(vals(C_CHECKS))) = 0 Then Exit Function

    On Error Resume Next
    old = CDbl(vals(C_STAMP))
    On Error GoTo 0

    CanReuse = (old <> 0) And (Abs(old - stamp) < 0.000001)
End Function


Private Sub RestoreRow(ByVal wsList As Worksheet, ByVal r As Long, ByVal vals As Variant)
    Dim c As Long
    For c = C_FILE To C_CHECKS
        wsList.Cells(r, c).Value = vals(c)
    Next c
    wsList.Cells(r, C_STAMP).Value = vals(C_STAMP)
    wsList.Cells(r, C_FILECHK).Value = vals(C_FILECHK)
End Sub


' Puts back the tick and any new-revision text the user had typed, matched by
' file name so it follows the row even if the order changed.
Private Sub RestoreTypedEntries(ByVal wsList As Worksheet, ByVal keep As Object, _
                                ByVal lastRow As Long)
    Dim r As Long, c As Long
    Dim f As String
    Dim vals As Variant

    For r = 2 To lastRow
        f = LCase$(AsText(wsList.Cells(r, C_FILE).Value))
        If keep.Exists(f) Then
            vals = keep(f)
            If Len(AsText(wsList.Cells(r, C_PICK).Value)) = 0 Then _
                wsList.Cells(r, C_PICK).Value = vals(C_PICK)
            For c = C_NEW_FIRST To C_NEWNAME
                If Len(AsText(wsList.Cells(r, c).Value)) = 0 Then wsList.Cells(r, c).Value = vals(c)
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
    TrimLog ws, LOG_RUNS - 1

    ' Newest run at the top, so the Log sheet opens on what just happened.
    ws.Rows("1:3").Insert Shift:=xlDown
    ws.Rows("1:3").ClearFormats

    ws.Range("A1").Value = what & " - " & Format$(Now, "dd/mm/yyyy hh:nn:ss")
    ws.Range("A1:C1").Font.Bold = True
    ws.Range("A1:C1").Interior.Color = RGB(221, 231, 244)
    ws.Range("E1").Value = "RUN"          ' marks where a run starts, for trimming

    ws.Range("A2").Value = "File"
    ws.Range("B2").Value = "Result"
    ws.Range("C2").Value = "Notes"
    ws.Range("A2:C2").Font.Bold = True

    ws.Columns("E").Hidden = True
    mLogRow = 3
End Sub


' Drops the oldest runs, keeping the newest `keep` of them.
Private Sub TrimLog(ByVal ws As Worksheet, ByVal keep As Long)
    Dim r As Long, lastRow As Long, runs As Long

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    For r = 1 To lastRow
        If StrComp(Trim$(CStr(ws.Cells(r, 5).Value)), "RUN", vbTextCompare) = 0 Then
            runs = runs + 1
            If runs > keep Then
                ws.Range(ws.Rows(r), ws.Rows(lastRow + 2)).Delete Shift:=xlUp
                Exit Sub
            End If
        End If
    Next r
End Sub


Private Sub LogLine(ByVal fileName As String, ByVal result As String, ByVal notes As String)
    Dim ws As Worksheet

    Set ws = GetSheet(ThisWorkbook, SH_LOG)
    If ws Is Nothing Then Exit Sub
    If mLogRow < 3 Then mLogRow = 3

    ws.Rows(mLogRow).Insert Shift:=xlDown
    ws.Rows(mLogRow).ClearFormats

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
    If Not ws Is Nothing Then
        ws.Columns("A:C").AutoFit
        If ws.Columns("C").ColumnWidth > 90 Then ws.Columns("C").ColumnWidth = 90
        ws.Columns("E").Hidden = True
    End If

    msg = okCount & " succeeded" & vbCrLf & _
          otherCount & " skipped / unchanged" & vbCrLf & _
          failCount & " failed" & vbCrLf & vbCrLf & _
          "Took " & Duration(seconds) & "."

    If Len(extra) > 0 Then msg = msg & vbCrLf & vbCrLf & extra
    msg = msg & vbCrLf & vbCrLf & "Line by line detail is at the top of the Log sheet, " & _
          "which keeps the last " & LOG_RUNS & " runs."

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
        v = Trim$(CStr(Opt(wsSetup, LBL_FOLDER, R_OPT_FOLDER).Value))
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
        If Not wsSetup Is Nothing Then Opt(wsSetup, LBL_FOLDER, R_OPT_FOLDER).Value = resolved
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

    If Not wsSetup Is Nothing Then Opt(wsSetup, LBL_FOLDER, R_OPT_FOLDER).Value = resolved
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


' The value cell for an option, found by its label in column A. Falls back to
' the row the current layout puts it on, for a sheet being built from scratch.
Private Function Opt(ByVal wsSetup As Worksheet, ByVal label As String, _
                     ByVal fallbackRow As Long) As Range
    Dim lbl As Range
    Set lbl = FindLabel(wsSetup, label, 6, 40)
    If lbl Is Nothing Then
        Set Opt = wsSetup.Cells(fallbackRow, 2)
    Else
        Set Opt = lbl.Offset(0, 1)
    End If
End Function


' The extra project fields, as a two-column array of name and the absolute
' address of the cell holding its value. Blank names are skipped, so gaps in
' the block are fine.
Private Function ProjectFields(ByVal wsSetup As Worksheet) As Variant
    Dim out() As Variant
    Dim r As Long, n As Long
    Dim nm As String
    Dim first As Range

    Set first = FindLabel(wsSetup, "PROJECT FIELDS", 6, 60)
    If first Is Nothing Then
        ProjectFields = Empty
        Exit Function
    End If

    ReDim out(1 To R_FLD_COUNT, 1 To 2)
    For r = first.Row + 1 To first.Row + R_FLD_COUNT
        nm = Trim$(CStr(wsSetup.Cells(r, 1).Value))
        If Len(nm) > 0 Then
            n = n + 1
            out(n, 1) = nm
            out(n, 2) = AbsRef(wsSetup.Cells(r, 2))
        End If
    Next r

    If n = 0 Then
        ProjectFields = Empty
    Else
        ReDim Preserve out(1 To R_FLD_COUNT, 1 To 2)
        ProjectFields = TrimFields(out, n)
    End If
End Function


Private Function TrimFields(ByVal src As Variant, ByVal n As Long) As Variant
    Dim out() As Variant
    Dim i As Long
    ReDim out(1 To n, 1 To 2)
    For i = 1 To n
        out(i, 1) = src(i, 1)
        out(i, 2) = src(i, 2)
    Next i
    TrimFields = out
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
    ' Start from a clean slate. Formats AND text, because the row a label sits
    ' on can move between versions and a leftover line is worse than no line.
    ' Column B is left alone - that is the user's data. So is column F, the
    ' suitability codes, and every other sheet in the workbook.
    ws.Range("A1:F60").ClearFormats
    ws.Range("A5:A60").ClearContents
    ws.Range("C1:C60").ClearContents
    ws.Range("H1:H30").ClearFormats
    ws.Range("H1:H30").ClearContents
    ws.Range("A1:A" & R_REV_FIRST + 6).Font.Color = CLR_LABEL

    ' Row heights back to standard. An earlier version shrank rows 2 and 5 as
    ' spacers, which squashed the suitability codes sharing those rows in F.
    ws.Rows(2).RowHeight = ws.StandardHeight
    ws.Rows(5).RowHeight = ws.StandardHeight
    ws.Rows(16).RowHeight = ws.StandardHeight
    ws.Rows(25).RowHeight = ws.StandardHeight

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
    ws.Cells(R_OPT_HFSRC, 1).Value = "Reference schedule"
    ws.Cells(R_OPT_HFIMG, 1).Value = "Header image"
    ws.Cells(R_OPT_HFSCALE, 1).Value = "Header image scale %"
    ws.Cells(R_OPT_SETUP, 1).Value = "Last setup run"
    ws.Cells(R_OPT_LIST, 1).Value = "Last list refresh"

    InputCell ws.Cells(R_OPT_FOLDER, 2)
    InputCell ws.Cells(R_OPT_BACKUP, 2)
    InputCell ws.Cells(R_OPT_AUTO, 2)
    InputCell ws.Cells(R_OPT_FULL, 2)
    InputCell ws.Cells(R_OPT_HFSRC, 2)
    InputCell ws.Cells(R_OPT_HFIMG, 2)
    InputCell ws.Cells(R_OPT_HFSCALE, 2)
    ReadOnlyCell ws.Cells(R_OPT_SETUP, 2)
    ReadOnlyCell ws.Cells(R_OPT_LIST, 2)

    ' Option rows have moved between versions, so a cell can be holding the
    ' value of whatever used to live on that row. Anything that is not a valid
    ' answer goes back to the default rather than being left to be read wrong.
    NormaliseYesNo ws.Cells(R_OPT_BACKUP, 2), "Yes"
    NormaliseYesNo ws.Cells(R_OPT_AUTO, 2), "No"
    NormaliseYesNo ws.Cells(R_OPT_FULL, 2), "No"

    If IsNumeric(ws.Cells(R_OPT_HFSRC, 2).Value) Then ws.Cells(R_OPT_HFSRC, 2).ClearContents
    If Not IsNumeric(ws.Cells(R_OPT_HFSCALE, 2).Value) Then ws.Cells(R_OPT_HFSCALE, 2).Value = 20
    If ws.Cells(R_OPT_HFSCALE, 2).Value <= 0 Then ws.Cells(R_OPT_HFSCALE, 2).Value = 20

    ' Written by the tool on its next run; whatever is there now is stale.
    ws.Cells(R_OPT_SETUP, 2).ClearContents
    ws.Cells(R_OPT_LIST, 2).ClearContents

    YesNoList ws.Cells(R_OPT_BACKUP, 2)
    YesNoList ws.Cells(R_OPT_AUTO, 2)
    YesNoList ws.Cells(R_OPT_FULL, 2)

    Note ws.Cells(R_OPT_FOLDER, 3), "blank = the folder this file is saved in"
    Note ws.Cells(R_OPT_BACKUP, 3), "copies every file into a timestamped folder first"
    Note ws.Cells(R_OPT_AUTO, 3), "reads every schedule when this file is opened"
    Note ws.Cells(R_OPT_FULL, 3), "No = only reopen files that changed since last time"
    Note ws.Cells(R_OPT_HFSRC, 3), "the schedule that is set up correctly; used by both copy buttons"
    Note ws.Cells(R_OPT_HFIMG, 3), "logo for the top-right of the header; blank = ask, or leave for none"
    Note ws.Cells(R_OPT_HFSCALE, 3), "size of that logo as a percentage of the image's own size"

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

    ' --- Extra project fields ---------------------------------------------
    SectionHeader ws, R_FLD_FIRST - 1, "PROJECT FIELDS"
    Note ws.Cells(R_FLD_FIRST - 1, 3), "anything else the whole project shares, e.g. DfE Code"
    InputCell ws.Range(ws.Cells(R_FLD_FIRST, 1), ws.Cells(R_FLD_FIRST + R_FLD_COUNT - 1, 2))

    ' --- Suitability codes -------------------------------------------------
    SectionHeader ws, 1, "SUITABILITY CODES", 6
    If SuitabilityCount(ws) = 0 Then SeedSuitabilityCodes ws
    If SuitabilityCount(ws) > 0 Then _
        InputCell ws.Range(ws.Cells(2, 6), ws.Cells(SuitabilityCount(ws) + 1, 6))
    SuitabilityList ws, ws.Cells(R_REV_FIRST + 1, 2)

    ' --- Notes beside the buttons -----------------------------------------
    ' Row 13 down, so the four buttons above never sit on top of them.
    Note ws.Range("H12"), "Cells shaded yellow are the ones you fill in."
    Note ws.Range("H13"), "Progress is shown in the status bar, bottom-left of the Excel window."
    Note ws.Range("H14"), "Every run writes a line per file to the Log sheet, then a summary."
    Note ws.Range("H16"), "Added a schedule? Press 'Set up / repair schedules' again - it is safe to re-run."
    Note ws.Range("H17"), "To reissue: on ScheduleList put an x in 'Add?', fill the blue 'New ...' columns,"
    Note ws.Range("H18"), "then press 'Add revision to ticked'. Blanks fall back to the block above."
    Note ws.Range("H20"), "Security classification: set the header/footer on one workbook by hand under"
    Note ws.Range("H21"), "Page Layout, then press 'Copy headers && footers' to push it to the rest."
    Note ws.Range("H22"), "Changed the cover or revision page layout itself? 'Copy cover && revision page'"
    Note ws.Range("H23"), "'Rename files' fills the New FileName column first so you can read it, then renames."
    Note ws.Range("H24"), "'Tidy sheets' orders every workbook the same way and hides Metadata. No cell is touched."
    Note ws.Range("H25"), "Schedule tool version " & TOOL_VERSION

    ' --- Layout ------------------------------------------------------------
    ws.Columns("A").ColumnWidth = 24
    ws.Columns("B").ColumnWidth = 34
    ws.Columns("C").ColumnWidth = 46
    ws.Columns("D:E").ColumnWidth = 3
    ws.Columns("F").ColumnWidth = 34
    ws.Columns("G").ColumnWidth = 3
    ws.Range("A1:F60").VerticalAlignment = xlCenter
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


Private Sub NormaliseYesNo(ByVal rng As Range, ByVal defaultValue As String)
    Select Case UCase$(Trim$(CStr(rng.Value)))
        Case "YES", "NO"
            ' fine as it is
        Case Else
            rng.Value = defaultValue
    End Select
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


' A hover note on a header cell. Explains a column without spending a row.
Private Sub HeaderNote(ByVal cell As Range, ByVal txt As String)
    On Error Resume Next
    cell.ClearComments
    cell.AddComment txt
    cell.Comment.Shape.TextFrame.AutoSize = True
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
              "New ApBy", "New Description", "New FileName", "_Stamp", "_FileChecks")

    For i = 0 To UBound(h)
        ws.Cells(1, i + 1).Value = h(i)
    Next i

    ' Own the header row completely. It inherited white text from the table
    ' style this sheet used to be, which was unreadable on a light fill.
    ws.Rows(1).ClearFormats
    ws.Rows(1).Font.Bold = True
    ws.Rows(1).Font.Color = RGB(0, 0, 0)
    ws.Range(ws.Cells(1, 1), ws.Cells(1, C_CHECKS)).Interior.Color = RGB(230, 230, 230)
    ws.Range(ws.Cells(1, C_NEW_FIRST), ws.Cells(1, C_NEWNAME)).Interior.Color = RGB(214, 232, 255)

    SetColumnState ws

    HeaderNote ws.Cells(1, C_PICK), _
        "Put an x here on every schedule you are reissuing, then press " & _
        "'Add revision to ticked' on the Setup sheet." & vbCrLf & vbCrLf & _
        "The revision itself comes from the blue 'New ...' columns on the same row. " & _
        "Anything you leave blank there is taken from the 'New revision' block on " & _
        "the Setup sheet, so common values only get typed once."
    HeaderNote ws.Cells(1, C_NEWNAME), _
        "The file name to rename this schedule to." & vbCrLf & vbCrLf & _
        "Press 'Rename files' once to fill this column in by find and replace, " & _
        "check what it proposes, then press it again to do the renaming. Or just " & _
        "type a name here yourself." & vbCrLf & vbCrLf & _
        "Leave the extension off and the current one is kept."
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


' Hidden-ness and number formats for the whole sheet, in one place.
'
' The columns have moved as inputs were added, and a column that was hidden
' under an old layout stayed hidden under the new one, with its old number
' format still on it. That is how the New FileName column arrived invisible
' and formatted as a date. Every column is now explicitly set, every time.
Private Sub SetColumnState(ByVal ws As Worksheet)
    Dim c As Long

    For c = 1 To C_FILECHK
        ws.Columns(c).Hidden = False
    Next c

    ws.Columns(C_STAMP).Hidden = True
    ws.Columns(C_FILECHK).Hidden = True

    ws.Columns(9).NumberFormat = "dd/mm/yyyy"                 ' Date
    ws.Columns(C_NEW_FIRST + 2).NumberFormat = "dd/mm/yyyy"   ' New Date
    ws.Columns(C_NEWNAME).NumberFormat = "@"                  ' file names are text
    ws.Columns(C_STAMP).NumberFormat = "dd/mm/yyyy hh:mm"
End Sub


Private Sub FormatList(ByVal ws As Worksheet, ByVal lastRow As Long)
    Dim wsSetup As Worksheet
    Dim codes As Long
    Dim rng As Range

    ' Clearing the rows resets number formats to General every refresh.
    SetColumnState ws

    ws.Range(ws.Cells(1, 1), ws.Cells(1, C_NEWNAME)).EntireColumn.AutoFit
    If ws.Columns(C_CHECKS).ColumnWidth > 60 Then ws.Columns(C_CHECKS).ColumnWidth = 60
    If ws.Columns(C_NEWNAME).ColumnWidth > 60 Then ws.Columns(C_NEWNAME).ColumnWidth = 60

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
    Dim i As Long
    Dim leftCol As Double, topRow As Double

    For i = ws.Buttons.Count To 1 Step -1
        ws.Buttons(i).Delete
    Next i

    leftCol = ws.Range("H2").Left
    topRow = ws.Range("H2").Top

    ' Two columns of three: routine jobs on the left, bulk edits on the right.
    AddButton ws, leftCol, topRow, "SetupProject", "Set up / repair schedules"
    AddButton ws, leftCol, topRow + 36, "RefreshScheduleList", "Refresh schedule list"
    AddButton ws, leftCol, topRow + 72, "AddRevisionToTicked", "Add revision to ticked"
    AddButton ws, leftCol, topRow + 108, "TidySheets", "Tidy sheets"

    AddButton ws, leftCol + 210, topRow, "CopyHeadersFooters", "Copy headers && footers"
    AddButton ws, leftCol + 210, topRow + 36, "CopyCommonSheets", "Copy cover && revision page"
    AddButton ws, leftCol + 210, topRow + 72, "RenameFiles", "Rename files"
End Sub


Private Sub AddButton(ByVal ws As Worksheet, ByVal x As Double, ByVal y As Double, _
                      ByVal macroName As String, ByVal label As String)
    Dim b As Object
    Set b = ws.Buttons.Add(x, y, 200, 30)
    b.OnAction = macroName
    b.Caption = label
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
