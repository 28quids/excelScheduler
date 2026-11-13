Attribute VB_Name = "modMain"
Option Explicit

' ---------------------------------------------------------------------------
' modMain - the buttons.
'
'   InstallTool           run once, after importing the modules
'   RefreshScheduleList   read every workbook in the folder back and QA it
'   AddRevisionToTicked   append a revision line to the ticked schedules
'   ExportToPdf           PDFs of the ticked schedules
'   SetupNewSchedules     everything a schedule new to the folder needs
'   SetupProject          link/repair schedules
'   CopyHeadersFooters    headers and footers from the reference schedule
'   CopyCommonSheets      cover and revision page from the reference schedule
'   SetClassification     the security marking, from the Setup sheet
'   DrawTitleBlocks       the title block and notes box on the schedule sheet
'   TidySheets            sheet order, names and visibility
'   RenameFiles           find and replace across file names
'
' The x in column A of ScheduleList is the selection for every button that
' changes a schedule: ticked rows only, and with nothing ticked the button
' says it is about to do ALL of them before it does.
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
Private Const R_OPT_CLASS   As Long = 14
Private Const R_OPT_SETUP   As Long = 15   ' read only
Private Const R_OPT_LIST    As Long = 16   ' read only
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
Private Const LBL_CLASS   As String = "Security classification"
Private Const LBL_SETUP   As String = "Last setup run"
Private Const LBL_LIST    As String = "Last list refresh"

' Setup sheet palette. Yellow means "you fill this in".
Private Const CLR_INPUT     As Long = 14810111   ' RGB(255, 251, 225)
Private Const CLR_READONLY  As Long = 15921906   ' RGB(242, 242, 242)
Private Const CLR_NOTE      As Long = 8421504    ' RGB(128, 128, 128)
Private Const CLR_SECTION   As Long = 6575172    ' RGB(68, 84, 100)
Private Const CLR_LABEL     As Long = 4210752    ' RGB(64, 64, 64)

' ScheduleList columns. Columns 3-15 are read from each schedule's Metadata
' sheet in FillRow, in the order of LIST_HEADERS below.
Private Const C_PICK      As Long = 1
Private Const C_FILE      As Long = 2
Private Const C_DATA_1    As Long = 3      ' ScheduleName
Private Const C_PROJNAME  As Long = 4
Private Const C_PROJNO    As Long = 5
Private Const C_CLIENT    As Long = 6
Private Const C_REV       As Long = 8
Private Const C_DATE      As Long = 9
Private Const C_REVDESC   As Long = 16     ' latest revision's Description, from RevisionTable
Private Const C_CLASS     As Long = 17     ' security marking, from headers / footers / cover
Private Const C_CHECKS    As Long = 18
Private Const C_NEW_FIRST As Long = 19     ' New Rev..New Description = 19..25
Private Const C_NEWNAME   As Long = 26     ' proposed file name, for the rename
Private Const C_STAMP     As Long = 27     ' hidden, file modified time
Private Const C_FILECHK   As Long = 28     ' hidden, checks that come from the file itself
Private Const REV_FIELDS  As Long = 7      ' Rev, Status, Date, Pr, Ch, Ap, Descr

' Every ScheduleList header, in column order. Also how a refresh recognises a
' list written under an older layout, whose rows it must not reuse.
Private Const LIST_HEADERS As String = _
    "Select|FileName|ScheduleName|ProjectName|ProjectNo|Client|DocType|Revision|Date|" & _
    "PrBy|ChBy|ApBy|DocumentNo|SuitabilitySt|SuitabilityDs|RevDescription|Classification|" & _
    "Checks|New Rev|New Status|New Date|New PrBy|New ChBy|New ApBy|New Description|" & _
    "New FileName|_Stamp|_FileChecks"

' What the refresh writes in the hidden file-checks column for a workbook it
' cannot compare: never set up, or not a schedule yet. Rows starting with
' this are left out of the project-wide comparisons.
Private Const NOT_READY   As String = "NOT "

' The values offered for the Security classification option. Anything the
' marking pattern in modSchedule knows is accepted if typed instead.
Private Const CLASS_CHOICES As String = _
    "None,OFFICIAL,OFFICIAL-SENSITIVE,CONFIDENTIAL,COMMERCIAL IN CONFIDENCE"

' Where the PDFs go by default, inside the folder the schedules came from.
Private Const PDF_SUBFOLDER As String = "_pdf"

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
    wsList.Range(wsList.Cells(2, 1), wsList.Cells(wsList.Rows.Count, C_FILECHK)).Hyperlinks.Delete
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
           "1. Fill in Client, Project Name and Project Number on this sheet, and the " & _
           "Security classification." & vbCrLf & _
           "2. Save this file into the project folder with the schedules." & vbCrLf & _
           "3. Press 'Refresh schedule list'." & vbCrLf & _
           "4. Press 'Set up / repair links' with nothing ticked, and say Yes to all.", _
           vbInformation, "Schedule tool"
End Sub


' ===========================================================================
' Set up / repair links - point the ticked schedules at this MAINPROJECTINFO
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
    Dim isAll As Boolean, leftOut As String
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

    Set files = TargetFiles(folderPath, "", isAll, leftOut)
    If files.Count = 0 Then
        MsgBox "Nothing to set up." & vbCrLf & vbCrLf & leftOut, vbInformation, "Set up / repair links"
        Exit Sub
    End If

    If Not Confirmed(PathFormWarning(folderPath) & _
              "Set up / repair " & ScopeText(files, isAll, leftOut) & vbCrLf & _
              "Each one is opened, linked to this file, its local links rebuilt, and " & _
              "saved. Safe to run again on a schedule that is already right." & vbCrLf & vbCrLf & _
              BackupLine(wsSetup) & vbCrLf & vbCrLf & "Continue?", _
              "Set up / repair links", isAll) Then Exit Sub

    statuses = GatherStatuses(wsSetup)
    backupDir = MakeBackupDir(wsSetup, folderPath)

    On Error GoTo Fail
    LogStart "Set up / repair links"
    BeginQuiet xlCalculationAutomatic
    ProgressStart files.Count, "Setting up schedules"
    started = Timer

    For i = 1 To files.Count
        fileName = files(i)
        fullPath = EndSep(folderPath) & fileName
        ProgressStep i - 1, fileName

        If Not OpenInExcel(fileName, skipped) Then
            BackUp backupDir, folderPath, fileName
            Set wbTgt = OpenQuiet(fullPath, False)

            If wbTgt Is Nothing Then
                failed = failed + 1
                LogLine fileName, "FAILED", "Could not open the file."
            ElseIf Not LooksLikeSchedule(wbTgt) Then
                wbTgt.Close SaveChanges:=False
                skipped = skipped + 1
                LogLine fileName, "Skipped", "No Revision Page yet. Use 'Set up new " & _
                        "schedule(s)', which gives it one from the reference schedule."
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
                On Error GoTo Fail
            End If
            Set wbTgt = Nothing
        End If
    Next i

    ProgressDone
    Opt(wsSetup, LBL_SETUP, R_OPT_SETUP).Value = Format$(Now, "dd/mm/yyyy hh:nn")
    EndQuiet

    ShowSummary "Set up / repair links", done, skipped, failed, _
                Timer - started, IIf(Len(backupDir) > 0, "Backup: " & backupDir, "")

    RefreshScheduleList
    Exit Sub

Fail:
    Recover "Set up / repair links", wbTgt
End Sub


' ===========================================================================
' Refresh schedule list - read every workbook in the folder back and check it
'
' Every workbook is listed, including one that is not a schedule yet, so a
' file just dropped into the folder shows up and can be ticked.
' ===========================================================================
Public Sub RefreshScheduleList()
    Dim wsSetup As Worksheet, wsList As Worksheet
    Dim folderPath As String, fileName As String, fullPath As String
    Dim wbTgt As Workbook
    Dim files As Collection, i As Long
    Dim r As Long
    Dim mpiName As String, mpiNo As String, mpiClient As String
    Dim wantClass As String
    Dim keep As Object
    Dim fullRefresh As Boolean, wasOpen As Boolean
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
    wantClass = ClassificationSetting(wsSetup)
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
    wsList.Range(wsList.Cells(2, 1), wsList.Cells(wsList.Rows.Count, C_FILECHK)).Hyperlinks.Delete
    wsList.Range(wsList.Cells(2, 1), wsList.Cells(wsList.Rows.Count, C_FILECHK)).Clear

    r = 2
    For i = 1 To files.Count
        fileName = files(i)
        fullPath = EndSep(folderPath) & fileName
        stamp = FileStamp(fullPath)
        ProgressStep i - 1, fileName

        If Not fullRefresh And CanReuse(keep, fileName, stamp) Then
            RestoreRow wsList, r, keep(LCase$(fileName))
            LinkFileCell wsList.Cells(r, C_FILE), fullPath
            LogLine fileName, "Unchanged", "Not reopened - same as the last refresh."
            reused = reused + 1
            r = r + 1
        Else
            ' One you have open is read as it stands and left open. Closing
            ' it would lose whatever you have not saved.
            wasOpen = IsWorkbookOpen(fileName)
            If wasOpen Then
                Set wbTgt = Workbooks(fileName)
            Else
                Set wbTgt = OpenQuiet(fullPath, True)
            End If

            If wbTgt Is Nothing Then
                failed = failed + 1
                LogLine fileName, "FAILED", "Could not open the file."
            Else
                If LooksLikeSchedule(wbTgt) Then
                    FillRow wsList, r, fileName, wbTgt
                    LogLine fileName, "Read", AsText(wsList.Cells(r, C_FILECHK).Value) & _
                            IIf(wasOpen, " Read from the copy open in Excel, including " & _
                                         "anything not saved yet.", "")
                Else
                    wsList.Cells(r, C_FILE).Value = fileName
                    wsList.Cells(r, C_FILECHK).Value = NOT_READY & "A SCHEDULE YET - " & _
                        "no Revision Page. Tick it and press 'Set up new schedule(s)'."
                    LogLine fileName, "Read", "Not a schedule yet - listed so it can be ticked."
                End If
                ' An open file's date on disk says nothing about what is in
                ' memory, so it is never reused next time.
                If Not wasOpen Then wsList.Cells(r, C_STAMP).Value = stamp
                LinkFileCell wsList.Cells(r, C_FILE), fullPath
                readCount = readCount + 1
                r = r + 1
                If Not wasOpen Then wbTgt.Close SaveChanges:=False
            End If
            Set wbTgt = Nothing
        End If
    Next i

    RestoreTypedEntries wsList, keep, r - 1
    ComposeChecks wsList, r - 1, mpiName, mpiNo, mpiClient, wantClass
    FlagOddOnesOut wsList, r - 1, (Len(wantClass) = 0)
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
    If wasOpen Then Set wbTgt = Nothing     ' never close a workbook you had open
    Recover "Refresh schedule list", wbTgt
End Sub


' ===========================================================================
' Add revision to ticked - append a revision line to every ticked schedule
'
' Each row can carry its own new revision in the "New ..." columns. Anything
' left blank there falls back to the block on the Setup sheet, so reissuing
' all 24 on one revision and reissuing 6 of them on different ones are the
' same operation.
'
' The tick is kept afterwards, so "Export ticked to PDF" straight after
' issues exactly the schedules that were just revised. Pressing this twice
' does no harm: a revision already in a table is refused.
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
               "Put an x in column A of ScheduleList next to each schedule being reissued, " & _
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
            ElseIf Not OpenInExcel(fileName, skipped) Then
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
                    Set wbTgt = Nothing
                End If
            End If
        End If
    Next r

    ProgressDone
    EndQuiet
    ShowSummary "Add revision", done, skipped, failed, Timer - started, _
                IIf(done > 0, "The ticks are kept, so 'Export ticked to PDF' now issues " & _
                              "these same schedules.", "")
    RefreshScheduleList
    Exit Sub

Fail:
    Recover "Add revision", wbTgt
End Sub


' ===========================================================================
' Copy headers & footers from the reference schedule to the ticked ones
'
' For copying a whole header and footer design across: a logo, a page number,
' a document reference. For only the security marking, "Set classification"
' is the direct way and touches nothing else.
' ===========================================================================
Public Sub CopyHeadersFooters()
    Dim wsSetup As Worksheet
    Dim folderPath As String, srcPath As String, fileName As String
    Dim wbSrc As Workbook, wbTgt As Workbook
    Dim files As Collection, i As Long
    Dim isAll As Boolean, leftOut As String
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

    If Not HeaderImageSize(wsSetup, folderPath, imgPath, imgW, imgH, scalePct) Then Exit Sub

    Set files = TargetFiles(folderPath, srcPath, isAll, leftOut)
    If files.Count = 0 Then
        MsgBox "Nothing to copy to." & vbCrLf & vbCrLf & leftOut, vbInformation, "Copy headers & footers"
        Exit Sub
    End If

    If Not Confirmed("Copy the headers and footers from" & vbCrLf & vbCrLf & _
              "    " & BaseName(srcPath) & vbCrLf & vbCrLf & _
              "into " & ScopeText(files, isAll, leftOut) & vbCrLf & _
              "Front Cover, Revision Page and the schedule sheets are matched up " & _
              "separately. Whatever those sheets currently have is replaced." & vbCrLf & vbCrLf & _
              IIf(Len(imgPath) > 0, "Header image: " & BaseName(imgPath) & " at " & _
                  CStr(scalePct) & "%, top right.", "No header image set.") & vbCrLf & vbCrLf & _
              BackupLine(wsSetup) & vbCrLf & vbCrLf & "Continue?", _
              "Copy headers & footers", isAll) Then Exit Sub

    backupDir = MakeBackupDir(wsSetup, folderPath)

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

        If Not OpenInExcel(fileName, skipped) Then
            BackUp backupDir, folderPath, fileName
            Set wbTgt = OpenQuiet(EndSep(folderPath) & fileName, False)
            If wbTgt Is Nothing Then
                failed = failed + 1
                LogLine fileName, "FAILED", "Could not open the file."
            Else
                oneLog = CopyHeadersFootersTo(wbSrc, wbTgt, imgPath, imgW, imgH)
                wbTgt.Close SaveChanges:=True
                Set wbTgt = Nothing
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
    Set wbSrc = Nothing
    ProgressDone
    EndQuiet

    ShowSummary "Copy headers & footers", done, skipped, failed, Timer - started, _
                IIf(Len(imgPath) > 0, "Header image set from " & BaseName(imgPath) & _
                    ". Check one print preview." & vbCrLf & vbCrLf, "") & _
                IIf(Len(backupDir) > 0, "Backup: " & backupDir, "")

    ' The Classification column reads the headers and footers.
    RefreshScheduleList
    Exit Sub

Fail:
    On Error Resume Next
    If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False
    On Error GoTo 0
    Recover "Copy headers & footers", wbTgt
End Sub


' The header logo and the size to put it in at. False means stop: the image
' set on the Setup sheet could not be read. No image at all is a valid answer
' and returns True with imgPath "".
Private Function HeaderImageSize(ByVal wsSetup As Worksheet, ByVal folderPath As String, _
                                 ByRef imgPath As String, ByRef imgW As Double, _
                                 ByRef imgH As Double, ByRef scalePct As Double) As Boolean
    imgPath = HeaderImagePath(wsSetup, folderPath)
    If Len(imgPath) > 0 Then
        If Not MeasureImage(imgPath, imgW, imgH) Then
            MsgBox "Could not read that image:" & vbCrLf & imgPath, vbExclamation
            Exit Function
        End If
        scalePct = 0
        On Error Resume Next
        scalePct = CDbl(Opt(wsSetup, LBL_HFSCALE, R_OPT_HFSCALE).Value)
        On Error GoTo 0
        If scalePct <= 0 Then scalePct = 20
        imgW = imgW * scalePct / 100
        imgH = imgH * scalePct / 100
    End If
    HeaderImageSize = True
End Function


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
' Copy cover & revision page from the reference schedule to the ticked ones
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
    Dim isAll As Boolean, leftOut As String
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

    Set files = TargetFiles(folderPath, srcPath, isAll, leftOut)
    If files.Count = 0 Then
        MsgBox "Nothing to copy to." & vbCrLf & vbCrLf & leftOut, vbInformation, _
               "Copy cover & revision page"
        Exit Sub
    End If

    If Not Confirmed("Replace the Front Cover and Revision Page in " & _
              ScopeText(files, isAll, leftOut) & vbCrLf & _
              "with the ones from" & vbCrLf & vbCrLf & _
              "    " & BaseName(srcPath) & vbCrLf & vbCrLf & _
              "Each schedule keeps its own:" & vbCrLf & _
              "  - schedule title" & vbCrLf & _
              "  - full revision history" & vbCrLf & _
              "  - document type, Delref, BSUID, trigger events" & vbCrLf & vbCrLf & _
              "Everything else on those two sheets is replaced, and the links " & _
              "are rebuilt to point at each schedule's own data." & vbCrLf & vbCrLf & _
              BackupLine(wsSetup) & vbCrLf & vbCrLf & "Continue?", _
              "Copy cover & revision page", True) Then Exit Sub

    statuses = GatherStatuses(wsSetup)
    backupDir = MakeBackupDir(wsSetup, folderPath)

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

        If Not OpenInExcel(fileName, skipped) Then
            BackUp backupDir, folderPath, fileName
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
                Set wbTgt = Nothing
            End If
        End If
    Next i

    wbSrc.Close SaveChanges:=False
    Set wbSrc = Nothing
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


' ===========================================================================
' Set up new schedule(s) - everything a schedule new to the folder needs, in
' one pass, for the ticked rows only
'
' What used to be four buttons pressed in the right order, which nobody
' remembers after a month away. Each ticked file is opened once and gets:
'
'   1. Front Cover and Revision Page from the reference schedule (its own
'      title and any revision history are kept, as with the copy button)
'   2. links set up to this file
'   3. the title block and notes box on the schedule sheet, added if missing
'   4. headers and footers from the reference, logo included
'   5. sheets put in order, Metadata hidden
'
' then saved once. A file that fails any step is closed without saving, so a
' half-done schedule is never written. It never runs on everything: with
' nothing ticked it explains how to tick, and stops.
' ===========================================================================
Public Sub SetupNewSchedules()
    Dim wsSetup As Worksheet
    Dim folderPath As String, srcPath As String, fileName As String
    Dim wbSrc As Workbook, wbTgt As Workbook
    Dim files As Collection, i As Long
    Dim isAll As Boolean, leftOut As String
    Dim backupDir As String
    Dim done As Long, skipped As Long, failed As Long
    Dim started As Double
    Dim oneLog As String, stepLog As String
    Dim projNameRef As String, projNoRef As String, clientRef As String
    Dim statuses As Variant
    Dim imgPath As String
    Dim imgW As Double, imgH As Double, scalePct As Double
    Dim changed As Long
    Const JOB As String = "Set up new schedule(s)"

    Set wsSetup = GetSheet(ThisWorkbook, SH_SETUP)
    If wsSetup Is Nothing Then
        MsgBox "No Setup sheet. Run InstallTool first.", vbExclamation
        Exit Sub
    End If

    If Not SetupRefs(wsSetup, projNameRef, projNoRef, clientRef) Then Exit Sub

    folderPath = SchedulesFolder()
    If Len(folderPath) = 0 Then Exit Sub

    ' Ticks first, before asking for a reference or a logo, so someone who
    ' has not ticked anything is told how before being asked anything else.
    Set files = TargetFiles(folderPath, "", isAll, leftOut)
    If isAll Or files.Count = 0 Then
        MsgBox "Nothing is ticked." & vbCrLf & vbCrLf & _
               "To bring a new schedule in:" & vbCrLf & _
               "  1. Put the file in the project folder." & vbCrLf & _
               "  2. Press 'Refresh schedule list'. It shows up as 'NOT A SCHEDULE YET'." & vbCrLf & _
               "  3. Put an x in column A next to it on ScheduleList." & vbCrLf & _
               "  4. Press this button again." & vbCrLf & vbCrLf & _
               "This button only ever works on ticked rows." & _
               IIf(Len(leftOut) > 0, vbCrLf & vbCrLf & leftOut, ""), vbInformation, JOB
        Exit Sub
    End If

    srcPath = HeaderSourcePath(wsSetup, folderPath)
    If Len(srcPath) = 0 Then Exit Sub

    ' TargetFiles was asked without the reference so the message above could
    ' come first. Take it out now if it is ticked.
    For i = files.Count To 1 Step -1
        If IsSamePath(EndSep(folderPath) & files(i), srcPath) Then
            files.Remove i
            leftOut = leftOut & "The reference schedule is ticked; it is the one being " & _
                   "copied from, so it is left out. "
        End If
    Next i
    If files.Count = 0 Then
        MsgBox "Only the reference schedule is ticked, and it is the one being copied from." & _
               vbCrLf & vbCrLf & "Tick the new schedule(s) instead.", vbInformation, JOB
        Exit Sub
    End If

    If Not HeaderImageSize(wsSetup, folderPath, imgPath, imgW, imgH, scalePct) Then Exit Sub

    If Not Confirmed("Set up " & ScopeText(files, False, leftOut) & vbCrLf & _
              "using " & BaseName(srcPath) & " as the reference. Each one gets:" & vbCrLf & vbCrLf & _
              "  1. the reference's Front Cover and Revision Page" & vbCrLf & _
              "     (its own title and any revision history are kept)" & vbCrLf & _
              "  2. its links set up to this file" & vbCrLf & _
              "  3. the title block and notes box, fitted to its columns" & vbCrLf & _
              "     (added above the schedule if it has none)" & vbCrLf & _
              "  4. the reference's headers and footers" & _
              IIf(Len(imgPath) > 0, ", with " & BaseName(imgPath), "") & vbCrLf & _
              "  5. its sheets put in order, Metadata hidden" & vbCrLf & vbCrLf & _
              "Then it is saved, once. A file that fails any step is not saved at all." & _
              vbCrLf & vbCrLf & BackupLine(wsSetup) & vbCrLf & vbCrLf & "Continue?", _
              JOB, False) Then Exit Sub

    statuses = GatherStatuses(wsSetup)
    backupDir = MakeBackupDir(wsSetup, folderPath)

    On Error GoTo Fail
    LogStart JOB
    BeginQuiet xlCalculationAutomatic
    ProgressStart files.Count, "Setting up new schedules"
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

        If Not OpenInExcel(fileName, skipped) Then
            BackUp backupDir, folderPath, fileName
            Set wbTgt = OpenQuiet(EndSep(folderPath) & fileName, False)
            If wbTgt Is Nothing Then
                failed = failed + 1
                LogLine fileName, "FAILED", "Could not open the file."
            Else
                oneLog = "1. " & CopyCommonSheetsTo(wbSrc, wbTgt, fileName)

                If InStr(1, oneLog, "PROBLEM:", vbTextCompare) = 0 Then
                    oneLog = oneLog & "2. " & _
                             RepairWorkbook(wbTgt, ThisWorkbook.FullName, SH_SETUP, _
                                            projNameRef, projNoRef, clientRef, statuses, _
                                            ProjectFields(wsSetup)) & " "
                    oneLog = oneLog & "3. " & BuildTitleBlocks(wbTgt, changed)
                    If InStr(1, oneLog, "PROBLEM:", vbTextCompare) = 0 Then
                        stepLog = CopyHeadersFootersTo(wbSrc, wbTgt, imgPath, imgW, imgH)
                        oneLog = oneLog & "4. " & stepLog
                        If InStr(1, stepLog, "PROBLEM:", vbTextCompare) = 0 Then
                            oneLog = oneLog & "5. " & TidySheetsIn(wbTgt)
                        End If
                    End If
                End If

                If InStr(1, oneLog, "PROBLEM:", vbTextCompare) > 0 Then
                    wbTgt.Close SaveChanges:=False
                    failed = failed + 1
                    LogLine fileName, "FAILED", oneLog & " Nothing was saved."
                Else
                    wbTgt.Close SaveChanges:=True
                    done = done + 1
                    LogLine fileName, "OK", oneLog
                End If
                Set wbTgt = Nothing
            End If
        End If
    Next i

    wbSrc.Close SaveChanges:=False
    Set wbSrc = Nothing
    ProgressDone
    EndQuiet

    ShowSummary JOB, done, skipped, failed, Timer - started, _
                "Next: check the Classification column on ScheduleList, and add a " & _
                "revision if this schedule is being issued." & vbCrLf & vbCrLf & _
                IIf(Len(backupDir) > 0, "Backup: " & backupDir, "")
    RefreshScheduleList
    Exit Sub

Fail:
    On Error Resume Next
    If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False
    On Error GoTo 0
    Recover JOB, wbTgt
End Sub


' ===========================================================================
' Set classification - put the marking from the Setup sheet on the ticked
' schedules, and nothing else
'
' Copy headers & footers can do this too, but only by copying the whole
' header and footer from a reference, which is a blunt tool for changing one
' word and easy to point at the wrong reference. This changes the marking
' where it already is and adds it where it is missing. See
' SetClassificationIn for exactly what it touches.
' ===========================================================================
Public Sub SetClassification()
    Dim wsSetup As Worksheet
    Dim folderPath As String, fileName As String
    Dim wbTgt As Workbook
    Dim files As Collection, i As Long
    Dim isAll As Boolean, leftOut As String
    Dim wanted As String, shown As String
    Dim backupDir As String
    Dim done As Long, skipped As Long, failed As Long
    Dim started As Double
    Dim oneLog As String
    Dim changed As Long
    Const JOB As String = "Set classification"

    Set wsSetup = GetSheet(ThisWorkbook, SH_SETUP)
    If wsSetup Is Nothing Then
        MsgBox "No Setup sheet. Run InstallTool first.", vbExclamation
        Exit Sub
    End If

    wanted = ClassificationSetting(wsSetup)
    If Len(wanted) = 0 Or (wanted <> "NONE" And Not IsKnownMarking(wanted)) Then
        MsgBox "Choose the marking in '" & LBL_CLASS & "' on the Setup sheet first:" & vbCrLf & _
               "None, OFFICIAL, OFFICIAL-SENSITIVE, CONFIDENTIAL or COMMERCIAL IN CONFIDENCE." & _
               vbCrLf & vbCrLf & "Then tick the schedules on ScheduleList and press this again.", _
               vbExclamation, JOB
        wsSetup.Activate
        Opt(wsSetup, LBL_CLASS, R_OPT_CLASS).Select
        Exit Sub
    End If
    If wanted = "NONE" Then
        shown = "no marking"
        wanted = ""
    Else
        wanted = NormaliseMarking(wanted)
        shown = wanted
    End If

    folderPath = SchedulesFolder()
    If Len(folderPath) = 0 Then Exit Sub

    Set files = TargetFiles(folderPath, "", isAll, leftOut)
    If files.Count = 0 Then
        MsgBox "Nothing to change." & vbCrLf & vbCrLf & leftOut, vbInformation, JOB
        Exit Sub
    End If

    If Not Confirmed("Set the security classification to " & shown & " on " & _
              ScopeText(files, isAll, leftOut) & vbCrLf & _
              IIf(Len(wanted) = 0, _
                  "Every marking is taken out of the headers and footers, and any cell " & _
                  "on the cover or revision page that says nothing but a marking is " & _
                  "emptied.", _
                  "Every marking already in a header or footer, or in a cell of its own " & _
                  "on the cover or revision page, becomes " & wanted & ", keeping its font " & _
                  "and size. A sheet with no marking at all gets " & wanted & " in the " & _
                  "centre of its footer, and of its header if that is empty.") & vbCrLf & vbCrLf & _
              "Nothing else in the headers and footers is touched: not the logo, the " & _
              "page numbers, orientation or margins. Each file is read back afterwards " & _
              "to prove it took." & vbCrLf & vbCrLf & _
              BackupLine(wsSetup) & vbCrLf & vbCrLf & "Continue?", JOB, isAll) Then Exit Sub

    backupDir = MakeBackupDir(wsSetup, folderPath)

    On Error GoTo Fail
    LogStart JOB & " to " & shown
    BeginQuiet xlCalculationManual
    ProgressStart files.Count, "Setting classification"
    started = Timer

    For i = 1 To files.Count
        fileName = files(i)
        ProgressStep i - 1, fileName

        If Not OpenInExcel(fileName, skipped) Then
            BackUp backupDir, folderPath, fileName
            Set wbTgt = OpenQuiet(EndSep(folderPath) & fileName, False)
            If wbTgt Is Nothing Then
                failed = failed + 1
                LogLine fileName, "FAILED", "Could not open the file."
            Else
                oneLog = SetClassificationIn(wbTgt, wanted, changed)
                If InStr(1, oneLog, "PROBLEM:", vbTextCompare) > 0 Then
                    wbTgt.Close SaveChanges:=False
                    failed = failed + 1
                    LogLine fileName, "FAILED", oneLog & "Nothing was saved."
                ElseIf changed = 0 Then
                    wbTgt.Close SaveChanges:=False
                    skipped = skipped + 1
                    LogLine fileName, "Unchanged", "Already " & shown & " throughout."
                Else
                    wbTgt.Close SaveChanges:=True
                    done = done + 1
                    LogLine fileName, "OK", oneLog
                End If
                Set wbTgt = Nothing
            End If
        End If
    Next i

    ProgressDone
    EndQuiet
    ShowSummary JOB, done, skipped, failed, Timer - started, _
                IIf(Len(backupDir) > 0, "Backup: " & backupDir, "")
    RefreshScheduleList
    Exit Sub

Fail:
    Recover JOB, wbTgt
End Sub


' The Security classification option, upper case. "" when blank, "NONE" for
' None. Blank means "do not check it against anything".
Private Function ClassificationSetting(ByVal wsSetup As Worksheet) As String
    ClassificationSetting = UCase$(Trim$(CStr(Opt(wsSetup, LBL_CLASS, R_OPT_CLASS).Value)))
End Function


' ===========================================================================
' Which schedules a button works on
'
' One rule for every button that changes a schedule: the rows with an x in
' column A of ScheduleList, and only those. With nothing ticked it is every
' workbook in the folder, and the confirmation box says ALL in capitals and
' defaults to No. Most of what these buttons do cannot be undone except from
' the backup, so doing the one schedule that needs it, rather than all 26, is
' the safer habit and the tool makes it the easy one.
' ===========================================================================

' The file names to work on. isAll comes back True when nothing was ticked.
' `skipPath`, the reference schedule, is never included. `leftOut` says what was
' ticked but left out, for the confirmation box.
Private Function TargetFiles(ByVal folderPath As String, ByVal skipPath As String, _
                             ByRef isAll As Boolean, ByRef leftOut As String) As Collection
    Dim c As New Collection
    Dim all As Collection
    Dim wsList As Worksheet
    Dim r As Long, lastRow As Long, i As Long
    Dim f As String, gone As String
    Dim tickedAny As Boolean

    isAll = False
    leftOut = ""
    Set TargetFiles = c

    Set wsList = GetSheet(ThisWorkbook, SH_LIST)
    If Not wsList Is Nothing Then
        lastRow = wsList.Cells(wsList.Rows.Count, C_FILE).End(xlUp).Row
        For r = 2 To lastRow
            If IsTicked(wsList, r) Then
                f = AsText(wsList.Cells(r, C_FILE).Value)
                If Len(f) > 0 Then
                    tickedAny = True
                    If IsSamePath(EndSep(folderPath) & f, skipPath) Then
                        leftOut = leftOut & "The reference schedule (" & f & ") is ticked; it is " & _
                               "the one being copied from, so it is left out. "
                    ElseIf Not FileExists(EndSep(folderPath) & f) Then
                        gone = gone & f & ", "
                    Else
                        c.Add f
                    End If
                End If
            End If
        Next r
    End If

    If Len(gone) > 0 Then
        leftOut = leftOut & "Ticked but no longer in the folder, so left out: " & _
               Left$(gone, Len(gone) - 2) & ". Refresh the list to tidy it up. "
    End If
    If tickedAny Then Exit Function

    isAll = True
    Set all = ScheduleFiles(folderPath)
    For i = 1 To all.Count
        If Not IsSamePath(EndSep(folderPath) & all(i), skipPath) Then c.Add all(i)
    Next i
End Function


' How a confirmation box names what is about to be done. Ticked files are
' listed by name, so what you are agreeing to is on the screen.
Private Function ScopeText(ByVal files As Collection, ByVal isAll As Boolean, _
                           ByVal leftOut As String) As String
    Dim s As String
    Dim i As Long

    If isAll Then
        s = "ALL " & files.Count & " workbook(s) in the folder, because nothing is " & _
            "ticked on ScheduleList." & vbCrLf & vbCrLf & _
            "To do only some, press No, put an x in column A of ScheduleList next " & _
            "to each one, and press this button again." & vbCrLf
    Else
        s = "the " & files.Count & " schedule(s) ticked on ScheduleList:" & vbCrLf
        For i = 1 To files.Count
            If i > 12 Then
                s = s & "    ...and " & (files.Count - 12) & " more" & vbCrLf
                Exit For
            End If
            s = s & "    " & files(i) & vbCrLf
        Next i
    End If

    If Len(leftOut) > 0 Then s = s & vbCrLf & leftOut & vbCrLf
    ScopeText = s
End Function


' The confirmation every button ends its questions with. Defaults to No when
' the run covers everything.
Private Function Confirmed(ByVal msg As String, ByVal title As String, _
                           ByVal isAll As Boolean) As Boolean
    Dim style As Long
    style = vbQuestion + vbYesNo
    If isAll Then style = vbExclamation + vbYesNo + vbDefaultButton2
    Confirmed = (MsgBox(msg, style, title) = vbYes)
End Function


Private Function IsSamePath(ByVal a As String, ByVal b As String) As Boolean
    If Len(b) = 0 Then Exit Function
    IsSamePath = (StrComp(a, b, vbTextCompare) = 0)
End Function


' A workbook already open in Excel is never opened again underneath you:
' Workbooks.Open hands back your copy, and closing it would throw away
' whatever is unsaved in it. Clicking a file name on ScheduleList opens it, so
' this is easy to run into. Logs the skip and returns True.
Private Function OpenInExcel(ByVal fileName As String, ByRef skipped As Long) As Boolean
    If Not IsWorkbookOpen(fileName) Then Exit Function
    skipped = skipped + 1
    LogLine fileName, "Skipped", "Open in Excel. Close it and run this again."
    OpenInExcel = True
End Function


' The backup folder for this run, or "" when backups are switched off.
Private Function MakeBackupDir(ByVal wsSetup As Worksheet, ByVal folderPath As String) As String
    Dim d As String
    If UCase$(Trim$(CStr(Opt(wsSetup, LBL_BACKUP, R_OPT_BACKUP).Value))) = "NO" Then Exit Function
    d = EndSep(folderPath) & "_backup " & Format$(Now, "yyyy-mm-dd hh-nn")
    On Error Resume Next
    MkDir d
    On Error GoTo 0
    MakeBackupDir = d
End Function


Private Sub BackUp(ByVal backupDir As String, ByVal folderPath As String, ByVal fileName As String)
    If Len(backupDir) = 0 Then Exit Sub
    On Error Resume Next
    FileCopy EndSep(folderPath) & fileName, EndSep(backupDir) & fileName
    On Error GoTo 0
End Sub


Private Function BackupLine(ByVal wsSetup As Worksheet) As String
    If UCase$(Trim$(CStr(Opt(wsSetup, LBL_BACKUP, R_OPT_BACKUP).Value))) = "NO" Then
        BackupLine = "Backups are switched OFF on the Setup sheet."
    Else
        BackupLine = "Each file is backed up first."
    End If
End Function


' ===========================================================================
' Rename files
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

    ' The same rule as every other button: ticked rows only, if any are.
    ' Nothing is renamed here anyway, only proposed for you to read.
    onlyTicked = (TickedCount(wsList) > 0)

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
        MsgBox """" & findText & """ is not in any of " & _
               IIf(onlyTicked, "the ticked", "the") & " file names. Nothing to do.", _
               vbInformation, "Rename files"
    Else
        MsgBox n & " file name(s) proposed in the 'New FileName' column" & _
               IIf(onlyTicked, ", for the ticked rows only.", ".") & vbCrLf & vbCrLf & _
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
' Title block - draw the block and notes box on the ticked schedules
'
' So nobody merges cells by hand: the block is fitted to each schedule's own
' columns, the notes box is merged across the full width, and both go in the
' print titles so every printed page carries them. A schedule with no block
' gets one. See BuildTitleBlocks in modSchedule for exactly what it does and
' what makes it stop.
' ===========================================================================
Public Sub DrawTitleBlocks()
    Dim wsSetup As Worksheet
    Dim folderPath As String, fileName As String
    Dim wbTgt As Workbook
    Dim files As Collection, i As Long
    Dim isAll As Boolean, leftOut As String
    Dim backupDir As String
    Dim done As Long, skipped As Long, failed As Long
    Dim started As Double
    Dim oneLog As String
    Dim changed As Long
    Const JOB As String = "Title block"

    Set wsSetup = GetSheet(ThisWorkbook, SH_SETUP)
    If wsSetup Is Nothing Then
        MsgBox "No Setup sheet. Run InstallTool first.", vbExclamation
        Exit Sub
    End If

    folderPath = SchedulesFolder()
    If Len(folderPath) = 0 Then Exit Sub

    Set files = TargetFiles(folderPath, "", isAll, leftOut)
    If files.Count = 0 Then
        MsgBox "Nothing to do." & vbCrLf & vbCrLf & leftOut, vbInformation, JOB
        Exit Sub
    End If

    If Not Confirmed("Draw the title block on " & ScopeText(files, isAll, leftOut) & vbCrLf & _
              "Rows 1-4 of each schedule sheet are redrawn to fit its columns, the " & _
              "General Notes box is merged across the full width, and both are set " & _
              "to print on every page. The notes text itself is not changed." & vbCrLf & vbCrLf & _
              "A schedule with no title block gets one: 12 rows are inserted above it " & _
              "for the block and an empty notes box." & vbCrLf & vbCrLf & _
              "A sheet whose top rows hold anything the tool does not recognise is " & _
              "left alone and named on the Log sheet." & vbCrLf & vbCrLf & _
              BackupLine(wsSetup) & vbCrLf & vbCrLf & "Continue?", JOB, isAll) Then Exit Sub

    backupDir = MakeBackupDir(wsSetup, folderPath)

    On Error GoTo Fail
    LogStart JOB
    BeginQuiet xlCalculationAutomatic
    ProgressStart files.Count, "Drawing title blocks"
    started = Timer

    For i = 1 To files.Count
        fileName = files(i)
        ProgressStep i - 1, fileName

        If Not OpenInExcel(fileName, skipped) Then
            BackUp backupDir, folderPath, fileName
            Set wbTgt = OpenQuiet(EndSep(folderPath) & fileName, False)
            If wbTgt Is Nothing Then
                failed = failed + 1
                LogLine fileName, "FAILED", "Could not open the file."
            Else
                oneLog = BuildTitleBlocks(wbTgt, changed)
                If InStr(1, oneLog, "PROBLEM:", vbTextCompare) > 0 Then
                    wbTgt.Close SaveChanges:=False
                    failed = failed + 1
                    LogLine fileName, "FAILED", oneLog & "Nothing was saved."
                ElseIf changed = 0 Then
                    wbTgt.Close SaveChanges:=False
                    skipped = skipped + 1
                    LogLine fileName, "Unchanged", "Already right. " & oneLog
                Else
                    wbTgt.Close SaveChanges:=True
                    done = done + 1
                    LogLine fileName, "OK", oneLog
                End If
                Set wbTgt = Nothing
            End If
        End If
    Next i

    ProgressDone
    EndQuiet
    ShowSummary JOB, done, skipped, failed, Timer - started, _
                "Check one print preview: page 2 onwards should carry the block and the notes." & _
                vbCrLf & vbCrLf & IIf(Len(backupDir) > 0, "Backup: " & backupDir, "")
    Exit Sub

Fail:
    Recover JOB, wbTgt
End Sub


' ===========================================================================
' Tidy sheets on the ticked schedules
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
    Dim isAll As Boolean, leftOut As String
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

    Set files = TargetFiles(folderPath, "", isAll, leftOut)
    If files.Count = 0 Then
        MsgBox "Nothing to tidy." & vbCrLf & vbCrLf & leftOut, vbInformation, "Tidy sheets"
        Exit Sub
    End If

    If Not Confirmed("Tidy the sheets in " & ScopeText(files, isAll, leftOut) & vbCrLf & _
              "Order becomes Front Cover, Revision Page, Schedule, then Metadata " & _
              "hidden at the end." & vbCrLf & vbCrLf & _
              "A workbook with one schedule sheet has it renamed to 'Schedule'. " & _
              "One with several keeps its names, since they are the only thing " & _
              "telling those sheets apart." & vbCrLf & vbCrLf & _
              "Only names, order and visibility change. No cell is touched.", _
              "Tidy sheets", isAll) Then Exit Sub

    backupDir = MakeBackupDir(wsSetup, folderPath)

    On Error GoTo Fail
    LogStart "Tidy sheets"
    BeginQuiet xlCalculationManual
    ProgressStart files.Count, "Tidying sheets"
    started = Timer

    For i = 1 To files.Count
        fileName = files(i)
        ProgressStep i - 1, fileName

        If Not OpenInExcel(fileName, skipped) Then
            BackUp backupDir, folderPath, fileName
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
                Set wbTgt = Nothing
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


' ===========================================================================
' Export ticked to PDF
'
' Unlike every other button this one only reads. The ticked rows on
' ScheduleList go to _pdf in the project folder after one confirmation, so
' re-issuing the same set is two clicks. With nothing ticked it falls back to
' asking: a whole folder, which can be anywhere, or files you pick.
' ===========================================================================
Public Sub ExportToPdf()
    Dim files As Collection, names As Collection
    Dim outFolder As String, clash As String
    Dim perSheet As Boolean, overwrite As Boolean
    Dim i As Long
    Dim fullPath As String, fileName As String
    Dim folderPath As String
    Dim isAll As Boolean, leftOut As String
    Dim fromTicks As Boolean, askedAll As Boolean
    Dim answer As VbMsgBoxResult
    Dim wbTgt As Workbook
    Dim notes As String
    Dim made As Long, madeBefore As Long
    Dim done As Long, skipped As Long, failed As Long
    Dim started As Double

    ' --- the ticked rows, when there are any ------------------------------
    folderPath = SchedulesFolder(False)
    If Len(folderPath) > 0 Then
        Set names = TargetFiles(folderPath, "", isAll, leftOut)
        fromTicks = Not isAll
    End If

    If fromTicks Then
        If names.Count = 0 Then
            MsgBox "Nothing to export." & vbCrLf & vbCrLf & leftOut, vbInformation, "Export to PDF"
            Exit Sub
        End If

        Set files = New Collection
        For i = 1 To names.Count
            files.Add EndSep(folderPath) & names(i)
        Next i

        outFolder = EndSep(folderPath) & PDF_SUBFOLDER
        answer = MsgBox("Export " & ScopeText(names, False, leftOut) & vbCrLf & _
                        "to" & vbCrLf & "    " & outFolder & vbCrLf & vbCrLf & _
                        "One PDF per workbook, replacing any PDF of the same name " & _
                        "already there. Nothing is written back to the schedules." & vbCrLf & vbCrLf & _
                        "Yes  - export" & vbCrLf & _
                        "No   - change where they go, or split them by sheet" & vbCrLf & _
                        "Cancel - stop", vbQuestion + vbYesNoCancel, "Export ticked to PDF")
        If answer = vbCancel Then Exit Sub

        If answer = vbYes Then
            outFolder = MakePdfFolder(folderPath)
            If Len(outFolder) = 0 Then Exit Sub
            perSheet = False
            overwrite = True
        Else
            outFolder = ChoosePdfFolder(folderPath)
            If Len(outFolder) = 0 Then Exit Sub
            If Not AskPdfLayout(perSheet) Then Exit Sub
            If Not AskPdfOverwrite(outFolder, overwrite) Then Exit Sub
            askedAll = True
        End If
    Else
        Set files = ChooseFilesToExport()
        If files Is Nothing Then Exit Sub          ' cancelled
        If files.Count = 0 Then
            MsgBox "No Excel files to export.", vbInformation, "Export to PDF"
            Exit Sub
        End If

        ' Found before anything is written, the way the rename button checks
        ' its whole batch first. Two files of the same name from different
        ' folders produce one PDF, and the second would look like it worked.
        clash = DuplicateName(files)
        If Len(clash) > 0 Then
            MsgBox "Two of the files chosen are both called:" & vbCrLf & vbCrLf & clash & _
                   vbCrLf & vbCrLf & "They would write the same PDF, so the second would " & _
                   "quietly replace the first. Export them in separate runs, or into " & _
                   "different folders.", vbExclamation, "Export to PDF"
            Exit Sub
        End If

        outFolder = ChoosePdfFolder(FolderOf(files(1)))
        If Len(outFolder) = 0 Then Exit Sub

        If Not AskPdfLayout(perSheet) Then Exit Sub
        If Not AskPdfOverwrite(outFolder, overwrite) Then Exit Sub
        askedAll = True
    End If

    If askedAll Then
        If MsgBox(files.Count & " workbook(s) will be exported to:" & vbCrLf & vbCrLf & _
                  outFolder & vbCrLf & vbCrLf & _
                  IIf(perSheet, "One PDF per sheet.", _
                                "One PDF per workbook, every sheet in tab order.") & vbCrLf & _
                  IIf(overwrite, "PDFs that already exist are overwritten.", _
                                 "PDFs that already exist are left alone.") & vbCrLf & vbCrLf & _
                  "Each workbook is opened read only, with its links updated and " & _
                  "recalculated, so the PDF matches what you see on screen. Nothing is " & _
                  "written back to any schedule." & vbCrLf & vbCrLf & _
                  "Continue?", vbQuestion + vbYesNo, "Export to PDF") = vbNo Then Exit Sub
    End If

    On Error GoTo Fail
    LogStart "Export to PDF"
    ' Events stay on and calculation stays automatic. See BeginQuiet.
    BeginQuiet xlCalculationAutomatic, keepEvents:=True
    ProgressStart files.Count, "Exporting to PDF"
    started = Timer

    For i = 1 To files.Count
        fullPath = files(i)
        fileName = BaseName(fullPath)
        ProgressStep i - 1, fileName

        If IsWorkbookOpen(fileName) Then
            ' Opening it again hands back the copy that is already open, and
            ' closing that would throw away whatever is unsaved in it.
            skipped = skipped + 1
            LogLine fileName, "Skipped", "Already open in Excel. Close it and run again."
        Else
            Set wbTgt = OpenQuiet(fullPath, True, 3)
            If wbTgt Is Nothing Then
                failed = failed + 1
                LogLine fileName, "FAILED", "Could not open the file."
            Else
                madeBefore = made
                notes = ExportWorkbookToPdf(wbTgt, outFolder, perSheet, overwrite, made)
                wbTgt.Close SaveChanges:=False
                Set wbTgt = Nothing

                If InStr(1, notes, "PROBLEM:", vbTextCompare) > 0 Then
                    failed = failed + 1
                    LogLine fileName, "FAILED", notes
                ElseIf made = madeBefore Then
                    skipped = skipped + 1
                    LogLine fileName, "Skipped", IIf(Len(notes) > 0, notes, _
                                                    "No sheet in it had anything to print.")
                Else
                    done = done + 1
                    LogLine fileName, "OK", notes
                End If
            End If
        End If
    Next i

    ProgressDone
    EndQuiet

    ShowSummary "Export to PDF", done, skipped, failed, Timer - started, _
                made & " PDF(s) written to " & outFolder
    Exit Sub

Fail:
    Recover "Export to PDF", wbTgt
End Sub


' Full paths of the workbooks to export. Nothing means the user backed out,
' an empty collection means they chose a folder with nothing in it.
Private Function ChooseFilesToExport() As Collection
    Dim answer As VbMsgBoxResult
    Dim folderPath As String
    Dim names As Collection
    Dim c As New Collection
    Dim i As Long

    answer = MsgBox("Nothing is ticked on ScheduleList, so which schedules do you " & _
                    "want as PDFs?" & vbCrLf & vbCrLf & _
                    "Yes  - every workbook in a folder" & vbCrLf & _
                    "No   - pick the files yourself" & vbCrLf & _
                    "Cancel - stop" & vbCrLf & vbCrLf & _
                    "Next time: put an x next to the ones you want on ScheduleList and " & _
                    "this goes straight to exporting them.", _
                    vbQuestion + vbYesNoCancel, "Export to PDF")
    If answer = vbCancel Then Exit Function

    If answer = vbNo Then
        Set ChooseFilesToExport = PickWorkbooks( _
            "Schedules to export (Ctrl or Shift to pick several)", SchedulesFolder(False))
        Exit Function
    End If

    ' Default to the project folder, but let them point somewhere else.
    folderPath = PickFolder("Folder holding the schedules to export")
    If Len(folderPath) = 0 Then Exit Function

    If Not FolderExists(folderPath) Then
        MsgBox "That folder cannot be read from Excel:" & vbCrLf & vbCrLf & folderPath & _
               vbCrLf & vbCrLf & "Pick it from the synced folder in File Explorer " & _
               "rather than from Filery or SharePoint in a browser.", _
               vbExclamation, "Export to PDF"
        Exit Function
    End If

    Set names = FolderWorkbooks(folderPath)
    For i = 1 To names.Count
        If StrComp(names(i), ThisWorkbook.Name, vbTextCompare) <> 0 Then
            c.Add EndSep(folderPath) & names(i)
        End If
    Next i
    Set ChooseFilesToExport = c
End Function


' "" means the user backed out.
Private Function ChoosePdfFolder(ByVal sourceFolder As String) As String
    Dim suggested As String
    Dim answer As VbMsgBoxResult
    Dim picked As String

    If FolderExists(sourceFolder) Then suggested = EndSep(sourceFolder) & PDF_SUBFOLDER

    If Len(suggested) > 0 Then
        answer = MsgBox("Where should the PDFs go?" & vbCrLf & vbCrLf & _
                        "Yes  - " & suggested & vbCrLf & _
                        "No   - somewhere else" & vbCrLf & _
                        "Cancel - stop", vbQuestion + vbYesNoCancel, "Export to PDF")
        If answer = vbCancel Then Exit Function

        If answer = vbYes Then
            ChoosePdfFolder = MakePdfFolder(sourceFolder)
            Exit Function
        End If
    End If

    picked = PickFolder("Where to put the PDFs")
    If Len(picked) = 0 Then Exit Function

    If Not FolderExists(picked) Then
        MsgBox "Excel cannot write to that folder:" & vbCrLf & vbCrLf & picked & _
               vbCrLf & vbCrLf & "Pick a folder on this PC, or the synced copy of " & _
               "the Filery one.", vbExclamation, "Export to PDF"
        Exit Function
    End If

    ChoosePdfFolder = EndSep(picked)
End Function


' The _pdf folder inside `sourceFolder`, created if need be, with its
' trailing separator. "" (after saying why) if it cannot be made.
Private Function MakePdfFolder(ByVal sourceFolder As String) As String
    Dim p As String

    p = EndSep(sourceFolder) & PDF_SUBFOLDER
    If Not FolderExists(p) Then
        On Error Resume Next
        MkDir p
        On Error GoTo 0
    End If
    If Not FolderExists(p) Then
        MsgBox "Could not create:" & vbCrLf & vbCrLf & p, vbExclamation, "Export to PDF"
        Exit Function
    End If
    MakePdfFolder = EndSep(p)
End Function


Private Function AskPdfLayout(ByRef perSheet As Boolean) As Boolean
    Dim answer As VbMsgBoxResult

    answer = MsgBox("One PDF per workbook, or one per sheet?" & vbCrLf & vbCrLf & _
                    "Yes  - one PDF per workbook, every sheet in it, in tab order" & vbCrLf & _
                    "No   - one PDF per sheet" & vbCrLf & _
                    "Cancel - stop" & vbCrLf & vbCrLf & _
                    "One per workbook is the usual answer: cover, revision page and " & _
                    "schedule come out as one document with continuous page numbers.", _
                    vbQuestion + vbYesNoCancel, "Export to PDF")
    If answer = vbCancel Then Exit Function

    perSheet = (answer = vbNo)
    AskPdfLayout = True
End Function


' Only worth asking when the destination already holds PDFs.
Private Function AskPdfOverwrite(ByVal outFolder As String, ByRef overwrite As Boolean) As Boolean
    Dim answer As VbMsgBoxResult

    If Not FolderHasPdfs(outFolder) Then
        overwrite = True
        AskPdfOverwrite = True
        Exit Function
    End If

    answer = MsgBox("There are already PDFs in:" & vbCrLf & vbCrLf & outFolder & vbCrLf & vbCrLf & _
                    "Yes  - overwrite the ones this run produces" & vbCrLf & _
                    "No   - leave those alone and skip them" & vbCrLf & _
                    "Cancel - stop", vbQuestion + vbYesNoCancel, "Export to PDF")
    If answer = vbCancel Then Exit Function

    overwrite = (answer = vbYes)
    AskPdfOverwrite = True
End Function


Private Function FolderHasPdfs(ByVal folderPath As String) As Boolean
    Dim f As Object

    If Not FolderExists(folderPath) Then Exit Function

    On Error Resume Next
    For Each f In Fso.GetFolder(folderPath).files
        If LCase$(Fso.GetExtensionName(f.Name)) = "pdf" Then
            FolderHasPdfs = True
            Exit For
        End If
    Next f
    On Error GoTo 0
End Function


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
    Dim rev As String, descr As String, gaps As String
    Dim found As Boolean
    Dim chk As String

    wsList.Cells(r, C_FILE).Value = fileName
    wsList.Cells(r, 3).Value = ReadMeta(wb, "ScheduleName")
    wsList.Cells(r, C_PROJNAME).Value = ReadMeta(wb, "Project Name")
    wsList.Cells(r, C_PROJNO).Value = ReadMeta(wb, "Project Number")
    wsList.Cells(r, C_CLIENT).Value = ReadMeta(wb, "Client")
    wsList.Cells(r, 7).Value = ReadMeta(wb, "DocumentType")
    wsList.Cells(r, C_REV).Value = ReadMeta(wb, "Revision")
    wsList.Cells(r, C_DATE).Value = ReadMeta(wb, "Date")
    wsList.Cells(r, 10).Value = ReadMeta(wb, "Prepared by")
    wsList.Cells(r, 11).Value = ReadMeta(wb, "Checked by")
    wsList.Cells(r, 12).Value = ReadMeta(wb, "Approved by")
    wsList.Cells(r, 13).Value = ReadMeta(wb, "DocumentNumber")
    wsList.Cells(r, 14).Value = ReadMeta(wb, "Suitability Status")
    wsList.Cells(r, 15).Value = ReadMeta(wb, "Suitability Description")

    chk = CheckFile(wb)

    ' The description sits only on the revision table, so it is read from
    ' there: the row of the revision the title block shows.
    rev = AsText(wsList.Cells(r, C_REV).Value)
    descr = RevisionDescription(wb, rev, found)
    wsList.Cells(r, C_REVDESC).Value = descr
    If Len(rev) > 0 And found And Len(descr) = 0 Then
        chk = chk & "Revision " & rev & " has no description. "
    ElseIf Len(rev) > 0 And Not found And Left$(chk, Len(NOT_READY)) <> NOT_READY Then
        chk = chk & "Revision " & rev & " is not in RevisionTable. "
    End If

    wsList.Cells(r, C_CLASS).Value = ClassificationOf(wb, gaps)
    chk = chk & gaps

    wsList.Cells(r, C_FILECHK).Value = chk
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
        CheckFile = NOT_READY & "SET UP - no Metadata sheet. Tick it and press " & _
                    "'Set up / repair links'. "
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


' True for a row the project-wide comparisons apply to: a schedule that has
' been set up. One that has not would only be reported as different from
' everything, which says nothing its own check does not already.
Private Function IsComparable(ByVal wsList As Worksheet, ByVal r As Long) As Boolean
    IsComparable = (Left$(AsText(wsList.Cells(r, C_FILECHK).Value), Len(NOT_READY)) <> NOT_READY)
End Function


' Builds the visible Checks column for every row, from the values on the row
' plus the file-derived half. Recomputed every refresh, so changing the
' project name or the classification on Setup updates the report even for
' files that were reused.
Private Sub ComposeChecks(ByVal wsList As Worksheet, ByVal lastRow As Long, _
                          ByVal mpiName As String, ByVal mpiNo As String, _
                          ByVal mpiClient As String, ByVal wantClass As String)
    Dim r As Long
    Dim out As String, v As String
    Dim expected As String

    If Len(wantClass) > 0 Then
        If wantClass = "NONE" Then expected = "None" Else expected = NormaliseMarking(wantClass)
    End If

    For r = 2 To lastRow
        out = AsText(wsList.Cells(r, C_FILECHK).Value)

        If IsComparable(wsList, r) Then
            v = AsText(wsList.Cells(r, C_PROJNAME).Value)
            If Len(mpiName) > 0 And StrComp(v, mpiName, vbTextCompare) <> 0 Then _
                out = out & "Project Name is '" & v & "', not '" & mpiName & "'. "

            v = AsText(wsList.Cells(r, C_PROJNO).Value)
            If Len(mpiNo) > 0 And StrComp(v, mpiNo, vbTextCompare) <> 0 Then _
                out = out & "Project Number is '" & v & "', not '" & mpiNo & "'. "

            v = AsText(wsList.Cells(r, C_CLIENT).Value)
            If Len(mpiClient) > 0 And StrComp(v, mpiClient, vbTextCompare) <> 0 Then _
                out = out & "Client is '" & v & "', not '" & mpiClient & "'. "

            v = AsText(wsList.Cells(r, C_CLASS).Value)
            If Len(expected) > 0 And StrComp(v, expected, vbTextCompare) <> 0 Then _
                out = out & "Marked " & v & ", Setup says " & expected & ". "
        End If

        If Len(Trim$(out)) = 0 Then out = "OK"
        wsList.Cells(r, C_CHECKS).Value = Trim$(out)
    Next r
End Sub


' Flags any schedule whose revision or date differs from most of the others.
' This is the "title block says P03 while everything else says P04" check.
' The classification gets the same treatment when the Setup sheet does not
' name one to check against, which is how one CONFIDENTIAL among 25 OFFICIAL
' still gets noticed.
Private Sub FlagOddOnesOut(ByVal wsList As Worksheet, ByVal lastRow As Long, _
                           ByVal checkClass As Boolean)
    Dim modeRev As String, modeDate As String, modeClass As String
    Dim r As Long

    If lastRow < 3 Then Exit Sub

    modeRev = MostCommon(wsList, C_REV, lastRow)
    modeDate = MostCommon(wsList, C_DATE, lastRow)
    If checkClass Then modeClass = MostCommon(wsList, C_CLASS, lastRow)

    For r = 2 To lastRow
        If IsComparable(wsList, r) Then
            If Len(modeRev) > 0 Then
                If StrComp(AsText(wsList.Cells(r, C_REV).Value), modeRev, vbTextCompare) <> 0 Then
                    AppendCheck wsList.Cells(r, C_CHECKS), _
                                "Revision differs from most schedules (" & modeRev & "). "
                End If
            End If
            If Len(modeDate) > 0 Then
                If StrComp(AsText(wsList.Cells(r, C_DATE).Value), modeDate, vbTextCompare) <> 0 Then
                    AppendCheck wsList.Cells(r, C_CHECKS), _
                                "Date differs from most schedules (" & modeDate & "). "
                End If
            End If
            If Len(modeClass) > 0 Then
                If StrComp(AsText(wsList.Cells(r, C_CLASS).Value), modeClass, vbTextCompare) <> 0 Then
                    AppendCheck wsList.Cells(r, C_CHECKS), _
                                "Marked differently from most schedules (" & modeClass & "). "
                End If
            End If
        End If
    Next r
End Sub


' The value most of the set-up schedules share, or "" when no value is held
' by more than half of them.
Private Function MostCommon(ByVal ws As Worksheet, ByVal col As Long, ByVal lastRow As Long) As String
    Dim d As Object, r As Long, v As String
    Dim best As String, bestN As Long, n As Long
    Dim k As Variant

    Set d = CreateObject("Scripting.Dictionary")
    For r = 2 To lastRow
        If IsComparable(ws, r) Then
            n = n + 1
            v = AsText(ws.Cells(r, col).Value)
            If Len(v) > 0 Then d(v) = d(v) + 1
        End If
    Next r
    For Each k In d.Keys
        If d(k) > bestN Then
            bestN = d(k)
            best = CStr(k)
        End If
    Next k
    ' Only meaningful if it really is the majority.
    If n > 0 And bestN * 2 > n Then MostCommon = best
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

    ' Rows written under an older column layout would be put back under the
    ' wrong headings, so they are not kept at all. The next refresh is then a
    ' full one, which is what a new layout needs anyway.
    If Not ListLayoutIsCurrent(wsList) Then Exit Function

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


Private Function ListLayoutIsCurrent(ByVal wsList As Worksheet) As Boolean
    Dim h As Variant
    Dim i As Long

    h = Split(LIST_HEADERS, "|")
    For i = 0 To UBound(h)
        If StrComp(AsText(wsList.Cells(1, i + 1).Value), CStr(h(i)), vbTextCompare) <> 0 Then Exit Function
    Next i
    ListLayoutIsCurrent = True
End Function


' The file name as a link to the file, so a schedule can be opened from the
' list to check something. A file you open this way is left alone by every
' button until you close it.
Private Sub LinkFileCell(ByVal cell As Range, ByVal fullPath As String)
    Dim txt As String
    txt = AsText(cell.Value)
    If Len(txt) = 0 Then Exit Sub
    On Error Resume Next
    cell.Hyperlinks.Delete
    cell.Worksheet.Hyperlinks.Add Anchor:=cell, Address:=fullPath, TextToDisplay:=txt
    On Error GoTo 0
End Sub


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


' Clears the typed revision once it is in the file. The tick stays: it is the
' selection for the next button too, PDFs most often.
Private Sub ClearRowEntry(ByVal wsList As Worksheet, ByVal r As Long)
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
' updateLinks stays 0 for every button that edits a schedule: those read what
' is saved in the file. The PDF export passes 3, because a document you issue
' has to match what you see on screen.
Private Function OpenQuiet(ByVal fullPath As String, ByVal readOnlyMode As Boolean, _
                           Optional ByVal updateLinks As Long = 0) As Workbook
    Dim wb As Workbook
    On Error Resume Next
    Set wb = Workbooks.Open(fileName:=fullPath, ReadOnly:=readOnlyMode, UpdateLinks:=updateLinks)
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
    Dim lbls As Variant, rowsAt As Variant
    Dim saved() As Variant
    Dim cell As Range
    Dim i As Long

    ' Option values are taken out by their label before anything is cleared,
    ' and put back on the row the current layout gives that label. Rows move
    ' when an option is added, and without this a value stays where it was
    ' and is read under whatever label lands next to it.
    lbls = Array(LBL_FOLDER, LBL_BACKUP, LBL_AUTO, LBL_FULL, LBL_HFSRC, LBL_HFIMG, _
                 LBL_HFSCALE, LBL_CLASS)
    rowsAt = Array(R_OPT_FOLDER, R_OPT_BACKUP, R_OPT_AUTO, R_OPT_FULL, R_OPT_HFSRC, _
                   R_OPT_HFIMG, R_OPT_HFSCALE, R_OPT_CLASS)
    ReDim saved(0 To UBound(lbls))
    For i = 0 To UBound(lbls)
        Set cell = FindLabel(ws, CStr(lbls(i)), R_OPT_FOLDER, R_REV_FIRST - 2)
        If Not cell Is Nothing Then saved(i) = cell.Offset(0, 1).Value
    Next i

    ' Start from a clean slate. Formats AND text, because the row a label sits
    ' on can move between versions and a leftover line is worse than no line.
    ' Left alone: B1:B4 (the project), the New revision values in column B,
    ' the PROJECT FIELDS block in columns A and B (schedules link to those
    ' cells by address, and the names live in column A), column F (the
    ' suitability codes), and every other sheet in the workbook.
    ws.Range("A1:F60").ClearFormats
    ws.Range("A5:A" & (R_FLD_FIRST - 1)).ClearContents
    ws.Range(ws.Cells(R_OPT_FOLDER, 2), ws.Cells(R_OPT_LIST, 2)).ClearContents
    On Error Resume Next
    ws.Range(ws.Cells(R_OPT_FOLDER, 2), ws.Cells(R_OPT_LIST, 2)).Validation.Delete
    On Error GoTo 0
    ws.Range("C1:C60").ClearContents
    ws.Range("H1:K40").ClearFormats
    ws.Range("H1:K40").ClearContents
    ws.Range("A1:A" & R_REV_FIRST + 6).Font.Color = CLR_LABEL

    ' Row heights back to standard. An earlier version shrank rows 2 and 5 as
    ' spacers, which squashed the suitability codes sharing those rows in F.
    ws.Rows("1:40").RowHeight = ws.StandardHeight

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
    ws.Cells(R_OPT_FOLDER, 1).Value = LBL_FOLDER
    ws.Cells(R_OPT_BACKUP, 1).Value = LBL_BACKUP
    ws.Cells(R_OPT_AUTO, 1).Value = LBL_AUTO
    ws.Cells(R_OPT_FULL, 1).Value = LBL_FULL
    ws.Cells(R_OPT_HFSRC, 1).Value = LBL_HFSRC
    ws.Cells(R_OPT_HFIMG, 1).Value = LBL_HFIMG
    ws.Cells(R_OPT_HFSCALE, 1).Value = LBL_HFSCALE
    ws.Cells(R_OPT_CLASS, 1).Value = LBL_CLASS
    ws.Cells(R_OPT_SETUP, 1).Value = LBL_SETUP
    ws.Cells(R_OPT_LIST, 1).Value = LBL_LIST

    For i = 0 To UBound(lbls)
        ws.Cells(rowsAt(i), 2).Value = saved(i)
    Next i

    InputCell ws.Range(ws.Cells(R_OPT_FOLDER, 2), ws.Cells(R_OPT_CLASS, 2))
    ReadOnlyCell ws.Cells(R_OPT_SETUP, 2)
    ReadOnlyCell ws.Cells(R_OPT_LIST, 2)

    ' A value that is not a valid answer, from a layout older than the
    ' save-and-restore above, goes back to the default rather than being
    ' left to be read wrong.
    NormaliseYesNo ws.Cells(R_OPT_BACKUP, 2), "Yes"
    NormaliseYesNo ws.Cells(R_OPT_AUTO, 2), "No"
    NormaliseYesNo ws.Cells(R_OPT_FULL, 2), "No"

    If IsNumeric(ws.Cells(R_OPT_HFSRC, 2).Value) Then ws.Cells(R_OPT_HFSRC, 2).ClearContents
    If Not IsNumeric(ws.Cells(R_OPT_HFSCALE, 2).Value) Then ws.Cells(R_OPT_HFSCALE, 2).Value = 20
    If ws.Cells(R_OPT_HFSCALE, 2).Value <= 0 Then ws.Cells(R_OPT_HFSCALE, 2).Value = 20
    If Len(AsText(ws.Cells(R_OPT_CLASS, 2).Value)) > 0 Then
        If UCase$(AsText(ws.Cells(R_OPT_CLASS, 2).Value)) <> "NONE" And _
           Not IsKnownMarking(AsText(ws.Cells(R_OPT_CLASS, 2).Value)) Then
            ws.Cells(R_OPT_CLASS, 2).ClearContents
        End If
    End If

    YesNoList ws.Cells(R_OPT_BACKUP, 2)
    YesNoList ws.Cells(R_OPT_AUTO, 2)
    YesNoList ws.Cells(R_OPT_FULL, 2)
    ClassList ws.Cells(R_OPT_CLASS, 2)

    Note ws.Cells(R_OPT_FOLDER, 3), "blank = the folder this file is saved in"
    Note ws.Cells(R_OPT_BACKUP, 3), "copies every file into a timestamped folder first"
    Note ws.Cells(R_OPT_AUTO, 3), "reads every schedule when this file is opened"
    Note ws.Cells(R_OPT_FULL, 3), "No = only reopen files that changed since last time"
    Note ws.Cells(R_OPT_HFSRC, 3), "the schedule that is set up correctly; copied from by three buttons"
    Note ws.Cells(R_OPT_HFIMG, 3), "logo for the top-right of the header; blank = ask, or leave for none"
    Note ws.Cells(R_OPT_HFSCALE, 3), "size of that logo as a percentage of the image's own size"
    Note ws.Cells(R_OPT_CLASS, 3), "what every schedule should be marked; ScheduleList flags any that are not"

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

    ' --- The buttons, in three groups, with notes under them ----------------
    ' Columns H, I and J are one group each; BuildButtons puts the buttons in
    ' the rows these headings leave free.
    SectionHeader ws, 1, "EVERY ISSUE", 8
    SectionHeader ws, 1, "ADDING A SCHEDULE", 9
    SectionHeader ws, 1, "ONE JOB AT A TIME", 10

    Note ws.Range("I5"), "1. Put the file in the folder."
    Note ws.Range("I6"), "2. Refresh schedule list."
    Note ws.Range("I7"), "3. Put an x next to it."
    Note ws.Range("I8"), "4. Press the button above."
    Note ws.Range("I10"), "Does cover, links, title block, headers"
    Note ws.Range("I11"), "and order in one go, ticked rows only."

    Note ws.Range("H16"), "WHICH SCHEDULES: every button works on the rows with an x in column A of ScheduleList."
    Note ws.Range("H17"), "Nothing ticked means all of them, and the button says ALL and defaults to No first."
    Note ws.Range("H18"), "Ticks stay until you delete them, so 'Add revision' then 'Export ticked to PDF' issues the same set."
    Note ws.Range("H20"), "To reissue: tick the rows, fill the blue 'New ...' columns (blanks come from NEW REVISION on the left),"
    Note ws.Range("H21"), "press 'Add revision to ticked', then 'Export ticked to PDF'."
    Note ws.Range("H23"), "Security marking: choose it in 'Security classification', tick, press 'Set classification'."
    Note ws.Range("H24"), "The Classification column on ScheduleList shows what each file is really marked, footers included."
    Note ws.Range("H26"), "'Copy headers & footers' and 'Copy cover & revision page' copy from the Reference schedule."
    Note ws.Range("H27"), "Click a file name on ScheduleList to open it. Buttons leave a file you have open alone."
    Note ws.Range("H28"), "'Title block' fits the block and notes to each schedule's columns. No merging by hand."
    Note ws.Range("H29"), "Cells shaded yellow are the ones you fill in. Progress shows in the status bar, bottom left."
    Note ws.Range("H30"), "Every run writes a line per file to the Log sheet, then a summary."
    Note ws.Range("H31"), "Schedule tool version " & TOOL_VERSION

    ' --- Layout ------------------------------------------------------------
    ws.Columns("A").ColumnWidth = 24
    ws.Columns("B").ColumnWidth = 34
    ws.Columns("C").ColumnWidth = 46
    ws.Columns("D:E").ColumnWidth = 3
    ws.Columns("F").ColumnWidth = 34
    ws.Columns("G").ColumnWidth = 3
    ws.Columns("H:J").ColumnWidth = 34
    ws.Range("A1:J60").VerticalAlignment = xlCenter
End Sub


' A section title: bold, coloured, with a rule underneath.
Private Sub SectionHeader(ByVal ws As Worksheet, ByVal r As Long, ByVal txt As String, _
                          Optional ByVal firstCol As Long = 1, Optional ByVal lastCol As Long = 3)
    Dim rng As Range

    If firstCol > 1 Then lastCol = firstCol     ' a heading over one column
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


' The classification dropdown. Only a warning on anything else, so a marking
' the tool knows but does not offer (SECRET, RESTRICTED) can still be typed;
' InstallTool clears one it does not know at all.
Private Sub ClassList(ByVal rng As Range)
    On Error Resume Next
    rng.Validation.Delete
    rng.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertWarning, _
        Operator:=xlBetween, Formula1:=CLASS_CHOICES
    rng.Validation.IgnoreBlank = True
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

    h = Split(LIST_HEADERS, "|")
    For i = 0 To UBound(h)
        ws.Cells(1, i + 1).Value = h(i)
    Next i

    ' Own the header row completely. It inherited white text from the table
    ' style this sheet used to be, which was unreadable on a light fill.
    ws.Rows(1).ClearFormats
    ws.Rows(1).Font.Bold = True
    ws.Rows(1).Font.Color = RGB(0, 0, 0)
    ws.Range(ws.Cells(1, 1), ws.Cells(1, C_CHECKS)).Interior.Color = RGB(230, 230, 230)
    ws.Cells(1, C_PICK).Interior.Color = RGB(255, 230, 150)
    ws.Range(ws.Cells(1, C_NEW_FIRST), ws.Cells(1, C_NEWNAME)).Interior.Color = RGB(214, 232, 255)

    SetColumnState ws

    HeaderNote ws.Cells(1, C_PICK), _
        "Put an x here next to every schedule a button should work on." & vbCrLf & vbCrLf & _
        "Every button on the Setup sheet that changes a schedule, and the PDF " & _
        "export, works on the ticked rows only. With nothing ticked it offers to " & _
        "do ALL of them, and says so first." & vbCrLf & vbCrLf & _
        "Ticks stay until you delete them, so 'Add revision to ticked' then " & _
        "'Export ticked to PDF' issues the same set."
    HeaderNote ws.Cells(1, C_REVDESC), _
        "The Description on the revision table row of the revision shown in " & _
        "the Revision column, i.e. the latest one."
    HeaderNote ws.Cells(1, C_CLASS), _
        "The security marking the file actually carries, read from the headers " & _
        "and footers of every printed sheet and from any cell on the cover or " & _
        "revision page that says nothing but a marking." & vbCrLf & vbCrLf & _
        "'A + B' means the sheets disagree. Checks names any sheet with no marking " & _
        "at all, and any file that does not match 'Security classification' on " & _
        "the Setup sheet (or most of the others, when that is blank)."
    HeaderNote ws.Cells(1, C_NEWNAME), _
        "The file name to rename this schedule to." & vbCrLf & vbCrLf & _
        "Press 'Rename files' once to fill this column in by find and replace, " & _
        "check what it proposes, then press it again to do the renaming. Or just " & _
        "type a name here yourself." & vbCrLf & vbCrLf & _
        "Leave the extension off and the current one is kept."
    HeaderNote ws.Cells(1, C_NEW_FIRST), _
        "The revision line to add to this schedule." & vbCrLf & vbCrLf & _
        "Fill in as much or as little as you like: blank cells fall back to the " & _
        "'New revision' block on the Setup sheet. Rows without an x in column A " & _
        "are ignored." & vbCrLf & vbCrLf & _
        "The new line goes in directly under the last revision in the table."

    ' The filter is rebuilt each time so it always spans the current columns.
    On Error Resume Next
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Range(ws.Cells(1, 1), ws.Cells(1, C_CHECKS)).AutoFilter
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

    ws.Columns(C_DATE).NumberFormat = "dd/mm/yyyy"            ' Date
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
    If ws.Columns(C_REVDESC).ColumnWidth > 45 Then ws.Columns(C_REVDESC).ColumnWidth = 45
    If ws.Columns(C_CLASS).ColumnWidth > 30 Then ws.Columns(C_CLASS).ColumnWidth = 30
    If ws.Columns(C_CHECKS).ColumnWidth > 60 Then ws.Columns(C_CHECKS).ColumnWidth = 60
    If ws.Columns(C_NEWNAME).ColumnWidth > 60 Then ws.Columns(C_NEWNAME).ColumnWidth = 60
    ws.Columns(C_PICK).ColumnWidth = 7

    ' Header row and the tick and file name columns stay in view while
    ' scrolling across to Checks and the New columns.
    FreezeListPanes ws

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


' Freezing works on a window, not a sheet, so the list is made the active
' sheet first, and if that did not happen nothing is touched: the active
' window could be some other workbook of yours. The refresh ends on this
' sheet anyway.
Private Sub FreezeListPanes(ByVal ws As Worksheet)
    On Error GoTo Done
    ws.Parent.Activate
    ws.Activate
    If Not ActiveSheet Is ws Then GoTo Done

    With ActiveWindow
        If Not (.FreezePanes And .SplitRow = 1 And .SplitColumn = C_FILE) Then
            .FreezePanes = False
            .ScrollRow = 1
            .ScrollColumn = 1
            .SplitColumn = C_FILE
            .SplitRow = 1
            .FreezePanes = True
        End If
    End With
Done:
End Sub


Private Sub BuildButtons(ByVal ws As Worksheet)
    Dim i As Long

    For i = ws.Buttons.Count To 1 Step -1
        ws.Buttons(i).Delete
    Next i

    ' Three groups, under the headings BuildSetupSheet puts in row 1. Read
    ' left to right they are how often you need them: every issue, when a
    ' schedule joins the project, and one-off fixes.
    ButtonAt ws, "H", 2, "RefreshScheduleList", "Refresh schedule list"
    ButtonAt ws, "H", 4, "AddRevisionToTicked", "Add revision to ticked"
    ButtonAt ws, "H", 6, "ExportToPdf", "Export ticked to PDF"

    ButtonAt ws, "I", 2, "SetupNewSchedules", "Set up new schedule(s)"

    ButtonAt ws, "J", 2, "SetupProject", "Set up / repair links"
    ButtonAt ws, "J", 4, "SetClassification", "Set classification"
    ButtonAt ws, "J", 6, "CopyHeadersFooters", "Copy headers && footers"
    ButtonAt ws, "J", 8, "CopyCommonSheets", "Copy cover && revision page"
    ButtonAt ws, "J", 10, "DrawTitleBlocks", "Title block"
    ButtonAt ws, "J", 12, "TidySheets", "Tidy sheets"
    ButtonAt ws, "J", 14, "RenameFiles", "Rename files"
End Sub


' A button filling two rows of one column, inset a little so neighbours do
' not touch.
Private Sub ButtonAt(ByVal ws As Worksheet, ByVal col As String, ByVal r As Long, _
                     ByVal macroName As String, ByVal label As String)
    Dim topCell As Range
    Dim b As Object
    Dim h As Double

    Set topCell = ws.Range(col & r)
    h = topCell.Height + topCell.Offset(1, 0).Height - 4
    Set b = ws.Buttons.Add(topCell.Left + 2, topCell.Top + 2, topCell.Width - 10, h)
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


' keepEvents is for the PDF export. A schedule opened with events off, under
' manual calculation and with screen updating off, exports with a line struck
' through every calculated value; the same file exported by hand is clean.
' Each of those settings is harmless on its own, the combination is not.
Private Sub BeginQuiet(ByVal calcMode As XlCalculation, _
                       Optional ByVal keepEvents As Boolean = False)
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    If Not keepEvents Then Application.EnableEvents = False
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
