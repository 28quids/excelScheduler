Attribute VB_Name = "modExport"
Option Explicit

' ---------------------------------------------------------------------------
' modExport - export schedules to PDF.
'
'   ExportToPdf   pick a folder or pick the files yourself, then write one PDF
'                 per workbook covering every sheet in it.
'
' NOT WIRED INTO THE TOOL YET. There is no button for it and InstallTool knows
' nothing about it: import the module and run ExportToPdf from Alt+F8.
'
' It uses modUtil and nothing else, so it can be imported on its own and
' dropped again without touching anything that works today. The logging and
' quiet-mode helpers at the bottom are local copies of modMain's, which are
' Private to that module and so cannot be called from here. Wiring this in
' means adding a button and either moving these entry points into modMain or
' making modMain's copies Public, then deleting the copies below.
' ---------------------------------------------------------------------------

' Default destination, created inside the folder the workbooks came from.
Private Const PDF_SUBFOLDER As String = "_pdf"

' Windows refuses a path longer than this, and a half-written PDF is worse
' than a refusal, so a name that would go over is reported instead.
Private Const MAX_PATH_LEN As Long = 255

Private Const EXPORT_LOG_RUNS As Long = 5

Private mLogRow As Long
Private mCalcSaved As XlCalculation


' ===========================================================================
' The one entry point
' ===========================================================================
Public Sub ExportToPdf()
    Dim files As Collection
    Dim outFolder As String
    Dim perSheet As Boolean, overwrite As Boolean
    Dim i As Long
    Dim fullPath As String, fileName As String
    Dim wb As Workbook
    Dim notes As String
    Dim made As Long, madeBefore As Long
    Dim done As Long, skipped As Long, failed As Long
    Dim clash As String
    Dim started As Double

    Set files = ChooseFiles()
    If files Is Nothing Then Exit Sub          ' cancelled
    If files.Count = 0 Then
        MsgBox "No Excel files to export.", vbInformation, "Export to PDF"
        Exit Sub
    End If

    ' Found before anything is written, the way the rename button checks its
    ' whole batch first. Two files of the same name from different folders
    ' produce one PDF, and the second run over it would look like it worked.
    clash = DuplicateName(files)
    If Len(clash) > 0 Then
        MsgBox "Two of the files chosen are both called:" & vbCrLf & vbCrLf & clash & _
               vbCrLf & vbCrLf & "They would write the same PDF, so the second would " & _
               "quietly replace the first. Export them in separate runs, or into " & _
               "different folders.", vbExclamation, "Export to PDF"
        Exit Sub
    End If

    outFolder = ChooseOutputFolder(FolderOf(files(1)))
    If Len(outFolder) = 0 Then Exit Sub

    If Not AskLayout(perSheet) Then Exit Sub
    If Not AskOverwrite(outFolder, overwrite) Then Exit Sub

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

    On Error GoTo Fail
    ExportLogStart "Export to PDF"
    BeginQuiet
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
            ExportLogLine fileName, "Skipped", "Already open in Excel. Close it and run again."
        Else
            Set wb = OpenQuiet(fullPath)
            If wb Is Nothing Then
                failed = failed + 1
                ExportLogLine fileName, "FAILED", "Could not open the file."
            Else
                madeBefore = made
                notes = ExportWorkbook(wb, outFolder, perSheet, overwrite, made)
                wb.Close SaveChanges:=False
                Set wb = Nothing

                If InStr(1, notes, "PROBLEM:", vbTextCompare) > 0 Then
                    failed = failed + 1
                    ExportLogLine fileName, "FAILED", notes
                ElseIf made = madeBefore Then
                    skipped = skipped + 1
                    ExportLogLine fileName, "Skipped", IIf(Len(notes) > 0, notes, _
                                                          "No sheet in it had anything to print.")
                Else
                    done = done + 1
                    ExportLogLine fileName, "OK", notes
                End If
            End If
        End If
    Next i

    ProgressDone
    EndQuiet

    ExportSummary "Export to PDF", done, skipped, failed, Timer - started, _
                  made & " PDF(s) written to " & outFolder
    Exit Sub

Fail:
    ExportRecover "Export to PDF", wb
End Sub


' ===========================================================================
' What to export, where to put it, how to lay it out
' ===========================================================================

' Returns full paths. Nothing means the user backed out, an empty collection
' means they chose a folder with nothing in it.
Private Function ChooseFiles() As Collection
    Dim answer As VbMsgBoxResult
    Dim folderPath As String
    Dim names As Collection
    Dim c As New Collection
    Dim i As Long

    answer = MsgBox("Which schedules do you want as PDFs?" & vbCrLf & vbCrLf & _
                    "Yes  - every workbook in a folder" & vbCrLf & _
                    "No   - pick the files yourself" & vbCrLf & _
                    "Cancel - stop", vbQuestion + vbYesNoCancel, "Export to PDF")
    If answer = vbCancel Then Exit Function

    If answer = vbNo Then
        Set ChooseFiles = PickWorkbooks("Schedules to export (Ctrl or Shift to pick several)")
        Exit Function
    End If

    folderPath = PickFolder("Folder holding the schedules to export")
    If Len(folderPath) = 0 Then Exit Function

    If Not FolderExists(folderPath) Then
        MsgBox "That folder cannot be read from Excel:" & vbCrLf & vbCrLf & folderPath & _
               vbCrLf & vbCrLf & "Pick it from the synced folder in File Explorer " & _
               "rather than from Filery or SharePoint in a browser.", _
               vbExclamation, "Export to PDF"
        Exit Function
    End If

    ' This workbook is the MPI, not a schedule.
    Set names = FolderWorkbooks(folderPath)
    For i = 1 To names.Count
        If StrComp(names(i), ThisWorkbook.Name, vbTextCompare) <> 0 Then
            c.Add EndSep(folderPath) & names(i)
        End If
    Next i
    Set ChooseFiles = c
End Function


Private Function PickWorkbooks(ByVal promptText As String) As Collection
    Dim fd As FileDialog
    Dim c As New Collection
    Dim i As Long

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = promptText
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "Excel files", "*.xls;*.xlsx;*.xlsm;*.xlsb"
        If FolderExists(ThisWorkbook.Path) Then .InitialFileName = EndSep(ThisWorkbook.Path)
        If .Show <> -1 Then Exit Function        ' cancelled, returns Nothing
        For i = 1 To .SelectedItems.Count
            c.Add .SelectedItems(i)
        Next i
    End With

    Set PickWorkbooks = c
End Function


' "" means the user backed out.
Private Function ChooseOutputFolder(ByVal sourceFolder As String) As String
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
            If Not FolderExists(suggested) Then
                On Error Resume Next
                MkDir suggested
                On Error GoTo 0
            End If
            If Not FolderExists(suggested) Then
                MsgBox "Could not create:" & vbCrLf & vbCrLf & suggested, _
                       vbExclamation, "Export to PDF"
                Exit Function
            End If
            ChooseOutputFolder = EndSep(suggested)
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

    ChooseOutputFolder = EndSep(picked)
End Function


Private Function AskLayout(ByRef perSheet As Boolean) As Boolean
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
    AskLayout = True
End Function


' Only worth asking when the destination already holds PDFs.
Private Function AskOverwrite(ByVal outFolder As String, ByRef overwrite As Boolean) As Boolean
    Dim answer As VbMsgBoxResult

    If Not FolderHasPdfs(outFolder) Then
        overwrite = True
        AskOverwrite = True
        Exit Function
    End If

    answer = MsgBox("There are already PDFs in:" & vbCrLf & vbCrLf & outFolder & vbCrLf & vbCrLf & _
                    "Yes  - overwrite the ones this run produces" & vbCrLf & _
                    "No   - leave those alone and skip them" & vbCrLf & _
                    "Cancel - stop", vbQuestion + vbYesNoCancel, "Export to PDF")
    If answer = vbCancel Then Exit Function

    overwrite = (answer = vbYes)
    AskOverwrite = True
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


' ===========================================================================
' Exporting one workbook
' ===========================================================================

' Returns the note for the log. A note containing "PROBLEM:" is a failure.
' `made` is raised by one for every PDF actually written, which is how the
' caller tells an export from a workbook that had nothing to give.
Private Function ExportWorkbook(ByVal wb As Workbook, ByVal outFolder As String, _
                                ByVal perSheet As Boolean, ByVal overwrite As Boolean, _
                                ByRef made As Long) As String
    Dim sheetList As Collection
    Dim stem As String, outPath As String
    Dim res As String, notes As String
    Dim madeHere As Long, skippedExisting As Long
    Dim i As Long

    Set sheetList = PrintableSheets(wb)
    If sheetList.Count = 0 Then Exit Function

    stem = NameWithoutExtension(wb.Name)

    If Not LooksLikeSchedule(wb) Then notes = "No Revision Page, exported anyway. "

    If perSheet Then
        For i = 1 To sheetList.Count
            outPath = EndSep(outFolder) & SafeName(stem & " - " & sheetList(i)) & ".pdf"
            If FileExists(outPath) And Not overwrite Then
                skippedExisting = skippedExisting + 1
            Else
                res = WritePdf(wb, Array(sheetList(i)), outPath)
                If Len(res) > 0 Then
                    ExportWorkbook = notes & res
                    Exit Function
                End If
                made = made + 1
                madeHere = madeHere + 1
            End If
        Next i

        If madeHere > 0 Then notes = notes & madeHere & " sheet(s) exported. "
        If skippedExisting > 0 Then
            notes = notes & skippedExisting & " sheet(s) already had a PDF and were skipped. "
        End If
    Else
        outPath = EndSep(outFolder) & SafeName(stem) & ".pdf"
        If FileExists(outPath) And Not overwrite Then
            ExportWorkbook = notes & "A PDF of that name is already there. "
            Exit Function
        End If

        If CoversAllVisible(wb, sheetList) Then
            res = WritePdf(wb, Empty, outPath)
        Else
            res = WritePdf(wb, ToArray(sheetList), outPath)
        End If
        If Len(res) > 0 Then
            ExportWorkbook = notes & res
            Exit Function
        End If
        made = made + 1
        notes = notes & sheetList.Count & " sheet(s): " & JoinCollection(sheetList, ", ") & ". "
    End If

    ExportWorkbook = notes
End Function


' The sheets worth printing, in tab order.
'
' Metadata is working data, never issued, so it is left out by name as well as
' by being hidden: a workbook that has not been through Tidy sheets may still
' have it visible.
Private Function PrintableSheets(ByVal wb As Workbook) As Collection
    Dim c As New Collection
    Dim ws As Worksheet

    Set PrintableSheets = c
    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible Then
            If LCase$(Trim$(ws.Name)) <> LCase$(SH_META) Then
                If HasSomethingToPrint(ws) Then c.Add ws.Name
            End If
        End If
    Next ws
End Function


' An empty sheet in the selection makes the whole export fail, so they are
' filtered out rather than left to break the file that would have been fine.
Private Function HasSomethingToPrint(ByVal ws As Worksheet) As Boolean
    Dim pa As String

    On Error Resume Next
    pa = ws.PageSetup.PrintArea
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0

    If Len(pa) > 0 Then
        HasSomethingToPrint = True
        Exit Function
    End If

    If ws.Shapes.Count > 0 Then
        HasSomethingToPrint = True
        Exit Function
    End If

    On Error Resume Next
    HasSomethingToPrint = (Application.WorksheetFunction.CountA(ws.UsedRange) > 0)
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Function


' Writes one PDF. `sheetNames` is a 0-based array of names, or Empty for the
' whole workbook. Returns "" on success, otherwise a note beginning "PROBLEM:".
'
' ExportAsFixedFormat is a method of a Workbook, a Worksheet or a Chart. It is
' NOT a method of the Sheets collection, so ActiveWindow.SelectedSheets.Export...
' does not even compile ("method or data member not found"). Exporting several
' sheets as one document is done the way the Publish dialog does it: select
' them, then export the ACTIVE SHEET, which means "active sheet(s)" and takes
' the whole selected group with continuous page numbers.
'
' When the wanted sheets are every visible sheet there is, the workbook itself
' is exported instead. That is the same PDF without touching the selection at
' all, which is worth having as the path the tidy schedules take.
'
' Selecting marks the workbook as changed, which is why every caller closes it
' with SaveChanges:=False.
Private Function WritePdf(ByVal wb As Workbook, ByVal sheetNames As Variant, _
                          ByVal outPath As String) As String
    Dim source As Object

    If Len(outPath) > MAX_PATH_LEN Then
        WritePdf = "PROBLEM: the PDF path would be " & Len(outPath) & " characters, " & _
                   "which is longer than Windows allows. Export to a folder nearer " & _
                   "the top of the drive. "
        Exit Function
    End If

    On Error Resume Next

    If IsEmpty(sheetNames) Then
        Set source = wb
    Else
        wb.Activate
        wb.Worksheets(sheetNames).Select
        If Err.Number <> 0 Then
            WritePdf = "PROBLEM: could not select the sheets - " & Err.Description & ". "
            Err.Clear
            Exit Function
        End If
        ' Late bound on purpose: ActiveSheet is a Worksheet here and a Workbook
        ' above, and both carry ExportAsFixedFormat with the same arguments.
        Set source = wb.ActiveSheet
    End If

    source.ExportAsFixedFormat _
        Type:=xlTypePDF, fileName:=outPath, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    If Err.Number <> 0 Then
        WritePdf = "PROBLEM: " & PdfErrorText(Err.Number, Err.Description) & " "
        Err.Clear
        Exit Function
    End If

    On Error GoTo 0

    ' Excel can report success on an export that produced no file at all, so
    ' the result is checked rather than assumed. Same reason the header and
    ' footer writes are read back.
    If Not FileExists(outPath) Then
        WritePdf = "PROBLEM: Excel reported success but no file appeared at " & outPath & ". "
    End If
End Function


' True when the wanted sheets are simply everything visible in the workbook,
' charts included, in which case the workbook can be exported as it stands.
Private Function CoversAllVisible(ByVal wb As Workbook, ByVal wanted As Collection) As Boolean
    Dim sh As Object
    Dim shown As Long

    For Each sh In wb.Sheets
        If sh.Visible = xlSheetVisible Then shown = shown + 1
    Next sh

    CoversAllVisible = (shown = wanted.Count)
End Function


' The two errors this actually throws in practice, said in English.
Private Function PdfErrorText(ByVal errNo As Long, ByVal errText As String) As String
    Select Case errNo
        Case 1004, 70
            PdfErrorText = "could not write the PDF (error " & errNo & "). It is " & _
                           "usually open in a PDF reader, or the folder is read only."
        Case Else
            PdfErrorText = "error " & errNo & " - " & errText & "."
    End Select
End Function


' ===========================================================================
' Small helpers
' ===========================================================================

' The first file name that appears twice in the list, or "".
Private Function DuplicateName(ByVal files As Collection) As String
    Dim seen As Object
    Dim i As Long
    Dim key As String

    Set seen = CreateObject("Scripting.Dictionary")
    For i = 1 To files.Count
        key = LCase$(BaseName(files(i)))
        If seen.Exists(key) Then
            DuplicateName = BaseName(files(i))
            Exit Function
        End If
        seen.Add key, True
    Next i
End Function


Private Function FolderOf(ByVal fullPath As String) As String
    Dim p As Long
    p = InStrRev(fullPath, Application.PathSeparator)
    If p > 1 Then FolderOf = Left$(fullPath, p - 1)
End Function


Private Function NameWithoutExtension(ByVal fileName As String) As String
    Dim p As Long
    p = InStrRev(fileName, ".")
    If p > 1 Then
        NameWithoutExtension = Left$(fileName, p - 1)
    Else
        NameWithoutExtension = fileName
    End If
End Function


' Sheet names can hold characters a file name cannot.
Private Function SafeName(ByVal name As String) As String
    Dim bad As String
    Dim i As Long

    bad = "\/:*?""<>|"
    SafeName = name
    For i = 1 To Len(bad)
        SafeName = Replace$(SafeName, Mid$(bad, i, 1), "-")
    Next i
    SafeName = Trim$(SafeName)
    Do While Right$(SafeName, 1) = "."       ' Windows drops a trailing dot
        SafeName = Left$(SafeName, Len(SafeName) - 1)
    Loop
End Function


' A Collection as the 0-based Variant array Worksheets() wants.
Private Function ToArray(ByVal c As Collection) As Variant
    Dim a() As Variant
    Dim i As Long

    ReDim a(0 To c.Count - 1)
    For i = 1 To c.Count
        a(i - 1) = c(i)
    Next i
    ToArray = a
End Function


Private Function JoinCollection(ByVal c As Collection, ByVal sep As String) As String
    Dim i As Long
    For i = 1 To c.Count
        If i > 1 Then JoinCollection = JoinCollection & sep
        JoinCollection = JoinCollection & c(i)
    Next i
End Function


' Opened read only so nothing can be written back, but otherwise exactly the
' way you would open it yourself: links updated and calculation on.
'
' It used to open with UpdateLinks:=0 under manual calculation, to print what
' was saved in the file rather than today's values, the same rule the schedule
' list reads by. A schedule opened cold that way came out of the PDF export
' with a line through every calculated value, while the same file exported by
' hand was clean. Opening it the ordinary way is not worth defending against
' for the sake of a rule that only ever mattered to the QA list.
Private Function OpenQuiet(ByVal fullPath As String) As Workbook
    Dim wb As Workbook

    On Error Resume Next
    Set wb = Workbooks.Open(fileName:=fullPath, ReadOnly:=True, UpdateLinks:=3)
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0

    ' Belt and braces: with automatic calculation the open recalculates, but
    ' the whole bug came from exporting a workbook that had not.
    On Error Resume Next
    Application.Calculate
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0

    Application.ScreenUpdating = False
    Set OpenQuiet = wb
End Function


' ===========================================================================
' Local copies of modMain's plumbing, which is Private to that module.
' Delete these when this is wired into the tool.
' ===========================================================================

' Quieter than the rest of the tool on purpose.
'
' Calculation is forced ON and events are left alone. Turning both off and
' then opening a schedule cold is what struck a line through every calculated
' value in the PDF. Each of those settings is harmless on its own; together,
' before the open, they are not. Only the two that cannot affect what is
' rendered are turned off here.
'
' Forced rather than merely left alone, because Excel may already be sitting
' in manual calculation from something else the user was doing. Whatever it
' was is put back by EndQuiet.
Private Sub BeginQuiet()
    mCalcSaved = Application.Calculation
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    On Error Resume Next
    Application.Calculation = xlCalculationAutomatic
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Sub


Private Sub EndQuiet()
    On Error Resume Next
    If mCalcSaved <> 0 Then Application.Calculation = mCalcSaved
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.AskToUpdateLinks = True
    Application.StatusBar = False
End Sub


Private Sub ExportRecover(ByVal what As String, ByRef wb As Workbook)
    Dim n As Long, d As String

    n = Err.Number
    d = Err.Description

    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    ProgressDone
    EndQuiet
    On Error GoTo 0

    MsgBox what & " stopped." & vbCrLf & vbCrLf & _
           "Error " & n & ": " & d & vbCrLf & vbCrLf & _
           "No schedule was changed. The PDFs written up to that point are on " & _
           "the Log sheet.", vbExclamation, what
End Sub


Private Sub ExportLogStart(ByVal what As String)
    Dim ws As Worksheet

    Set ws = LogSheet()
    If ws Is Nothing Then Exit Sub

    TrimExportLog ws, EXPORT_LOG_RUNS - 1

    ws.Rows("1:3").Insert Shift:=xlDown
    ws.Rows("1:3").ClearFormats

    ws.Range("A1").Value = what & " - " & Format$(Now, "dd/mm/yyyy hh:nn:ss")
    ws.Range("A1:C1").Font.Bold = True
    ws.Range("A1:C1").Interior.Color = RGB(221, 231, 244)
    ws.Range("E1").Value = "RUN"

    ws.Range("A2").Value = "File"
    ws.Range("B2").Value = "Result"
    ws.Range("C2").Value = "Notes"
    ws.Range("A2:C2").Font.Bold = True

    ws.Columns("E").Hidden = True
    mLogRow = 3
End Sub


Private Sub TrimExportLog(ByVal ws As Worksheet, ByVal keep As Long)
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


Private Sub ExportLogLine(ByVal fileName As String, ByVal result As String, ByVal notes As String)
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


Private Function LogSheet() As Worksheet
    Dim ws As Worksheet

    Set ws = GetSheet(ThisWorkbook, SH_LOG)
    If ws Is Nothing Then
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        If Err.Number <> 0 Then
            Err.Clear
            Exit Function
        End If
        ws.Name = SH_LOG
        On Error GoTo 0
    End If

    Set LogSheet = ws
End Function


Private Sub ExportSummary(ByVal what As String, ByVal okCount As Long, ByVal otherCount As Long, _
                          ByVal failCount As Long, ByVal seconds As Double, ByVal extra As String)
    Dim ws As Worksheet
    Dim msg As String

    Set ws = GetSheet(ThisWorkbook, SH_LOG)
    If Not ws Is Nothing Then
        ws.Columns("A:C").AutoFit
        If ws.Columns("C").ColumnWidth > 90 Then ws.Columns("C").ColumnWidth = 90
        ws.Columns("E").Hidden = True
    End If

    msg = okCount & " exported" & vbCrLf & _
          otherCount & " skipped" & vbCrLf & _
          failCount & " failed" & vbCrLf & vbCrLf & _
          "Took " & Duration(seconds) & "."

    If Len(extra) > 0 Then msg = msg & vbCrLf & vbCrLf & extra
    msg = msg & vbCrLf & vbCrLf & "Line by line detail is at the top of the Log sheet."

    MsgBox msg, IIf(failCount > 0, vbExclamation, vbInformation), what
End Sub
