Attribute VB_Name = "modUtil"
Option Explicit

' ---------------------------------------------------------------------------
' modUtil - small shared helpers. No entry points here.
' ---------------------------------------------------------------------------

Public Const SH_FRONT As String = "Front Cover"
Public Const SH_META  As String = "Metadata"
Public Const SH_REV   As String = "Revision Page"
Public Const SH_SETUP As String = "Setup"
Public Const SH_LIST  As String = "ScheduleList"
Public Const SH_LOG   As String = "Log"

' How far down/across we look for the "SCHEDULE OF ..." title cell.
Public Const TITLE_MAX_ROW As Long = 60
Public Const TITLE_MAX_COL As Long = 10

' Cached file system object (built into Windows, nothing to install).
Private mFso As Object

' Progress state.
Private mProgTotal As Long
Private mProgStart As Double
Private mProgCaption As String


' Case-insensitive worksheet lookup. Returns Nothing if absent.
Public Function GetSheet(ByVal wb As Workbook, ByVal wantedName As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If LCase$(Trim$(ws.Name)) = LCase$(Trim$(wantedName)) Then
            Set GetSheet = ws
            Exit Function
        End If
    Next ws
End Function


Public Function IsCommonSheet(ByVal ws As Worksheet) As Boolean
    Select Case LCase$(Trim$(ws.Name))
        Case LCase$(SH_FRONT), LCase$(SH_META), LCase$(SH_REV)
            IsCommonSheet = True
    End Select
End Function


' First exact (trimmed, case-insensitive) match for labelText in column A.
' Top-down, so on the Revision Page "Revision" finds the title-block label at
' A14 and never the revision table header further down the sheet.
Public Function FindLabel(ByVal ws As Worksheet, ByVal labelText As String, _
                          Optional ByVal firstRow As Long = 1, _
                          Optional ByVal lastRow As Long = 0) As Range
    Dim r As Long, endRow As Long
    If lastRow > 0 Then
        endRow = lastRow
    Else
        endRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    End If
    If endRow > ws.Rows.Count Then endRow = ws.Rows.Count
    For r = firstRow To endRow
        If LCase$(Trim$(CStr(ws.Cells(r, 1).Value))) = LCase$(Trim$(labelText)) Then
            Set FindLabel = ws.Cells(r, 1)
            Exit Function
        End If
    Next r
End Function


' The cell holding the schedule title, i.e. text starting "SCHEDULE OF".
' Deliberately bounded so it can never wander into the body of a big schedule.
Public Function FindTitleCell(ByVal ws As Worksheet) As Range
    Dim r As Long, c As Long
    Dim v As Variant
    For r = 1 To TITLE_MAX_ROW
        For c = 1 To TITLE_MAX_COL
            v = ws.Cells(r, c).Value
            If VarType(v) = vbString Then
                If StrComp(Left$(Trim$(v), 11), "SCHEDULE OF", vbTextCompare) = 0 Then
                    Set FindTitleCell = ws.Cells(r, c)
                    Exit Function
                End If
            End If
        Next c
    Next r
End Function


' "$B$3"
Public Function AbsRef(ByVal rng As Range) As String
    AbsRef = rng.Address(True, True, xlA1)
End Function


' "'Revision Page'"
Public Function SheetRef(ByVal sheetName As String) As String
    SheetRef = "'" & Replace$(sheetName, "'", "''") & "'"
End Function


' Writes a formula, preferring Formula2 so dynamic-array functions land correctly.
Public Sub PutFormula(ByVal target As Range, ByVal f As String)
    On Error Resume Next
    target.Formula2 = f
    If Err.Number <> 0 Then
        Err.Clear
        target.Formula = f
    End If
    On Error GoTo 0
End Sub


Public Function BaseName(ByVal fullPath As String) As String
    Dim p As Long
    p = InStrRev(fullPath, Application.PathSeparator)
    If p = 0 Then p = InStrRev(fullPath, "/")
    If p > 0 Then
        BaseName = Mid$(fullPath, p + 1)
    Else
        BaseName = fullPath
    End If
End Function


Public Function EndSep(ByVal folderPath As String) As String
    EndSep = folderPath
    If Len(EndSep) > 0 Then
        If Right$(EndSep, 1) <> Application.PathSeparator Then
            EndSep = EndSep & Application.PathSeparator
        End If
    End If
End Function


' True when a workbook is one of ours: it has the three common sheets.
Public Function LooksLikeSchedule(ByVal wb As Workbook) As Boolean
    LooksLikeSchedule = Not (GetSheet(wb, SH_REV) Is Nothing)
End Function


' The underlying value of a cell, never its displayed text.
'
' .Text returns "#####" when a column is too narrow, which is a display
' artefact and has nothing to do with the value. Reading it was making the
' QA report claim revisions were broken when they were fine.
Public Function CellValue(ByVal rng As Range) As Variant
    CellValue = ""
    On Error Resume Next
    If Not IsError(rng.Value) Then CellValue = rng.Value
    On Error GoTo 0
End Function


' A value as text, for comparisons. Dates are formatted, not left as serials.
Public Function AsText(ByVal v As Variant) As String
    On Error Resume Next
    If IsError(v) Then Exit Function
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    If VarType(v) = vbDate Then
        AsText = Format$(v, "dd/mm/yyyy")
    Else
        AsText = Trim$(CStr(v))
    End If
    On Error GoTo 0
End Function


' ---------------------------------------------------------------------------
' Progress on the status bar, with an estimate of the time left. Cheap, always
' visible, and it does not flicker the way a repainting sheet does.
' ---------------------------------------------------------------------------

Public Sub ProgressStart(ByVal total As Long, ByVal jobName As String)
    mProgTotal = total
    mProgCaption = jobName
    mProgStart = Timer
    Application.StatusBar = jobName & ": starting..."
    DoEvents
End Sub


Public Sub ProgressStep(ByVal doneCount As Long, ByVal itemName As String)
    Dim elapsed As Double
    Dim remaining As Double
    Dim msg As String

    If mProgTotal < 1 Then Exit Sub

    elapsed = Timer - mProgStart
    If elapsed < 0 Then elapsed = elapsed + 86400   ' rolled past midnight

    msg = mProgCaption & ": " & doneCount & " of " & mProgTotal & _
          "  (" & Format$(doneCount / mProgTotal, "0%") & ")"

    If doneCount > 0 And elapsed > 0 Then
        remaining = (elapsed / doneCount) * (mProgTotal - doneCount)
        msg = msg & "  -  about " & Duration(remaining) & " left"
    End If

    If Len(itemName) > 0 Then msg = msg & "  -  " & itemName

    Application.StatusBar = msg
    DoEvents
End Sub


Public Sub ProgressDone()
    mProgTotal = 0
    Application.StatusBar = False
End Sub


Public Function Duration(ByVal seconds As Double) As String
    Dim m As Long, sec As Long
    If seconds < 1 Then
        Duration = "a moment"
        Exit Function
    End If
    m = Int(seconds / 60)
    sec = Int(seconds - m * 60)
    If m > 0 Then
        Duration = m & "m " & sec & "s"
    Else
        Duration = sec & "s"
    End If
End Function


' ---------------------------------------------------------------------------
' Folder handling.
'
' When a workbook is opened from SharePoint / OneDrive / Filery, its .Path is
' a URL like "https://host/personal/.../Documents/Project". Dir() cannot read
' one and raises error 52, so everything below goes through the file system
' object and a URL is resolved to the local synced folder first.
' ---------------------------------------------------------------------------

Public Function Fso() As Object
    If mFso Is Nothing Then Set mFso = CreateObject("Scripting.FileSystemObject")
    Set Fso = mFso
End Function


Public Function IsUrlPath(ByVal p As String) As Boolean
    IsUrlPath = (InStr(1, p, "://", vbTextCompare) > 0)
End Function


Public Function FolderExists(ByVal p As String) As Boolean
    If Len(Trim$(p)) = 0 Then Exit Function
    If IsUrlPath(p) Then Exit Function
    On Error Resume Next
    FolderExists = Fso.FolderExists(p)
    If Err.Number <> 0 Then
        Err.Clear
        FolderExists = False
    End If
    On Error GoTo 0
End Function


' All .xls* files directly in a folder. Uses the file system object rather
' than Dir(), which keeps global state and breaks if two loops interleave.
Public Function FolderWorkbooks(ByVal folderPath As String) As Collection
    Dim c As New Collection
    Dim f As Object
    Dim n As String

    Set FolderWorkbooks = c
    If Not FolderExists(folderPath) Then Exit Function

    On Error Resume Next
    For Each f In Fso.GetFolder(folderPath).files
        n = f.Name
        If Left$(n, 2) <> "~$" Then
            If LCase$(Left$(Fso.GetExtensionName(n), 3)) = "xls" Then c.Add n
        End If
    Next f
    On Error GoTo 0
End Function


' Turns a SharePoint/OneDrive URL into the local synced folder, or returns ""
' if it cannot be worked out. Nothing is guessed: every candidate is tested
' against the file system before it is returned.
Public Function ResolveLocalFolder(ByVal p As String) As String
    Dim roots As Variant
    Dim tail As String
    Dim i As Long, cut As Long
    Dim candidate As String
    Dim markers As Variant
    Dim m As Long

    If Len(Trim$(p)) = 0 Then Exit Function
    If Not IsUrlPath(p) Then
        If FolderExists(p) Then ResolveLocalFolder = p
        Exit Function
    End If

    roots = Array(Environ$("OneDriveCommercial"), Environ$("OneDrive"), _
                  Environ$("OneDriveConsumer"))

    ' The local sync root replaces everything up to and including one of these.
    markers = Array("/Documents/", "/Shared Documents/")

    For m = LBound(markers) To UBound(markers)
        cut = InStr(1, p, CStr(markers(m)), vbTextCompare)
        If cut > 0 Then
            tail = Mid$(p, cut + Len(CStr(markers(m))))
            tail = Replace$(tail, "/", Application.PathSeparator)
            tail = Replace$(tail, "%20", " ")

            For i = LBound(roots) To UBound(roots)
                If Len(CStr(roots(i))) > 0 Then
                    candidate = EndSep(CStr(roots(i))) & tail
                    If FolderExists(candidate) Then
                        ResolveLocalFolder = candidate
                        Exit Function
                    End If
                End If
            Next i
        End If
    Next m
End Function


Public Function FileExists(ByVal fullPath As String) As Boolean
    If Len(Trim$(fullPath)) = 0 Then Exit Function
    If IsUrlPath(fullPath) Then Exit Function
    On Error Resume Next
    FileExists = Fso.FileExists(fullPath)
    If Err.Number <> 0 Then
        Err.Clear
        FileExists = False
    End If
    On Error GoTo 0
End Function


Public Function PickWorkbook(ByVal promptText As String, ByVal startFolder As String) As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = promptText
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel files", "*.xls;*.xlsx;*.xlsm;*.xlsb"
        If Len(startFolder) > 0 Then .InitialFileName = EndSep(startFolder)
        If .Show = -1 Then PickWorkbook = .SelectedItems(1)
    End With
End Function


Public Function PickFolder(ByVal promptText As String) As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = promptText
    If fd.Show = -1 Then PickFolder = fd.SelectedItems(1)
End Function


' Last modified time of a file, or 0 if it cannot be read.
Public Function FileStamp(ByVal fullPath As String) As Double
    On Error Resume Next
    FileStamp = CDbl(Fso.GetFile(fullPath).DateLastModified)
    If Err.Number <> 0 Then
        Err.Clear
        FileStamp = 0
    End If
    On Error GoTo 0
End Function
