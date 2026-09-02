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

' How far down/across we look for the "SCHEDULE OF ..." title cell.
Public Const TITLE_MAX_ROW As Long = 60
Public Const TITLE_MAX_COL As Long = 10


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


' Safe string of a cell, turning errors into "" so reports never blow up.
Public Function CellText(ByVal rng As Range) As String
    On Error Resume Next
    If IsError(rng.Value) Then
        CellText = "#ERROR"
    Else
        CellText = Trim$(CStr(rng.Text))
    End If
    On Error GoTo 0
End Function
