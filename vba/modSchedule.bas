Attribute VB_Name = "modSchedule"
Option Explicit

' ---------------------------------------------------------------------------
' modSchedule - everything that reads from, or writes to, one schedule
' workbook. Nothing in here is a button; modMain drives it.
'
' The contract this enforces in every schedule:
'
'   Metadata!B4:B6  are the ONLY cells that link to MAINPROJECTINFO.
'   Everything else on Front Cover / Revision Page reads the local Metadata
'   or the local Revision Page. So a schedule copied away from the project
'   folder keeps working; only Project Name / Number / Client go stale.
' ---------------------------------------------------------------------------

' Rows written on the Metadata sheet. Order matters, labels are the API that
' the ScheduleList reads back, so do not rename them casually.
Private Const META_LAST_ROW As Long = 14


' Repairs / sets up one open schedule workbook. Returns a log string
' (empty means "nothing worth reporting"). Never saves - the caller decides.
Public Function RepairWorkbook(ByVal wbTgt As Workbook, _
                               ByVal mpiFullPath As String, _
                               ByVal setupSheetName As String, _
                               ByVal projNameRef As String, _
                               ByVal projNoRef As String, _
                               ByVal clientRef As String, _
                               ByVal statuses As Variant) As String

    Dim wsFront As Worksheet, wsMeta As Worksheet, wsRev As Worksheet
    Dim log As String
    Dim mpiName As String
    Dim mpiPrefix As String
    Dim revTitle As Range, frontTitle As Range
    Dim recipRef As String, dateRef As String
    Dim r As Long

    mpiName = BaseName(mpiFullPath)
    mpiPrefix = "'[" & mpiName & "]" & setupSheetName & "'!"

    Set wsRev = GetSheet(wbTgt, SH_REV)
    If wsRev Is Nothing Then
        RepairWorkbook = "SKIPPED - no '" & SH_REV & "' sheet."
        Exit Function
    End If

    Set wsFront = GetSheet(wbTgt, SH_FRONT)
    Set wsMeta = GetSheet(wbTgt, SH_META)

    ' Point any pre-existing link to the MPI at THIS MPI before we touch
    ' formulas, otherwise Excel may bind new formulas to the stale link.
    log = log & RepointMpiLink(wbTgt, mpiFullPath)

    If wsMeta Is Nothing Then
        Set wsMeta = wbTgt.Worksheets.Add(After:=wbTgt.Worksheets(wbTgt.Worksheets.Count))
        wsMeta.Name = SH_META
        log = log & "Created Metadata sheet. "
    End If

    ' --- Metadata --------------------------------------------------------
    log = log & WriteMetadata(wsMeta, wsRev, mpiPrefix, projNameRef, projNoRef, clientRef)

    ' --- Revision Page ---------------------------------------------------
    Set revTitle = FindTitleCell(wsRev)
    If revTitle Is Nothing Then
        log = log & "No 'SCHEDULE OF...' title found on Revision Page. "
    ElseIf revTitle.Row > 1 Then
        ' The cell directly above the title is the project name.
        PutFormula wsRev.Cells(revTitle.Row - 1, revTitle.Column), _
                   "=" & SheetRef(wsMeta.Name) & "!B4"
    End If

    WriteIfLabelled wsRev, "Project Name", "=" & SheetRef(wsMeta.Name) & "!B4", log
    WriteIfLabelled wsRev, "Project no.", "=" & SheetRef(wsMeta.Name) & "!B5", log
    WriteIfLabelled wsRev, "Recipient", "=" & SheetRef(wsMeta.Name) & "!B6", log
    WriteIfLabelled wsRev, "Document no", "=" & SheetRef(wsMeta.Name) & "!B2", log

    ' --- Front Cover -----------------------------------------------------
    If wsFront Is Nothing Then
        log = log & "No 'Front Cover' sheet. "
    Else
        recipRef = AbsRefOfLabelValue(wsRev, "Recipient")
        dateRef = AbsRefOfLabelValue(wsRev, "Date")

        If Len(recipRef) = 0 Then
            log = log & "Front Cover 'Intended for' left alone (no Recipient on Revision Page). "
        Else
            WriteBelowLabel wsFront, "Intended for", _
                            "=" & SheetRef(wsRev.Name) & "!" & recipRef, log
        End If

        If Len(dateRef) = 0 Then
            log = log & "Front Cover 'Date' left alone (no Date on Revision Page). "
        Else
            WriteBelowLabel wsFront, "Date", _
                            "=" & SheetRef(wsRev.Name) & "!" & dateRef, log
        End If

        Set frontTitle = FindTitleCell(wsFront)
        If frontTitle Is Nothing Then
            log = log & "No 'SCHEDULE OF...' title found on Front Cover. "
        Else
            PutFormula frontTitle, "=" & SheetRef(wsMeta.Name) & "!B3"
            If frontTitle.Row > 1 Then
                PutFormula wsFront.Cells(frontTitle.Row - 1, frontTitle.Column), _
                           "=" & SheetRef(wsMeta.Name) & "!B4"
            End If
        End If
    End If

    ' --- Schedule sheet title -------------------------------------------
    log = log & LinkScheduleSheetTitle(wbTgt, wsMeta, wsRev)

    ' --- Suitability dropdown, kept local so it survives without the MPI --
    log = log & WriteStatusList(wsMeta, wsRev, statuses)

    ' --- Housekeeping ----------------------------------------------------
    log = log & RemoveDeadNames(wbTgt)
    log = log & TidyLinks(wbTgt, mpiName)

    RepairWorkbook = Trim$(log)
End Function


' Writes the Metadata sheet content. Rows 4-6 are the only MPI links.
Private Function WriteMetadata(ByVal wsMeta As Worksheet, ByVal wsRev As Worksheet, _
                               ByVal mpiPrefix As String, _
                               ByVal projNameRef As String, ByVal projNoRef As String, _
                               ByVal clientRef As String) As String
    Dim lo As ListObject
    Dim log As String
    Dim rv As String

    rv = SheetRef(wsRev.Name) & "!"

    ' Grow the table first so the writes land inside it.
    On Error Resume Next
    Set lo = wsMeta.ListObjects(1)
    On Error GoTo 0
    If Not lo Is Nothing Then
        On Error Resume Next
        lo.Resize wsMeta.Range("A1:B" & META_LAST_ROW)
        On Error GoTo 0
    End If

    wsMeta.Range("A1").Value = "Header"
    wsMeta.Range("B1").Value = "Value"

    wsMeta.Range("A2").Value = "DocumentNumber"
    PutFormula wsMeta.Range("B2"), _
        "=TRIM(TEXTBEFORE(TEXTBEFORE(TEXTAFTER(CELL(""filename"",A1),""[""),""]""),""-"",-1))"

    wsMeta.Range("A3").Value = "ScheduleName"
    PutFormula wsMeta.Range("B3"), "=TRIM(" & rv & "$A$4)"

    wsMeta.Range("A4").Value = "Project Name"
    PutFormula wsMeta.Range("B4"), "=" & mpiPrefix & projNameRef

    wsMeta.Range("A5").Value = "Project Number"
    PutFormula wsMeta.Range("B5"), "=" & mpiPrefix & projNoRef

    wsMeta.Range("A6").Value = "Client"
    PutFormula wsMeta.Range("B6"), "=" & mpiPrefix & clientRef

    ' Rows 7-14 mirror the Revision Page title block so the MPI can read one
    ' sheet per schedule instead of hunting for labels in a formatted page.
    log = log & MetaRow(wsMeta, 7, "DocumentType", wsRev, "Document type")
    log = log & MetaRow(wsMeta, 8, "Revision", wsRev, "Revision")
    log = log & MetaRow(wsMeta, 9, "Date", wsRev, "Date")
    log = log & MetaRow(wsMeta, 10, "Prepared by", wsRev, "Prepared by")
    log = log & MetaRow(wsMeta, 11, "Checked by", wsRev, "Checked by")
    log = log & MetaRow(wsMeta, 12, "Approved by", wsRev, "Approved by")
    log = log & MetaRow(wsMeta, 13, "Suitability Status", wsRev, "Suitability Status")
    log = log & MetaRow(wsMeta, 14, "Suitability Description", wsRev, "Suitability Description")

    wsMeta.Range("B9").NumberFormat = "dd/mm/yyyy"
    wsMeta.Columns("A:B").AutoFit

    ' The title cell on the Revision Page is the one place a schedule name is
    ' typed. Make sure it is not itself a formula pointing somewhere else.
    If wsRev.Range("A4").HasFormula Then
        log = log & "Revision Page A4 (schedule name) is a formula - it should be typed text. "
    End If

    WriteMetadata = log
End Function


Private Function MetaRow(ByVal wsMeta As Worksheet, ByVal metaRowNo As Long, _
                         ByVal header As String, ByVal wsRev As Worksheet, _
                         ByVal revLabel As String) As String
    Dim lbl As Range
    wsMeta.Cells(metaRowNo, 1).Value = header
    Set lbl = FindLabel(wsRev, revLabel)
    If lbl Is Nothing Then
        wsMeta.Cells(metaRowNo, 2).ClearContents
        MetaRow = "Revision Page has no '" & revLabel & "' label. "
    Else
        PutFormula wsMeta.Cells(metaRowNo, 2), _
            "=" & SheetRef(wsRev.Name) & "!" & AbsRef(lbl.Offset(0, 1))
    End If
End Function


' "$B$15" for the value cell next to a Revision Page label, or "" if the
' label is not there. Never guesses - a wrong guess is worse than no link.
Private Function AbsRefOfLabelValue(ByVal ws As Worksheet, ByVal labelText As String) As String
    Dim lbl As Range
    Set lbl = FindLabel(ws, labelText)
    If lbl Is Nothing Then Exit Function
    AbsRefOfLabelValue = AbsRef(lbl.Offset(0, 1))
End Function


Private Sub WriteIfLabelled(ByVal ws As Worksheet, ByVal labelText As String, _
                            ByVal f As String, ByRef log As String)
    Dim lbl As Range
    Set lbl = FindLabel(ws, labelText)
    If lbl Is Nothing Then
        log = log & "No '" & labelText & "' label on " & ws.Name & ". "
    Else
        PutFormula lbl.Offset(0, 1), f
    End If
End Sub


Private Sub WriteBelowLabel(ByVal ws As Worksheet, ByVal labelText As String, _
                            ByVal f As String, ByRef log As String)
    Dim lbl As Range
    Set lbl = FindLabel(ws, labelText)
    If lbl Is Nothing Then
        log = log & "No '" & labelText & "' label on " & ws.Name & ". "
    Else
        PutFormula lbl.Offset(1, 0), f
    End If
End Sub


' The schedule sheet's own title is linked to Metadata only when it already
' matches the Revision Page title. If they disagree we report it rather than
' silently overwrite whichever one the engineer meant.
Private Function LinkScheduleSheetTitle(ByVal wb As Workbook, ByVal wsMeta As Worksheet, _
                                        ByVal wsRev As Worksheet) As String
    Dim ws As Worksheet
    Dim t As Range
    Dim master As String

    master = Trim$(CStr(wsRev.Range("A4").Value))
    If Len(master) = 0 Then Exit Function

    For Each ws In wb.Worksheets
        If Not IsCommonSheet(ws) Then
            Set t = FindTitleCell(ws)
            If Not t Is Nothing Then
                If t.HasFormula Then
                    ' already linked, leave it
                ElseIf StrComp(Trim$(CStr(t.Value)), master, vbTextCompare) = 0 Then
                    PutFormula t, "=" & SheetRef(wsMeta.Name) & "!B3"
                Else
                    LinkScheduleSheetTitle = LinkScheduleSheetTitle & _
                        "Title on '" & ws.Name & "' (" & Trim$(CStr(t.Value)) & _
                        ") differs from Revision Page A4 (" & master & ") - left alone. "
                End If
            End If
        End If
    Next ws
End Function


' Puts the suitability code list on the Metadata sheet and points the
' revision table's Status column at it.
Private Function WriteStatusList(ByVal wsMeta As Worksheet, ByVal wsRev As Worksheet, _
                                 ByVal statuses As Variant) As String
    Dim i As Long, n As Long
    Dim lo As ListObject
    Dim col As ListColumn
    Dim rng As Range

    If Not IsArray(statuses) Then Exit Function
    n = UBound(statuses) - LBound(statuses) + 1
    If n < 1 Then Exit Function

    wsMeta.Range("D1:D200").ClearContents
    wsMeta.Range("D1").Value = "Suitability Codes"
    For i = 0 To n - 1
        wsMeta.Cells(2 + i, 4).Value = statuses(LBound(statuses) + i)
    Next i
    wsMeta.Columns("D").AutoFit

    On Error Resume Next
    Set lo = wsRev.ListObjects("RevisionTable")
    On Error GoTo 0
    If lo Is Nothing Then
        WriteStatusList = "No 'RevisionTable' on the Revision Page. "
        Exit Function
    End If

    On Error Resume Next
    Set col = lo.ListColumns("Status")
    On Error GoTo 0
    If col Is Nothing Then
        WriteStatusList = "RevisionTable has no 'Status' column. "
        Exit Function
    End If

    Set rng = col.DataBodyRange
    If rng Is Nothing Then Exit Function

    On Error Resume Next
    rng.Validation.Delete
    rng.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, _
        Formula1:="=" & SheetRef(wsMeta.Name) & "!$D$2:$D$" & (n + 1)
    rng.Validation.IgnoreBlank = True
    rng.Validation.InCellDropdown = True
    On Error GoTo 0
End Function


' Removes defined names left over from copied templates (#REF! or pointing at
' some other workbook). These are what make Excel nag about updating links.
Private Function RemoveDeadNames(ByVal wb As Workbook) As String
    Dim nm As Name
    Dim i As Long
    Dim killed As Long
    Dim rt As String
    Dim re As Object

    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.pattern = "\[[^\]\[]+\.xls[a-z]*\]"

    For i = wb.Names.Count To 1 Step -1
        Set nm = wb.Names(i)
        If Left$(nm.Name, 6) <> "_xlnm." And InStr(nm.Name, "!_xlnm.") = 0 Then
            On Error Resume Next
            rt = nm.RefersTo
            On Error GoTo 0
            If InStr(1, rt, "#REF!", vbTextCompare) > 0 _
               Or re.Test(rt) Then
                On Error Resume Next
                nm.Delete
                If Err.Number = 0 Then killed = killed + 1
                Err.Clear
                On Error GoTo 0
            End If
        End If
    Next i

    If killed > 0 Then RemoveDeadNames = "Removed " & killed & " broken/external defined name(s). "
End Function


' Repoints an existing link to a workbook with the MPI's file name at the
' current MPI. Returns a log fragment.
Private Function RepointMpiLink(ByVal wb As Workbook, ByVal mpiFullPath As String) As String
    Dim links As Variant
    Dim i As Long
    Dim mpiName As String

    mpiName = BaseName(mpiFullPath)

    On Error Resume Next
    links = wb.LinkSources(xlExcelLinks)
    On Error GoTo 0
    If IsEmpty(links) Then Exit Function

    For i = LBound(links) To UBound(links)
        If StrComp(BaseName(CStr(links(i))), mpiName, vbTextCompare) = 0 Then
            If StrComp(CStr(links(i)), mpiFullPath, vbTextCompare) <> 0 Then
                On Error Resume Next
                wb.ChangeLink Name:=CStr(links(i)), NewName:=mpiFullPath, Type:=xlExcelLinks
                On Error GoTo 0
            End If
        End If
    Next i
End Function


' Breaks external links that no formula uses any more (the usual leftovers of
' copying a schedule). Links that ARE still referenced are reported, never
' broken - breaking one would silently hardcode a live formula.
Private Function TidyLinks(ByVal wb As Workbook, ByVal mpiName As String) As String
    Dim links As Variant
    Dim used As Object
    Dim i As Long
    Dim src As String, nameOnly As String
    Dim broke As Long
    Dim log As String

    On Error Resume Next
    links = wb.LinkSources(xlExcelLinks)
    On Error GoTo 0
    If IsEmpty(links) Then Exit Function

    Set used = ReferencedWorkbooks(wb)

    For i = LBound(links) To UBound(links)
        src = CStr(links(i))
        nameOnly = LCase$(BaseName(src))
        If nameOnly <> LCase$(mpiName) Then
            If Not used.Exists(nameOnly) Then
                On Error Resume Next
                wb.BreakLink Name:=src, Type:=xlExcelLinks
                Err.Clear
                On Error GoTo 0
                broke = broke + 1
            Else
                log = log & "STILL LINKED to " & BaseName(src) & " - check this. "
            End If
        End If
    Next i

    If broke > 0 Then log = log & "Cleared " & broke & " unused external link(s). "
    TidyLinks = log
End Function


' Set of lower-case workbook file names referenced by any formula in wb.
Private Function ReferencedWorkbooks(ByVal wb As Workbook) As Object
    Dim d As Object
    Dim ws As Worksheet
    Dim rng As Range, cell As Range
    Dim re As Object, matches As Object, m As Object
    Dim f As String

    Set d = CreateObject("Scripting.Dictionary")
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.pattern = "\[([^\]\[]+\.xls[a-z]*)\]"
    re.IgnoreCase = True

    For Each ws In wb.Worksheets
        Set rng = Nothing
        On Error Resume Next
        Set rng = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
        On Error GoTo 0
        If Not rng Is Nothing Then
            For Each cell In rng.Cells
                f = cell.Formula
                If InStr(f, "[") > 0 Then
                    Set matches = re.Execute(f)
                    For Each m In matches
                        d(LCase$(m.SubMatches(0))) = True
                    Next m
                End If
            Next cell
        End If
    Next ws

    Set ReferencedWorkbooks = d
End Function


' ---------------------------------------------------------------------------
' Reading side - used by the ScheduleList refresh.
' ---------------------------------------------------------------------------

' Reads a value from a schedule's Metadata sheet by its header text.
Public Function ReadMeta(ByVal wb As Workbook, ByVal header As String) As String
    Dim wsMeta As Worksheet
    Dim lbl As Range

    Set wsMeta = GetSheet(wb, SH_META)
    If wsMeta Is Nothing Then Exit Function

    Set lbl = FindLabel(wsMeta, header, 1, META_LAST_ROW + 5)
    If lbl Is Nothing Then Exit Function

    ReadMeta = CellText(lbl.Offset(0, 1))
End Function


' Distinct Status values already present in a schedule's revision table.
Public Sub CollectStatuses(ByVal wb As Workbook, ByVal bag As Object)
    Dim wsRev As Worksheet
    Dim lo As ListObject
    Dim col As ListColumn
    Dim cell As Range
    Dim v As String

    Set wsRev = GetSheet(wb, SH_REV)
    If wsRev Is Nothing Then Exit Sub

    On Error Resume Next
    Set lo = wsRev.ListObjects("RevisionTable")
    If lo Is Nothing Then Exit Sub
    Set col = lo.ListColumns("Status")
    On Error GoTo 0
    If col Is Nothing Then Exit Sub
    If col.DataBodyRange Is Nothing Then Exit Sub

    For Each cell In col.DataBodyRange.Cells
        v = Trim$(CStr(cell.Value))
        If Len(v) > 0 Then bag(v) = True
    Next cell
End Sub


' Appends one row to a schedule's revision table.
Public Function AppendRevision(ByVal wb As Workbook, ByVal rev As String, ByVal status As String, _
                               ByVal issueDate As Variant, ByVal prep As String, _
                               ByVal chk As String, ByVal app As String, _
                               ByVal descr As String) As String
    Dim wsRev As Worksheet
    Dim lo As ListObject
    Dim newRow As ListRow
    Dim cell As Range

    Set wsRev = GetSheet(wb, SH_REV)
    If wsRev Is Nothing Then
        AppendRevision = "no Revision Page"
        Exit Function
    End If

    On Error Resume Next
    Set lo = wsRev.ListObjects("RevisionTable")
    On Error GoTo 0
    If lo Is Nothing Then
        AppendRevision = "no RevisionTable"
        Exit Function
    End If

    ' Refuse to add a revision that is already there.
    If Not lo.ListColumns("Revision").DataBodyRange Is Nothing Then
        For Each cell In lo.ListColumns("Revision").DataBodyRange.Cells
            If StrComp(Trim$(CStr(cell.Value)), rev, vbTextCompare) = 0 Then
                AppendRevision = "revision " & rev & " already present"
                Exit Function
            End If
        Next cell
    End If

    Set newRow = lo.ListRows.Add
    SetCol lo, newRow, "Revision", rev
    SetCol lo, newRow, "Status", status
    SetCol lo, newRow, "Date", issueDate
    SetCol lo, newRow, "Prepared by", prep
    SetCol lo, newRow, "Checked by", chk
    SetCol lo, newRow, "Approved by", app
    SetCol lo, newRow, "Description", descr
End Function


Private Sub SetCol(ByVal lo As ListObject, ByVal lr As ListRow, _
                   ByVal colName As String, ByVal v As Variant)
    Dim idx As Long
    On Error Resume Next
    idx = lo.ListColumns(colName).Index
    On Error GoTo 0
    If idx > 0 Then lr.Range.Cells(1, idx).Value = v
End Sub
