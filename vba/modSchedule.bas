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

' Revision Page cells that are typed per document and must not be inherited
' when the common sheets are copied from a reference schedule.
Private Const KEEP_LABELS As String = _
    "Document type,Delref Classification,BSUID,Trigger Events"


' A schedule's revision history, lifted out before its sheet is replaced.
Private Type RevData
    Valid As Boolean
    Headers() As String
    Cells() As Variant
    RowCount As Long
    ColCount As Long
End Type


' One sheet's header and footer, captured from the source workbook.
Public Type HFSet
    Valid As Boolean
    LeftHeader As String
    CenterHeader As String
    RightHeader As String
    LeftFooter As String
    CenterFooter As String
    RightFooter As String
    DiffFirstPage As Boolean
    OddAndEven As Boolean
    AlignMargins As Boolean
    ScaleWithDoc As Boolean
    FirstLeftHeader As String
    FirstCenterHeader As String
    FirstRightHeader As String
    FirstLeftFooter As String
    FirstCenterFooter As String
    FirstRightFooter As String
    EvenLeftHeader As String
    EvenCenterHeader As String
    EvenRightHeader As String
    EvenLeftFooter As String
    EvenCenterFooter As String
    EvenRightFooter As String
End Type

' Revision families in priority order, lowest first. "P" preliminary,
' "C" construction, "AF" as fitted, so AF01 beats C09 beats P12.
' Add a family here and re-run setup to push it into every schedule.
Public Const REV_PREFIXES As String = "P,C,AF"


' The suitability codes seeded onto the Setup sheet and pushed into every
' schedule's Metadata sheet. ISO 19650 CDE status codes.
'
' Format is "<code> - <description>". The title block splits on the " - ",
' so no description may contain one.
'
' S5 is deliberately absent: the standard leaves it to project guidance, so a
' shared dropdown cannot say anything useful about it. Historic revision rows
' that already say "S5 - ..." keep their text, as they should - the revision
' table is a record of what was issued, not something to rewrite. Add a row to
' column F of the Setup sheet if this project needs S5 back.
Public Function DefaultSuitabilityCodes() As Variant
    DefaultSuitabilityCodes = Array( _
        "S0 - Work in Progress", _
        "S1 - Suitable for Coordination", _
        "S2 - Suitable for Information", _
        "S3 - Suitable for Review and Comment", _
        "S4 - Suitable for Stage Approval", _
        "S6 - Suitable for PIM Authorisation", _
        "S7 - Suitable for AIM Authorisation", _
        "A1 - Authorised and Accepted", _
        "A2 - Authorised and Accepted", _
        "B1 - Published with Comments", _
        "B2 - Published with Comments")
End Function


' Repairs / sets up one open schedule workbook. Returns a log string
' (empty means "nothing worth reporting"). Never saves - the caller decides.
Public Function RepairWorkbook(ByVal wbTgt As Workbook, _
                               ByVal mpiFullPath As String, _
                               ByVal setupSheetName As String, _
                               ByVal projNameRef As String, _
                               ByVal projNoRef As String, _
                               ByVal clientRef As String, _
                               ByVal statuses As Variant, _
                               ByVal fields As Variant) As String

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
    log = log & WriteMetadata(wsMeta, wsRev, mpiPrefix, projNameRef, projNoRef, clientRef, fields)

    ' Wire any Revision Page label that matches a project field.
    log = log & LinkProjectFields(wsRev, wsMeta, fields)

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

    ' --- Revision Page title block, driven by the revision table ---------
    log = log & WriteRevisionFormulas(wsRev)

    ' --- Any other cell anywhere that ranks the revision table -----------
    log = log & UpgradeRevisionFormulas(wbTgt)

    ' --- Suitability dropdown, kept local so it survives without the MPI --
    log = log & WriteStatusList(wsMeta, wsRev, statuses)

    ' --- Anything hand-written that got externalised by a sheet copy ------
    log = log & LocaliseFormulas(wbTgt)

    ' --- Housekeeping ----------------------------------------------------
    log = log & RemoveDeadNames(wbTgt)
    log = log & TidyLinks(wbTgt, mpiName)

    RepairWorkbook = Trim$(log)
End Function


' Writes the Metadata sheet content. Rows 4-6 are the only MPI links.
Private Function WriteMetadata(ByVal wsMeta As Worksheet, ByVal wsRev As Worksheet, _
                               ByVal mpiPrefix As String, _
                               ByVal projNameRef As String, ByVal projNoRef As String, _
                               ByVal clientRef As String, ByVal fields As Variant) As String
    Dim lo As ListObject
    Dim log As String
    Dim rv As String
    Dim lastRow As Long
    Dim i As Long

    rv = SheetRef(wsRev.Name) & "!"

    ' Grow the table first so the writes land inside it.
    On Error Resume Next
    Set lo = wsMeta.ListObjects(1)
    On Error GoTo 0
    lastRow = META_LAST_ROW + FieldCount(fields)
    If Not lo Is Nothing Then
        On Error Resume Next
        lo.Resize wsMeta.Range("A1:B" & lastRow)
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

    ' Extra project fields from the MPI, one row each after the fixed ones.
    ' Rewritten in full every run, so removing a field on the MPI removes it
    ' from every schedule too.
    wsMeta.Range("A" & (META_LAST_ROW + 1) & ":B" & (META_LAST_ROW + 60)).ClearContents
    For i = 1 To FieldCount(fields)
        wsMeta.Cells(META_LAST_ROW + i, 1).Value = fields(i, 1)
        PutFormula wsMeta.Cells(META_LAST_ROW + i, 2), "=" & mpiPrefix & fields(i, 2)
    Next i

    wsMeta.Range("B9").NumberFormat = "dd/mm/yyyy"
    wsMeta.Columns("A:B").AutoFit

    ' The title cell on the Revision Page is the one place a schedule name is
    ' typed. Make sure it is not itself a formula pointing somewhere else.
    If wsRev.Range("A4").HasFormula Then
        log = log & "Revision Page A4 (schedule name) is a formula - it should be typed text. "
    End If

    WriteMetadata = log
End Function


Public Function FieldCount(ByVal fields As Variant) As Long
    On Error Resume Next
    If IsArray(fields) Then FieldCount = UBound(fields, 1)
    If Err.Number <> 0 Then
        Err.Clear
        FieldCount = 0
    End If
    On Error GoTo 0
End Function


' Points any Revision Page label that matches a project field at the Metadata
' row holding it. Add "DfE Code" to column A of the reference schedule''s
' Revision Page, and every schedule picks the value up from the MPI.
Private Function LinkProjectFields(ByVal wsRev As Worksheet, ByVal wsMeta As Worksheet, _
                                   ByVal fields As Variant) As String
    Dim i As Long, n As Long
    Dim lbl As Range
    Dim wired As Long

    n = FieldCount(fields)
    If n = 0 Then Exit Function

    For i = 1 To n
        Set lbl = FindLabel(wsRev, CStr(fields(i, 1)))
        If Not lbl Is Nothing Then
            PutFormula lbl.Offset(0, 1), _
                "=" & SheetRef(wsMeta.Name) & "!B" & (META_LAST_ROW + i)
            wired = wired + 1
        End If
    Next i

    If wired > 0 Then _
        LinkProjectFields = "Linked " & wired & " project field(s) on the Revision Page. "
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


' Rewrites references that point at another workbook's copy of a sheet this
' workbook has itself.
'
' Copying a sheet in from a reference turns every formula on it that referred
' to a sheet in the reference into an external one:
'
'     ='Revision Page'!B26   becomes   ='[Golden.xlsx]Revision Page'!B26
'
' The repair fixes the cells it writes itself, but not one somebody added by
' hand on the cover. This catches all of them: any external reference naming a
' sheet that exists here is pointed at the local sheet instead. References to
' sheets this workbook does not have are left alone, because those are real
' links to somewhere else and breaking them silently would be worse.
Private Function LocaliseFormulas(ByVal wb As Workbook) As String
    Dim ws As Worksheet, target As Worksheet
    Dim rng As Range, cell As Range
    Dim re As Object, matches As Object, m As Object
    Dim f As String, newF As String
    Dim sheetName As String
    Dim changed As Long

    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    ' '[Book.xlsx]Sheet Name'!  or  [Book.xlsx]SheetName!
    re.pattern = "'\[[^\]\[]+\]([^']+)'!|\[[^\]\[]+\]([A-Za-z0-9_.]+)!"

    For Each ws In wb.Worksheets
        Set rng = Nothing
        On Error Resume Next
        Set rng = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
        On Error GoTo 0
        If Not rng Is Nothing Then
            For Each cell In rng.Cells
                f = CellFormula(cell)
                If InStr(f, "[") > 0 Then
                    newF = f
                    Set matches = re.Execute(f)
                    For Each m In matches
                        sheetName = m.SubMatches(0)
                        If Len(sheetName) = 0 Then sheetName = m.SubMatches(1)
                        Set target = GetSheet(wb, sheetName)
                        If Not target Is Nothing Then
                            newF = Replace(newF, m.Value, SheetRef(target.Name) & "!")
                        End If
                    Next m

                    ' Table references lose their workbook prefix the same way.
                    newF = RegexReplace(newF, "\[[^\]\[]+\]!(?=[A-Za-z_])", "")

                    If StrComp(newF, f, vbBinaryCompare) <> 0 Then
                        PutFormula cell, newF
                        changed = changed + 1
                    End If
                End If
            Next cell
        End If
    Next ws

    If changed > 0 Then _
        LocaliseFormulas = "Pointed " & changed & " formula(s) back at this workbook's own sheets. "
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


' Rewrites the seven title-block cells that summarise the revision table, so
' that they rank revisions by family (AF > C > P) and then by number.
Private Function WriteRevisionFormulas(ByVal wsRev As Worksheet) As String
    Dim lo As ListObject
    Dim log As String

    On Error Resume Next
    Set lo = wsRev.ListObjects("RevisionTable")
    On Error GoTo 0
    If lo Is Nothing Then
        WriteRevisionFormulas = "No 'RevisionTable' - revision formulas left alone. "
        Exit Function
    End If

    WriteIfLabelled wsRev, "Revision", RevFormula("XLOOKUP(MAX(rank),rank,rev)", False), log
    WriteIfLabelled wsRev, "Date", RevFormula("XLOOKUP(MAX(rank),rank,RevisionTable[Date])", False), log
    WriteIfLabelled wsRev, "Prepared by", RevFormula("XLOOKUP(MAX(rank),rank,RevisionTable[Prepared by])", False), log
    WriteIfLabelled wsRev, "Checked by", RevFormula("XLOOKUP(MAX(rank),rank,RevisionTable[Checked by])", False), log
    WriteIfLabelled wsRev, "Approved by", RevFormula("XLOOKUP(MAX(rank),rank,RevisionTable[Approved by])", False), log
    WriteIfLabelled wsRev, "Suitability Status", _
        RevFormula("IFERROR(LEFT(stat,FIND("" - "",stat)-1),stat)", True), log
    WriteIfLabelled wsRev, "Suitability Description", _
        RevFormula("IFERROR(TEXTAFTER(stat,"" - ""),"""")", True), log

    WriteRevisionFormulas = log
End Function


' Rewrites every formula ANYWHERE in the workbook that picks the latest row
' out of the revision table, not just the seven on the Revision Page. Catches
' the copies that live on the front cover or in a schedule sheet header.
'
' Only touches cells whose formula already ranks RevisionTable[Revision] with
' MAX(), so it cannot wander into unrelated formulas.
Private Function UpgradeRevisionFormulas(ByVal wb As Workbook) As String
    Dim ws As Worksheet
    Dim rng As Range, cell As Range
    Dim f As String, newF As String
    Dim changed As Long

    For Each ws In wb.Worksheets
        Set rng = Nothing
        On Error Resume Next
        Set rng = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
        On Error GoTo 0
        If Not rng Is Nothing Then
            For Each cell In rng.Cells
                f = CellFormula(cell)
                If InStr(1, f, "RevisionTable[Revision]", vbTextCompare) > 0 _
                   And InStr(1, f, "MAX(", vbTextCompare) > 0 Then
                    newF = RebuildRevFormula(f)
                    If Len(newF) > 0 And StrComp(newF, f, vbBinaryCompare) <> 0 Then
                        PutFormula cell, newF
                        changed = changed + 1
                    End If
                End If
            Next cell
        End If
    Next ws

    If changed > 0 Then _
        UpgradeRevisionFormulas = "Rebuilt " & changed & " revision formula(s) with AF > C > P ranking. "
End Function


Private Function CellFormula(ByVal cell As Range) As String
    On Error Resume Next
    CellFormula = cell.Formula2
    If Len(CellFormula) = 0 Then CellFormula = cell.Formula
    On Error GoTo 0
End Function


' Works out what an existing revision formula returns, then regenerates it with
' the family-aware ranking. Returns "" if it cannot tell, in which case the
' original formula is left exactly as it is.
Private Function RebuildRevFormula(ByVal f As String) As String
    Dim re As Object, matches As Object
    Dim i As Long
    Dim col As String

    ' The two suitability formulas split "S5 - Suitable for ..." apart.
    If InStr(1, f, " - ", vbTextCompare) > 0 Then
        If InStr(1, f, "FIND(", vbTextCompare) > 0 And InStr(1, f, "LEFT(", vbTextCompare) > 0 Then
            RebuildRevFormula = RevFormula("IFERROR(LEFT(stat,FIND("" - "",stat)-1),stat)", True)
            Exit Function
        End If
        If InStr(1, f, "TEXTAFTER(", vbTextCompare) > 0 Then
            RebuildRevFormula = RevFormula("IFERROR(TEXTAFTER(stat,"" - ""),"""")", True)
            Exit Function
        End If
    End If

    ' Otherwise the answer column is the last RevisionTable[...] that is not
    ' the Revision column itself.
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.pattern = "RevisionTable\[([^\]\[#]+)\]"
    Set matches = re.Execute(f)

    For i = 0 To matches.Count - 1
        If StrComp(matches(i).SubMatches(0), "Revision", vbTextCompare) <> 0 Then
            col = matches(i).SubMatches(0)
        End If
    Next i

    If Len(col) = 0 Then
        RebuildRevFormula = RevFormula("XLOOKUP(MAX(rank),rank,rev)", False)
    Else
        RebuildRevFormula = RevFormula("XLOOKUP(MAX(rank),rank,RevisionTable[" & col & "])", False)
    End If
End Function


' Builds the shared LET wrapper that ranks the revision table, then returns
' whatever the caller asked for from the winning row.
'
'   num   the digits, e.g. AF01 -> 1
'   pri   the family, P=1 C=2 AF=3, anything unrecognised = 0
'   rank  pri*1000 + num, so AF01 (3001) beats C09 (2009) beats P12 (1012)
Private Function RevFormula(ByVal resultExpr As String, ByVal needStatus As Boolean) As String
    Dim names As Variant
    Dim order As Variant
    Dim subs As String, prio As String, closers As String
    Dim i As Long, idx As Long

    names = Split(REV_PREFIXES, ",")
    order = ByLengthDesc(names)

    subs = "t"
    For i = LBound(order) To UBound(order)
        idx = order(i)
        subs = "SUBSTITUTE(" & subs & ",""" & UCase$(Trim$(names(idx))) & ""","""")"
    Next i

    For i = LBound(order) To UBound(order)
        idx = order(i)
        prio = prio & "IF(LEFT(t," & Len(Trim$(names(idx))) & ")=""" & _
               UCase$(Trim$(names(idx))) & """," & (idx + 1) & ","
        closers = closers & ")"
    Next i
    prio = prio & "0" & closers

    RevFormula = "=LET(rev,RevisionTable[Revision]," & _
                 "t,UPPER(TRIM(rev))," & _
                 "num,IFERROR(--" & subs & ",0)," & _
                 "pri," & prio & "," & _
                 "rank,pri*1000+num,"
    If needStatus Then
        RevFormula = RevFormula & "stat,INDEX(RevisionTable[Status],XMATCH(MAX(rank),rank)),"
    End If
    RevFormula = RevFormula & resultExpr & ")"
End Function


' Indexes into arr, longest string first, so "AF" is stripped and matched
' before "A" would be if someone ever adds one.
Private Function ByLengthDesc(ByVal arr As Variant) As Variant
    Dim idx() As Long
    Dim i As Long, j As Long, t As Long
    Dim n As Long

    n = UBound(arr) - LBound(arr) + 1
    ReDim idx(0 To n - 1)
    For i = 0 To n - 1
        idx(i) = i
    Next i

    For i = 0 To n - 2
        For j = 0 To n - 2 - i
            If Len(Trim$(arr(idx(j)))) < Len(Trim$(arr(idx(j + 1)))) Then
                t = idx(j): idx(j) = idx(j + 1): idx(j + 1) = t
            End If
        Next j
    Next i

    ByLengthDesc = idx
End Function


' ---------------------------------------------------------------------------
' Reading side - used by the ScheduleList refresh.
' ---------------------------------------------------------------------------

' Reads a value from a schedule's Metadata sheet by its header text.
Public Function ReadMeta(ByVal wb As Workbook, ByVal header As String) As Variant
    Dim wsMeta As Worksheet
    Dim lbl As Range

    ReadMeta = ""
    Set wsMeta = GetSheet(wb, SH_META)
    If wsMeta Is Nothing Then Exit Function

    Set lbl = FindLabel(wsMeta, header, 1, META_LAST_ROW + 5)
    If lbl Is Nothing Then Exit Function

    ReadMeta = CellValue(lbl.Offset(0, 1))
End Function


' Appends one row to a schedule's revision table.
Public Function AppendRevision(ByVal wb As Workbook, ByVal rev As String, ByVal status As String, _
                               ByVal issueDate As Variant, ByVal prep As String, _
                               ByVal chk As String, ByVal app As String, _
                               ByVal descr As String) As String
    Dim wsRev As Worksheet
    Dim lo As ListObject
    Dim target As Range
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

    Set target = NextRevisionRow(lo)
    If target Is Nothing Then
        AppendRevision = "could not find a row to write to"
        Exit Function
    End If

    SetCol lo, target, "Revision", rev
    SetCol lo, target, "Status", status
    SetCol lo, target, "Date", issueDate
    SetCol lo, target, "Prepared by", prep
    SetCol lo, target, "Checked by", chk
    SetCol lo, target, "Approved by", app
    SetCol lo, target, "Description", descr
End Function


' The row a new revision belongs in: the first empty row of the table, so the
' line lands directly under the previous revision.
'
' Revision tables are usually drawn with spare rows below the last entry.
' ListRows.Add would jump past them and leave a gap, which is what makes a
' title block look like it skipped a revision. Only when every row is used
' does the table actually grow by one.
Private Function NextRevisionRow(ByVal lo As ListObject) As Range
    Dim body As Range
    Dim i As Long, lastUsed As Long

    Set body = lo.DataBodyRange
    If body Is Nothing Then
        Set NextRevisionRow = lo.ListRows.Add.Range
        Exit Function
    End If

    For i = 1 To body.Rows.Count
        If Application.WorksheetFunction.CountA(body.Rows(i)) > 0 Then lastUsed = i
    Next i

    If lastUsed < body.Rows.Count Then
        Set NextRevisionRow = body.Rows(lastUsed + 1)
    Else
        Set NextRevisionRow = lo.ListRows.Add.Range
    End If
End Function


Private Sub SetCol(ByVal lo As ListObject, ByVal rowRange As Range, _
                   ByVal colName As String, ByVal v As Variant)
    Dim idx As Long
    On Error Resume Next
    idx = lo.ListColumns(colName).Index
    On Error GoTo 0
    If idx > 0 Then rowRange.Cells(1, idx).Value = v
End Sub


' ===========================================================================
' Headers and footers
'
' Used for the security classification banner (OFFICIAL, OFFICIAL-SENSITIVE,
' CONFIDENTIAL, or nothing). Set one workbook up by hand, then copy it out.
' ===========================================================================

' Reads one sheet's header and footer.
Public Function CaptureHF(ByVal ws As Worksheet) As HFSet
    Dim h As HFSet

    On Error GoTo Finish
    With ws.PageSetup
        h.LeftHeader = .LeftHeader
        h.CenterHeader = .CenterHeader
        h.RightHeader = .RightHeader
        h.LeftFooter = .LeftFooter
        h.CenterFooter = .CenterFooter
        h.RightFooter = .RightFooter
        h.DiffFirstPage = .DifferentFirstPageHeaderFooter
        h.OddAndEven = .OddAndEvenPagesHeaderFooter
        h.AlignMargins = .AlignMarginsHeaderFooter
        h.ScaleWithDoc = .ScaleWithDocHeaderFooter

        On Error Resume Next
        If h.DiffFirstPage Then
            h.FirstLeftHeader = .FirstPage.LeftHeader.Text
            h.FirstCenterHeader = .FirstPage.CenterHeader.Text
            h.FirstRightHeader = .FirstPage.RightHeader.Text
            h.FirstLeftFooter = .FirstPage.LeftFooter.Text
            h.FirstCenterFooter = .FirstPage.CenterFooter.Text
            h.FirstRightFooter = .FirstPage.RightFooter.Text
        End If
        If h.OddAndEven Then
            h.EvenLeftHeader = .EvenPage.LeftHeader.Text
            h.EvenCenterHeader = .EvenPage.CenterHeader.Text
            h.EvenRightHeader = .EvenPage.RightHeader.Text
            h.EvenLeftFooter = .EvenPage.LeftFooter.Text
            h.EvenCenterFooter = .EvenPage.CenterFooter.Text
            h.EvenRightFooter = .EvenPage.RightFooter.Text
        End If
        Err.Clear
        On Error GoTo Finish
    End With

    h.Valid = True

Finish:
    CaptureHF = h
End Function


' Writes one sheet's header and footer, then reads it back and checks it.
'
' An earlier version wrapped this in Application.PrintCommunication = False for
' speed. That is a documented optimisation but it queues the changes rather
' than applying them, and they can be dropped. Correctness first: 24 files is
' a few seconds either way.
'
' Returns "" on success, or a PROBLEM line naming what did not stick.
Public Function ApplyHF(ByVal ws As Worksheet, ByRef h As HFSet) As String
    If Not h.Valid Then Exit Function

    On Error GoTo HFError

    With ws.PageSetup
        .DifferentFirstPageHeaderFooter = h.DiffFirstPage
        .OddAndEvenPagesHeaderFooter = h.OddAndEven
        .AlignMarginsHeaderFooter = h.AlignMargins
        .ScaleWithDocHeaderFooter = h.ScaleWithDoc

        .LeftHeader = h.LeftHeader
        .CenterHeader = h.CenterHeader
        .RightHeader = h.RightHeader
        .LeftFooter = h.LeftFooter
        .CenterFooter = h.CenterFooter
        .RightFooter = h.RightFooter

        On Error Resume Next
        If h.DiffFirstPage Then
            .FirstPage.LeftHeader.Text = h.FirstLeftHeader
            .FirstPage.CenterHeader.Text = h.FirstCenterHeader
            .FirstPage.RightHeader.Text = h.FirstRightHeader
            .FirstPage.LeftFooter.Text = h.FirstLeftFooter
            .FirstPage.CenterFooter.Text = h.FirstCenterFooter
            .FirstPage.RightFooter.Text = h.FirstRightFooter
        End If
        If h.OddAndEven Then
            .EvenPage.LeftHeader.Text = h.EvenLeftHeader
            .EvenPage.CenterHeader.Text = h.EvenCenterHeader
            .EvenPage.RightHeader.Text = h.EvenRightHeader
            .EvenPage.LeftFooter.Text = h.EvenLeftFooter
            .EvenPage.CenterFooter.Text = h.EvenCenterFooter
            .EvenPage.RightFooter.Text = h.EvenRightFooter
        End If
        Err.Clear
        On Error GoTo HFError
    End With

    ApplyHF = VerifyHF(ws, h)
    Exit Function

HFError:
    ApplyHF = "PROBLEM: '" & ws.Name & "' error " & Err.Number & " - " & Err.Description & ". "
End Function


' Reads the six sections back and reports any that did not take. This is the
' difference between the tool saying it worked and it actually having worked.
Private Function VerifyHF(ByVal ws As Worksheet, ByRef h As HFSet) As String
    Dim bad As String

    On Error Resume Next
    With ws.PageSetup
        bad = bad & Mismatch("LeftHeader", .LeftHeader, h.LeftHeader)
        bad = bad & Mismatch("CenterHeader", .CenterHeader, h.CenterHeader)
        bad = bad & Mismatch("RightHeader", .RightHeader, h.RightHeader)
        bad = bad & Mismatch("LeftFooter", .LeftFooter, h.LeftFooter)
        bad = bad & Mismatch("CenterFooter", .CenterFooter, h.CenterFooter)
        bad = bad & Mismatch("RightFooter", .RightFooter, h.RightFooter)
    End With
    On Error GoTo 0

    If Len(bad) > 0 Then _
        VerifyHF = "PROBLEM: '" & ws.Name & "' did not take - " & bad
End Function


Private Function Mismatch(ByVal what As String, ByVal got As String, ByVal want As String) As String
    If StrComp(Trim$(got), Trim$(want), vbBinaryCompare) <> 0 Then
        Mismatch = what & " is [" & got & "] not [" & want & "]; "
    End If
End Function




' A one-line summary of a header/footer, for the log.
Public Function DescribeHF(ByRef h As HFSet) As String
    Dim hdr As String, ftr As String
    hdr = Trim$(h.LeftHeader & " " & h.CenterHeader & " " & h.RightHeader)
    ftr = Trim$(h.LeftFooter & " " & h.CenterFooter & " " & h.RightFooter)
    If Len(hdr) = 0 Then hdr = "(none)"
    If Len(ftr) = 0 Then ftr = "(none)"
    DescribeHF = "header " & hdr & " / footer " & ftr
End Function


' Copies headers and footers from one open workbook into another.
'
' Sheet matching, in order:
'   1. a source sheet with the same name
'   2. Front Cover and Revision Page to their namesakes
'   3. anything else to the source's schedule sheet
'
' Schedule tabs are not always called "Schedule", which is why the name match
' comes first and the role fallback second.
'
' Only header and footer properties are written. Orientation, paper size and
' margins are left alone, so a landscape schedule stays landscape. Header and
' footer text is positioned relative to whatever page the sheet is set to, so
' the same banner is correct on both.
Public Function CopyHeadersFootersTo(ByVal wbSrc As Workbook, ByVal wbTgt As Workbook, _
                                     ByVal imgPath As String, ByVal imgW As Double, _
                                     ByVal imgH As Double) As String
    Dim wsTgt As Worksheet, wsSrc As Worksheet
    Dim h As HFSet
    Dim log As String
    Dim n As Long

    For Each wsTgt In wbTgt.Worksheets
        If LCase$(Trim$(wsTgt.Name)) <> LCase$(SH_META) Then
            Set wsSrc = MatchSourceSheet(wbSrc, wsTgt)
            If wsSrc Is Nothing Then
                log = log & "No source sheet for '" & wsTgt.Name & "'. "
            Else
                h = CaptureHF(wsSrc)
                If h.Valid Then
                    log = log & ApplyHF(wsTgt, h)
                    log = log & ApplyHeaderImage(wsTgt, imgPath, imgW, imgH)
                    n = n + 1
                End If
            End If
        End If
    Next wsTgt

    CopyHeadersFootersTo = "Set " & n & " sheet(s). " & log
End Function


Private Function MatchSourceSheet(ByVal wbSrc As Workbook, ByVal wsTgt As Worksheet) As Worksheet
    Dim ws As Worksheet

    ' 1. same name
    Set ws = GetSheet(wbSrc, wsTgt.Name)
    If Not ws Is Nothing Then
        Set MatchSourceSheet = ws
        Exit Function
    End If

    ' 2. the two named sheets
    Select Case LCase$(Trim$(wsTgt.Name))
        Case LCase$(SH_FRONT)
            Set MatchSourceSheet = GetSheet(wbSrc, SH_FRONT)
            Exit Function
        Case LCase$(SH_REV)
            Set MatchSourceSheet = GetSheet(wbSrc, SH_REV)
            Exit Function
    End Select

    ' 3. whatever the source uses for its schedule
    Set MatchSourceSheet = GetSheet(wbSrc, "Schedule")
    If MatchSourceSheet Is Nothing Then Set MatchSourceSheet = FirstScheduleSheet(wbSrc)
End Function


' Puts an image in the top-right of the header at the given size.
'
' Header images cannot be read out of another workbook: Excel does not keep
' the original path, so PageSetup.RightHeaderPicture.Filename comes back empty
' for an embedded picture. It CAN be written, though, so the image is taken
' from a file the user picks rather than from the source workbook.
Public Function ApplyHeaderImage(ByVal ws As Worksheet, ByVal imgPath As String, _
                                 ByVal wPts As Double, ByVal hPts As Double) As String
    Dim rh As String

    If Len(imgPath) = 0 Then Exit Function

    On Error GoTo ImgError

    ' &G is the placeholder that makes the picture show. Keep any text that is
    ' already in the right section rather than throwing it away.
    rh = ws.PageSetup.RightHeader
    If InStr(1, rh, "&G", vbTextCompare) = 0 Then ws.PageSetup.RightHeader = "&G" & rh

    With ws.PageSetup.RightHeaderPicture
        .fileName = imgPath
        .LockAspectRatio = msoFalse
        .Width = wPts
        .Height = hPts
    End With

    If InStr(1, ws.PageSetup.RightHeader, "&G", vbTextCompare) = 0 Then
        ApplyHeaderImage = "PROBLEM: '" & ws.Name & "' header image did not take. "
    End If
    Exit Function

ImgError:
    ApplyHeaderImage = "PROBLEM: '" & ws.Name & "' image error " & Err.Number & _
                       " - " & Err.Description & ". "
End Function


' Native size of an image file, in points. Measured by dropping it on a
' scratch sheet, which is the only reliable way to ask Excel.
Public Function MeasureImage(ByVal imgPath As String, ByRef wPts As Double, _
                             ByRef hPts As Double) As Boolean
    Dim ws As Worksheet
    Dim shp As Shape

    If Not FileExists(imgPath) Then Exit Function

    On Error GoTo Clean
    Set ws = ThisWorkbook.Worksheets.Add
    Set shp = ws.Shapes.AddPicture(fileName:=imgPath, LinkToFile:=msoFalse, _
                                   SaveWithDocument:=msoTrue, Left:=0, Top:=0, _
                                   Width:=-1, Height:=-1)
    wPts = shp.Width
    hPts = shp.Height
    MeasureImage = (wPts > 0 And hPts > 0)

Clean:
    On Error Resume Next
    If Not ws Is Nothing Then ws.Delete
    On Error GoTo 0
End Function


' The first sheet that is not one of the three common ones.
Public Function FirstScheduleSheet(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If Not IsCommonSheet(ws) Then
            Set FirstScheduleSheet = ws
            Exit Function
        End If
    Next ws
End Function


' ===========================================================================
' Copying the common sheets from a reference schedule
'
' Replaces a target's Front Cover and Revision Page wholesale with the
' reference's, so a layout change made once - dropping a security
' classification, moving a logo, changing the wording - lands everywhere.
'
' What survives the copy, because it belongs to the document rather than the
' template:
'
'   the schedule title            ('Revision Page' A4)
'   the whole revision history    (the RevisionTable rows)
'   Document type, Delref Classification, BSUID, Trigger Events
'
' Everything else comes from the reference. The links are then rebuilt by the
' normal repair, so nothing points back at the reference workbook.
' ===========================================================================


' Returns "" when nothing needed saying, or a log fragment. The caller runs
' the repair afterwards and saves.
Public Function CopyCommonSheetsTo(ByVal wbSrc As Workbook, ByVal wbTgt As Workbook, _
                                   ByVal fileName As String) As String
    Dim wsSrcFront As Worksheet, wsSrcRev As Worksheet
    Dim wsRev As Worksheet, wsFront As Worksheet
    Dim rd As RevData
    Dim keeps As Object
    Dim title As String
    Dim idxFront As Long, idxRev As Long
    Dim log As String

    Set wsSrcFront = GetSheet(wbSrc, SH_FRONT)
    Set wsSrcRev = GetSheet(wbSrc, SH_REV)
    If wsSrcRev Is Nothing Then
        CopyCommonSheetsTo = "PROBLEM: the reference has no '" & SH_REV & "' sheet. "
        Exit Function
    End If

    Set wsRev = GetSheet(wbTgt, SH_REV)
    Set wsFront = GetSheet(wbTgt, SH_FRONT)

    ' --- everything that belongs to this document, before anything is deleted
    If wsRev Is Nothing Then
        ' A schedule that has never had these sheets. It gets them from the
        ' reference with an empty revision table, and its title is taken from
        ' the schedule sheet, or failing that from the file name.
        log = log & "No Revision Page - creating one from the reference. "
        title = TitleFromWorkbook(wbTgt, fileName)
        idxRev = wbTgt.Worksheets.Count + 1
        idxFront = idxRev
    Else
        title = CapturedTitle(wsRev)
        rd = CaptureRevisions(wsRev)
        Set keeps = CaptureKeeps(wsRev)

        If Not rd.Valid Then log = log & "No RevisionTable found to preserve. "
        If Len(title) = 0 Then title = TitleFromWorkbook(wbTgt, fileName)

        idxRev = wsRev.Index
        If wsFront Is Nothing Then idxFront = idxRev Else idxFront = wsFront.Index
    End If

    If Len(title) = 0 Then log = log & "PROBLEM: could not work out a schedule title. "

    ' --- swap the sheets -------------------------------------------------
    If Not wsFront Is Nothing Then wsFront.Delete
    If Not wsRev Is Nothing Then wsRev.Delete

    If Not wsSrcFront Is Nothing Then
        log = log & PlaceCopy(wsSrcFront, wbTgt, SH_FRONT, idxFront)
    End If
    log = log & PlaceCopy(wsSrcRev, wbTgt, SH_REV, idxRev)

    Set wsRev = GetSheet(wbTgt, SH_REV)
    If wsRev Is Nothing Then
        CopyCommonSheetsTo = log & "PROBLEM: the copied Revision Page did not arrive. "
        Exit Function
    End If

    ' --- put the document's own content back ------------------------------
    NameRevisionTable wsRev
    If Len(title) > 0 Then SetTitle wsRev, title
    log = log & RestoreRevisions(wsRev, rd)
    RestoreKeeps wsRev, keeps

    CopyCommonSheetsTo = log
End Function


' Copies one sheet in, gives it the right name and puts it back where the old
' one was. Excel names a copy "Revision Page (2)" when the name is taken, so
' the old sheet has to be gone first, which it is.
Private Function PlaceCopy(ByVal wsSrc As Worksheet, ByVal wbTgt As Workbook, _
                           ByVal wantName As String, ByVal wantIndex As Long) As String
    Dim ws As Worksheet

    On Error GoTo Failed
    wsSrc.Copy After:=wbTgt.Worksheets(wbTgt.Worksheets.Count)
    Set ws = wbTgt.Worksheets(wbTgt.Worksheets.Count)
    ws.Name = wantName

    If wantIndex >= 1 And wantIndex <= wbTgt.Worksheets.Count Then
        ws.Move Before:=wbTgt.Worksheets(wantIndex)
    End If
    Exit Function

Failed:
    PlaceCopy = "PROBLEM: could not place '" & wantName & "' - " & Err.Description & ". "
End Function


' A title for a schedule that has no Revision Page to take one from: the
' "SCHEDULE OF ..." heading on its own sheet, or the tail of the file name.
Private Function TitleFromWorkbook(ByVal wb As Workbook, ByVal fileName As String) As String
    Dim ws As Worksheet
    Dim t As Range
    Dim tail As String

    For Each ws In wb.Worksheets
        If Not IsCommonSheet(ws) Then
            Set t = FindTitleCell(ws)
            If Not t Is Nothing Then
                TitleFromWorkbook = Trim$(CStr(t.Value))
                Exit Function
            End If
        End If
    Next ws

    tail = fileName
    If InStrRev(tail, ".") > 1 Then tail = Left$(tail, InStrRev(tail, ".") - 1)
    If InStr(tail, " - ") > 0 Then tail = Mid$(tail, InStrRev(tail, " - ") + 3)
    TitleFromWorkbook = Trim$(tail)
End Function


Private Function CapturedTitle(ByVal wsRev As Worksheet) As String
    Dim t As Range
    Set t = FindTitleCell(wsRev)
    If t Is Nothing Then Exit Function
    CapturedTitle = Trim$(CStr(t.Value))
End Function


Private Sub SetTitle(ByVal wsRev As Worksheet, ByVal title As String)
    Dim t As Range
    Set t = FindTitleCell(wsRev)
    If t Is Nothing Then Exit Sub
    t.Value = title
End Sub


' The revision history, headers included so it can be written back by name
' even if the reference has reordered the columns.
Private Function CaptureRevisions(ByVal wsRev As Worksheet) As RevData
    Dim rd As RevData
    Dim lo As ListObject
    Dim body As Range
    Dim r As Long, c As Long, used As Long

    On Error Resume Next
    Set lo = wsRev.ListObjects("RevisionTable")
    On Error GoTo 0
    If lo Is Nothing Then
        CaptureRevisions = rd
        Exit Function
    End If

    rd.ColCount = lo.ListColumns.Count
    ReDim rd.Headers(1 To rd.ColCount)
    For c = 1 To rd.ColCount
        rd.Headers(c) = CStr(lo.ListColumns(c).Name)
    Next c

    Set body = lo.DataBodyRange
    If Not body Is Nothing Then
        For r = 1 To body.Rows.Count
            If Application.WorksheetFunction.CountA(body.Rows(r)) > 0 Then used = r
        Next r
    End If

    rd.RowCount = used
    If used > 0 Then
        ReDim rd.Cells(1 To used, 1 To rd.ColCount)
        For r = 1 To used
            For c = 1 To rd.ColCount
                rd.Cells(r, c) = body.Cells(r, c).Value
            Next c
        Next r
    End If

    rd.Valid = True
    CaptureRevisions = rd
End Function


' Writes the captured history into the freshly copied table. The body is
' always cleared first, so a schedule with no history does not inherit the
' reference's.
Private Function RestoreRevisions(ByVal wsRev As Worksheet, ByRef rd As RevData) As String
    Dim lo As ListObject
    Dim r As Long, c As Long, tgtCol As Long
    Dim need As Long

    On Error Resume Next
    Set lo = wsRev.ListObjects("RevisionTable")
    On Error GoTo 0
    If lo Is Nothing Then
        RestoreRevisions = "PROBLEM: no RevisionTable on the copied Revision Page. "
        Exit Function
    End If

    On Error GoTo Failed

    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.ClearContents

    If Not rd.Valid Or rd.RowCount = 0 Then Exit Function

    ' Grow the table if this schedule has more revisions than the reference.
    need = rd.RowCount
    Do While lo.ListRows.Count < need
        lo.ListRows.Add
    Loop

    For r = 1 To rd.RowCount
        For c = 1 To rd.ColCount
            tgtCol = ColumnIndexByName(lo, rd.Headers(c))
            If tgtCol > 0 Then lo.DataBodyRange.Cells(r, tgtCol).Value = rd.Cells(r, c)
        Next c
    Next r
    Exit Function

Failed:
    RestoreRevisions = "PROBLEM: revision history not fully restored - " & Err.Description & ". "
End Function


Private Function ColumnIndexByName(ByVal lo As ListObject, ByVal colName As String) As Long
    Dim i As Long
    For i = 1 To lo.ListColumns.Count
        If StrComp(lo.ListColumns(i).Name, colName, vbTextCompare) = 0 Then
            ColumnIndexByName = i
            Exit Function
        End If
    Next i
End Function


' Typed values next to the labels that are per document, not per template.
Private Function CaptureKeeps(ByVal wsRev As Worksheet) As Object
    Dim d As Object
    Dim names As Variant
    Dim i As Long
    Dim lbl As Range

    Set d = CreateObject("Scripting.Dictionary")
    Set CaptureKeeps = d

    names = Split(KEEP_LABELS, ",")
    For i = LBound(names) To UBound(names)
        Set lbl = FindLabel(wsRev, Trim$(CStr(names(i))))
        If Not lbl Is Nothing Then
            If Not lbl.Offset(0, 1).HasFormula Then
                d(Trim$(CStr(names(i)))) = lbl.Offset(0, 1).Value
            End If
        End If
    Next i
End Function


Private Sub RestoreKeeps(ByVal wsRev As Worksheet, ByVal keeps As Object)
    Dim k As Variant
    Dim lbl As Range

    If keeps Is Nothing Then Exit Sub
    For Each k In keeps.Keys
        Set lbl = FindLabel(wsRev, CStr(k))
        If Not lbl Is Nothing Then lbl.Offset(0, 1).Value = keeps(k)
    Next k
End Sub


' A copied table can arrive as "RevisionTable1" if the name was taken. It is
' not, because the old sheet is deleted first, but make sure of it: every
' revision formula in the workbook refers to it by name.
Private Sub NameRevisionTable(ByVal wsRev As Worksheet)
    Dim lo As ListObject
    If wsRev.ListObjects.Count = 0 Then Exit Sub
    Set lo = wsRev.ListObjects(1)
    If StrComp(lo.Name, "RevisionTable", vbTextCompare) = 0 Then Exit Sub
    On Error Resume Next
    lo.Name = "RevisionTable"
    On Error GoTo 0
End Sub
