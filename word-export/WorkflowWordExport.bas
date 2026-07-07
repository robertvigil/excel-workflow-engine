Attribute VB_Name = "WorkflowWordExport"
Option Explicit

' Start every top-level (Heading 1) section on a fresh page. Set via the
' Heading 1 style's "page break before" property (not manual breaks), so it
' reflows cleanly and leaves the TOC / cross-refs / bookmarks untouched. Set
' to False for a continuous document. Note: this also pushes the first section
' to its own page, leaving Title + TOC on page 1.
Private Const PAGE_BREAK_BEFORE_H1 As Boolean = True

' Put a page number (centered) in the footer of every page. It's a Word PAGE
' field referencing the same pagination the TOC uses, so the two correspond.
' Set to False to omit page numbers.
Private Const PAGE_NUMBERS As Boolean = True

' Show the outline number in front of each heading (e.g. "1.2 Step name") and
' in [[name]] cross-refs (e.g. "2.3.3 Configure the build"). This is display
' only -- the No. column still drives heading depth and outline order either
' way; setting False just suppresses the visible prefix. Set to True to number
' the output, False for name-only headings and cross-refs.
Private Const SHOW_NUMBERS As Boolean = False

' WorkflowWordExport ---------------------------------------------------------
'
' Companion to the excelscript-workflow-engine Office Script. Exports the
' active workflow sheet to a Word document:
'   * a Table of Contents at the top (clickable),
'   * one heading per step, nested by tree depth,
'   * each step's Notes as the body text beneath it, and
'   * screenshots embedded inline via ![[key]] image tokens, with optional
'     captions and an auto-built List of Figures.
'
' This is the "Excel macro drives Word via COM" path: it reads the sheet
' directly and builds the doc in one button-press -- no blob, no clipboard
' hand-off. It can't call the TypeScript engine, so it derives structure from
' the No. column: heading depth = dot-count of the number, order = outline
' order of the numbers. So the tree must be numbered first.
'
' PRECONDITION -- export is gated on a valid outline. Run Renumber first.
' Export refuses (and says what to fix) unless the No. column is present,
' every step is numbered, and the numbers form a valid outline. It never
' renumbers for you (that would mutate the source as a side effect).
'
' IMAGES -- put screenshots on a separate sheet and name it in a config row
' "Images Sheet | <sheet name>". That row is repeatable -- add several to group
' screenshots across multiple sheets; their shapes pool into one key map (first
' listed wins on a name collision). Give each screenshot a stable key by
' selecting it and typing the key into the Name Box (top-left). Reference it
' from any Notes cell with ![[key]] (or ![[key|Caption text]]). See README.md.
'
' LINKS -- reference another step from a Notes cell with [[Exact Step Name]];
' on export it becomes a clickable cross-reference that jumps to that step's
' heading and shows its current number + name (e.g. "2.3.3 Configure the
' build"), recomputed every run so it can't go stale. An unresolved name is
' left verbatim. Word-export only -- the source keeps the durable [[name]] token.
' A plain markdown link [text](url) also becomes a Word hyperlink to an external
' URL or file path (vs. [[name]], which links to a step heading in the doc).
'
' MARKDOWN -- Notes support lightweight markup, authored as plain text in the
' cell (so it has no 255-char limit): **bold**, *italic*, `code`; "- " bullet
' and "N. " numbered lists (indent 2 spaces per sub-level to nest); and
' simplified pipe tables -- a run of 2+ lines
' starting with "|", first row the header, the GFM "|---|" separator row
' optional. Native Excel cell formatting (bold applied with the toolbar) is NOT
' read -- that path is capped at ~255 chars by the Characters API, so markup is
' the dependable, length-agnostic route. See README.md.
'
' RUN IT: Alt+F11 -> Insert > Module -> paste this -> Alt+F8 ->
'         ExportWorkflowToWord. Windows desktop Excel + Word, macros enabled.
'
' CAVEATS: Windows-only (Excel->Word COM is unreliable on Mac, absent on Web).
'   Built-in style names ("Title", "Heading 1", "Caption") assume an English
'   Word UI. Embedding images uses the clipboard, so it overwrites whatever
'   you had copied.
' ---------------------------------------------------------------------------

Public Sub ExportWorkflowToWord()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim data As Variant
    data = ws.UsedRange.Value
    If Not IsArray(data) Then
        MsgBox "This sheet has no workflow to export.", vbExclamation, "Export to Word"
        Exit Sub
    End If

    Dim nRows As Long, nCols As Long
    nRows = UBound(data, 1)
    nCols = UBound(data, 2)

    ' --- locate the step-table header row and its columns -------------------
    Dim headerRow As Long, colNo As Long, colStep As Long, colLabel As Long, colNotes As Long
    Dim r As Long, c As Long

    For r = 1 To nRows
        For c = 1 To nCols
            If NormalizeHeader(data(r, c)) = "workflow step" Then headerRow = r: Exit For
        Next c
        If headerRow > 0 Then Exit For
    Next r

    If headerRow = 0 Then
        MsgBox "Couldn't find a step table (no 'Workflow Step' header) on this sheet.", _
               vbExclamation, "Export to Word"
        Exit Sub
    End If

    For c = 1 To nCols
        Select Case NormalizeHeader(data(headerRow, c))
            Case "no.":           colNo = c
            Case "workflow step": colStep = c
            Case "label":         colLabel = c
            Case "notes":         colNotes = c
        End Select
    Next c

    ' --- numbering gate: refuse and instruct -------------------------------
    If colNo = 0 Then
        MsgBox "Add a 'No.' column and run Renumber before exporting.", _
               vbExclamation, "Export to Word"
        Exit Sub
    End If

    ' Collect step rows (non-blank Workflow Step) as Array(no, displayLabel, notes),
    ' and remember every number so we can validate the outline.
    Dim steps As Collection: Set steps = New Collection
    Dim seen As Object:      Set seen = CreateObject("Scripting.Dictionary")
    Dim names As Object:     Set names = CreateObject("Scripting.Dictionary")
    names.CompareMode = vbTextCompare   ' [[name]] cross-refs resolve case-insensitively
    Dim idx As Long

    For r = headerRow + 1 To nRows
        Dim stepName As String
        stepName = Trim(CStr(data(r, colStep) & ""))
        If Len(stepName) > 0 Then
            Dim noVal As String
            noVal = Trim(CStr(data(r, colNo) & ""))
            If Len(noVal) = 0 Then
                MsgBox "Some steps aren't numbered. Run Renumber before exporting.", _
                       vbExclamation, "Export to Word"
                Exit Sub
            End If
            seen(noVal) = True
            idx = idx + 1

            ' Reset per-iteration: VBA does NOT clear Dim'd locals each loop pass,
            ' so without this disp/notes would retain the previous step's value
            ' whenever its source column is absent (e.g. no Label column).
            Dim disp As String
            disp = ""
            If colLabel > 0 Then disp = Trim(CStr(data(r, colLabel) & ""))
            If Len(disp) = 0 Then disp = stepName

            Dim notes As String
            notes = ""
            If colNotes > 0 Then notes = CStr(data(r, colNotes) & "")

            ' Each step gets a heading bookmark; a [[name]] cross-ref hyperlinks to
            ' it and displays "<number> <name>". Map: stepName -> Array(text, bookmark).
            Dim bmName As String
            bmName = "wfstep" & idx
            If Not names.Exists(stepName) Then _
                names.Add stepName, Array(HeadingLabel(noVal, disp), bmName)

            steps.Add Array(noVal, disp, notes, bmName)
        End If
    Next r

    If steps.Count = 0 Then
        MsgBox "There are no steps to export.", vbExclamation, "Export to Word"
        Exit Sub
    End If

    ' Outline validity: every multi-part number's parent prefix must exist.
    ' (Catches stale/partial numbering -- a believable-but-wrong TOC.)
    Dim i As Long
    For i = 1 To steps.Count
        Dim k As String: k = steps(i)(0)
        If InStr(k, ".") > 0 Then
            If Not seen.Exists(ParentPrefix(k)) Then
                MsgBox "Numbering is out of date (no parent for '" & k & "'). " & _
                       "Run Renumber before exporting.", vbExclamation, "Export to Word"
                Exit Sub
            End If
        End If
    Next i

    ' --- sort into outline order -------------------------------------------
    Dim arr() As Variant
    ReDim arr(1 To steps.Count)
    For i = 1 To steps.Count: arr(i) = steps(i): Next i
    SortByOutline arr

    Dim title As String
    title = ConfigValue(data, nRows, nCols, "workflow name")
    If Len(title) = 0 Then title = ws.name

    ' --- resolve images: shape-name -> Shape across every Images Sheet --------
    ' "Images Sheet" is repeatable (one sheet per row) so screenshots can be
    ' grouped across several sheets; shapes from all of them merge into one map.
    Dim images As Object
    Set images = CollectImages(ws.Parent, ConfigValues(data, nRows, nCols, "images sheet"))

    Dim anyCaptions As Boolean
    For i = 1 To steps.Count
        If HasResolvableCaption(CStr(steps(i)(2)), images) Then anyCaptions = True: Exit For
    Next i

    ' --- build the Word document -------------------------------------------
    Dim wd As Object
    On Error Resume Next
    Set wd = GetObject(, "Word.Application")
    On Error GoTo 0
    If wd Is Nothing Then Set wd = CreateObject("Word.Application")
    wd.Visible = True

    Dim doc As Object, sel As Object
    Set doc = wd.Documents.Add
    Set sel = wd.Selection

    ' Each Heading 1 starts on its own page (style property, not a hard break).
    If PAGE_BREAK_BEFORE_H1 Then
        On Error Resume Next
        doc.Styles("Heading 1").ParagraphFormat.PageBreakBefore = True
        On Error GoTo 0
    End If

    ' Tighten list spacing so a sub-list sits snug under its parent line
    ' (the Normal -> List Bullet boundary otherwise stacks both paragraphs'
    ' spacing, unlike consecutive same-style items which collapse).
    Dim lstNames As Variant, li As Long
    lstNames = Array("List Bullet", "List Bullet 2", "List Bullet 3", _
                     "List Bullet 4", "List Bullet 5", _
                     "List Number", "List Number 2", "List Number 3", _
                     "List Number 4", "List Number 5")
    For li = LBound(lstNames) To UBound(lstNames)
        On Error Resume Next
        With doc.Styles(lstNames(li)).ParagraphFormat
            .SpaceBefore = 0
            .SpaceAfter = 0
        End With
        On Error GoTo 0
    Next li

    ' Title
    sel.Style = "Title"
    sel.TypeText title
    sel.TypeParagraph

    ' Table of Contents (populates from the headings below on update).
    ' UseHyperlinks:=True adds the field's \h switch so entries are
    ' Ctrl+clickable jumps to their headings.
    Dim toc As Object
    Set toc = doc.TablesOfContents.Add(Range:=sel.Range, _
        UseHeadingStyles:=True, UpperHeadingLevel:=1, LowerHeadingLevel:=9, _
        UseHyperlinks:=True)
    sel.EndKey Unit:=6      ' wdStory -- past the TOC, to end of document
    sel.TypeParagraph

    ' Optional List of Figures (only if a captioned image will resolve).
    Dim tof As Object
    If anyCaptions Then
        sel.Style = "Normal"
        sel.Font.Bold = True
        sel.TypeText "List of Figures"
        sel.Font.Bold = False
        sel.TypeParagraph
        Set tof = doc.TablesOfFigures.Add(Range:=sel.Range, _
            caption:="Figure", IncludePageNumbers:=True, _
            RightAlignPageNumbers:=True, UseHyperlinks:=True)
        sel.EndKey Unit:=6
        sel.TypeParagraph
    End If

    ' Body: one heading per step (nested by depth), Notes beneath (with images).
    For i = 1 To UBound(arr)
        Dim depth As Long
        depth = OutlineDepth(CStr(arr(i)(0)))
        If depth > 9 Then depth = 9          ' Word has 9 heading levels
        sel.Style = "Heading " & depth
        Dim hStart As Long
        hStart = sel.Range.Start
        sel.TypeText HeadingLabel(CStr(arr(i)(0)), CStr(arr(i)(1)))  ' "1.2 Step name" or just "Step name"
        ' Bookmark the heading text so [[name]] cross-refs can hyperlink to it.
        doc.Bookmarks.Add name:=CStr(arr(i)(3)), Range:=doc.Range(hStart, sel.Range.Start)
        sel.TypeParagraph
        If Len(Trim(CStr(arr(i)(2)))) > 0 Then
            sel.Style = "Normal"
            RenderNotes sel, doc, CStr(arr(i)(2)), images, names
            sel.TypeParagraph
        End If
    Next i

    If PAGE_NUMBERS Then AddPageNumbers doc
    toc.Update
    If Not tof Is Nothing Then tof.Update
    doc.Fields.Update
    sel.HomeKey Unit:=6  ' back to the top
End Sub

' Add a centered page-number field to the document footer. The export uses
' page-break-before (not section breaks), so it's a single section -- one
' primary footer covers every page, and its numbers match the TOC because the
' TOC paginates the same document.
Private Sub AddPageNumbers(doc As Object)
    Dim ftr As Object
    Set ftr = doc.Sections(1).Footers(1)              ' wdHeaderFooterPrimary
    ftr.Range.Text = ""
    ftr.Range.Fields.Add Range:=ftr.Range, Type:=33   ' wdFieldPage
    ftr.Range.ParagraphFormat.Alignment = 1           ' wdAlignParagraphCenter
End Sub

' --- notes rendering -------------------------------------------------------

' Render a Notes string into the document. Three layers:
'   * block   -- split on line breaks; a run of 2+ lines starting with "|" is a
'                 simplified markdown table (see RenderTable), everything else is
'                 a normal line / list item.
'   * line    -- "- " => bullet, "N. " => numbered list, else Normal paragraph.
'   * inline  -- ![[key]] / ![[key|Caption]] image, [[name]] cross-ref link, and
'                 lightweight markup: **bold**, *italic*, `code` (see RenderInline).
Private Sub RenderNotes(sel As Object, doc As Object, ByVal notes As String, images As Object, names As Object)
    notes = Replace(notes, vbCrLf, vbLf)
    notes = Replace(notes, vbCr, vbLf)
    Dim lines() As String
    lines = Split(notes, vbLf)

    Dim i As Long, firstBlock As Boolean
    i = 0: firstBlock = True
    Do While i <= UBound(lines)
        ' Table block: this line AND the next both start with "|" (run of 2+).
        ' Nested If, not a single And: VBA's And doesn't short-circuit, so
        ' lines(i + 1) must be guarded behind the i < UBound test.
        Dim isTable As Boolean
        isTable = False
        If i < UBound(lines) Then
            If StartsWithPipe(lines(i)) And StartsWithPipe(lines(i + 1)) Then isTable = True
        End If
        If isTable Then
            Dim tbl As Collection: Set tbl = New Collection
            Do While i <= UBound(lines)
                If StartsWithPipe(lines(i)) Then
                    tbl.Add lines(i): i = i + 1
                Else
                    Exit Do
                End If
            Loop
            If Not firstBlock Then sel.TypeParagraph
            RenderTable sel, doc, tbl, images, names
            firstBlock = False
        Else
            If Not firstBlock Then sel.TypeParagraph
            RenderLine sel, doc, lines(i), images, names
            firstBlock = False
            i = i + 1
        End If
    Loop
End Sub

' One non-table line -> a paragraph. Detect a list prefix and pick the Word list
' style (nested by leading indent, 2 spaces per sub-level); otherwise Normal.
' Then render the line's text inline.
Private Sub RenderLine(sel As Object, doc As Object, ByVal line As String, images As Object, names As Object)
    Dim t As String: t = LTrim(line)
    Dim content As String
    If Left(t, 2) = "- " Then
        SetListStyle sel, "List Bullet", ListLevel(line)
        RenderInline sel, doc, Mid(t, 3), images, names
    ElseIf NumberedListContent(t, content) Then
        SetListStyle sel, "List Number", ListLevel(line)
        RenderInline sel, doc, content, images, names
    Else
        sel.Style = "Normal"
        RenderInline sel, doc, line, images, names
    End If
End Sub

' Nesting level from leading spaces: 2 spaces per sub-level (0-1 sp => 1,
' 2-3 => 2, ...), capped at 5 (Word's built-in leveled list styles stop there).
Private Function ListLevel(ByVal line As String) As Long
    Dim sp As Long: sp = 0
    Do While Mid(line, sp + 1, 1) = " "      ' Mid past end returns "", so this stops
        sp = sp + 1
    Loop
    ListLevel = sp \ 2 + 1
    If ListLevel > 5 Then ListLevel = 5
End Function

' Apply "List Bullet"/"List Number" (level 1) or "List Bullet 2".."5" etc.
' Falls back to the level-1 style if the leveled built-in isn't available.
Private Sub SetListStyle(sel As Object, ByVal base As String, ByVal level As Long)
    Dim nm As String
    If level <= 1 Then nm = base Else nm = base & " " & level
    On Error Resume Next
    sel.Style = nm
    If Err.Number <> 0 Then
        Err.Clear
        sel.Style = base
    End If
    On Error GoTo 0
End Sub

' Scan one line of text, resolving tokens and markup at the exact spot they
' appear: ![[key]] / ![[key|Caption]] image, [[name]] cross-ref, **bold**,
' *italic*, `code`. Literal text is buffered and typed in runs. Emphasis spans
' recurse so [[name]]/markup nest inside them; `code` is literal (no nesting).
' Any marker that doesn't close (or is space-flanked) is left as literal text.
Private Sub RenderInline(sel As Object, doc As Object, ByVal s As String, images As Object, names As Object)
    Dim pos As Long, n As Long, buf As String
    pos = 1: n = Len(s): buf = ""
    Do While pos <= n
        Dim m3 As String, m2 As String, ch As String
        m3 = Mid(s, pos, 3): m2 = Mid(s, pos, 2): ch = Mid(s, pos, 1)

        If m3 = "![[" Then
            Dim ei As Long: ei = InStr(pos + 3, s, "]]")
            If ei > 0 Then
                Flush sel, buf
                Dim inner As String: inner = Mid(s, pos + 3, ei - (pos + 3))
                Dim key As String, caption As String, bar As Long
                bar = InStr(inner, "|")
                If bar > 0 Then
                    key = Trim(Left(inner, bar - 1)): caption = Trim(Mid(inner, bar + 1))
                Else
                    key = Trim(inner): caption = ""
                End If
                EmitImage sel, doc, key, caption, images
                pos = ei + 2
            Else
                buf = buf & ch: pos = pos + 1
            End If

        ElseIf m2 = "[[" Then
            Dim el As Long: el = InStr(pos + 2, s, "]]")
            If el > 0 Then
                Flush sel, buf
                EmitLink sel, doc, Trim(Mid(s, pos + 2, el - (pos + 2))), names
                pos = el + 2
            Else
                buf = buf & ch: pos = pos + 1
            End If

        ElseIf ch = "[" Then            ' [text](url) markdown link ([[ handled above)
            Dim cs As Long, ce As Long
            cs = InStr(pos + 1, s, "]")
            If cs > 0 And Mid(s, cs + 1, 1) = "(" Then
                ce = InStr(cs + 2, s, ")")
            Else
                ce = 0
            End If
            If ce > 0 Then
                Flush sel, buf
                EmitUrlLink sel, doc, Mid(s, pos + 1, cs - (pos + 1)), _
                            Trim(Mid(s, cs + 2, ce - (cs + 2)))
                pos = ce + 1
            Else
                buf = buf & ch: pos = pos + 1
            End If

        ElseIf m2 = "**" Then
            ' Nested If, not one And: VBA's And doesn't short-circuit, so the
            ' Mid(s, cb - 1, ...) flanking test must be guarded behind cb > 0
            ' (an unmatched ** gives cb = 0 -> Mid(s, -1, ..) would error).
            Dim cb As Long: cb = InStr(pos + 2, s, "**")
            If cb > pos + 2 Then
                If Mid(s, pos + 2, 1) <> " " And Mid(s, cb - 1, 1) <> " " Then
                    Flush sel, buf
                    EmitEmphasis sel, doc, "bold", Mid(s, pos + 2, cb - (pos + 2)), images, names
                    pos = cb + 2
                Else
                    buf = buf & ch: pos = pos + 1
                End If
            Else
                buf = buf & ch: pos = pos + 1
            End If

        ElseIf ch = "*" Then
            Dim cit As Long: cit = InStr(pos + 1, s, "*")
            If cit > pos + 1 Then
                If Mid(s, pos + 1, 1) <> " " And Mid(s, cit - 1, 1) <> " " Then
                    Flush sel, buf
                    EmitEmphasis sel, doc, "italic", Mid(s, pos + 1, cit - (pos + 1)), images, names
                    pos = cit + 1
                Else
                    buf = buf & ch: pos = pos + 1
                End If
            Else
                buf = buf & ch: pos = pos + 1
            End If

        ElseIf ch = "`" Then
            Dim cc As Long: cc = InStr(pos + 1, s, "`")
            If cc > pos + 1 Then
                Flush sel, buf
                EmitEmphasis sel, doc, "code", Mid(s, pos + 1, cc - (pos + 1)), images, names
                pos = cc + 1
            Else
                buf = buf & ch: pos = pos + 1
            End If

        Else
            buf = buf & ch: pos = pos + 1
        End If
    Loop
    Flush sel, buf
End Sub

' Type any buffered literal text (in the current font) and clear the buffer.
Private Sub Flush(sel As Object, ByRef buf As String)
    If Len(buf) > 0 Then
        sel.TypeText buf
        buf = ""
    End If
End Sub

' Emit an emphasis span. bold/italic toggle the font and recurse so nested
' tokens/markup resolve; code types literally in a monospace font.
Private Sub EmitEmphasis(sel As Object, doc As Object, ByVal kind As String, ByVal content As String, images As Object, names As Object)
    Select Case kind
        Case "bold"
            Dim pb As Variant: pb = sel.Font.Bold
            sel.Font.Bold = True
            RenderInline sel, doc, content, images, names
            sel.Font.Bold = pb
        Case "italic"
            Dim pit As Variant: pit = sel.Font.Italic
            sel.Font.Italic = True
            RenderInline sel, doc, content, images, names
            sel.Font.Italic = pit
        Case "code"
            Dim pn As String: pn = sel.Font.name
            sel.Font.name = "Consolas"
            sel.TypeText content
            sel.Font.name = pn
    End Select
End Sub

' --- simplified markdown tables --------------------------------------------
'
' A table is a run of 2+ consecutive lines starting with "|". The first row is
' the header (no GFM "|---|" separator required) -- but a separator row, if
' present, is tolerated and skipped, so a pasted-in standard GFM table also
' works. Cells are split on "|" (outer pipes dropped) and rendered inline, so
' **bold** / *italic* / [[link]] work inside a cell. No alignment syntax, no
' "\|" escaping.
Private Sub RenderTable(sel As Object, doc As Object, tblLines As Collection, images As Object, names As Object)
    Dim rows As Collection: Set rows = New Collection
    Dim i As Long
    For i = 1 To tblLines.Count
        If Not IsSeparatorRow(CStr(tblLines(i))) Then rows.Add SplitRow(CStr(tblLines(i)))
    Next i
    If rows.Count = 0 Then Exit Sub

    Dim nCols As Long, r As Long, c As Long
    nCols = UBound(rows(1)) + 1
    For r = 2 To rows.Count                       ' widen for any ragged row
        If UBound(rows(r)) + 1 > nCols Then nCols = UBound(rows(r)) + 1
    Next r

    Dim wtbl As Object
    Set wtbl = doc.Tables.Add(Range:=sel.Range, NumRows:=rows.Count, NumColumns:=nCols)
    On Error Resume Next
    wtbl.Style = "Table Grid"                     ' simple bordered grid if available
    On Error GoTo 0
    wtbl.Borders.Enable = True

    For r = 1 To rows.Count
        Dim cells As Variant: cells = rows(r)
        For c = 1 To nCols
            Dim cellText As String
            If c - 1 <= UBound(cells) Then cellText = CStr(cells(c - 1)) Else cellText = ""
            wtbl.Cell(r, c).Range.Select
            sel.Collapse 1                        ' wdCollapseStart -- cursor at cell start
            sel.Style = "Normal"
            If r = 1 Then sel.Font.Bold = True    ' header row
            If Len(cellText) > 0 Then RenderInline sel, doc, cellText, images, names
            If r = 1 Then sel.Font.Bold = False
        Next c
    Next r

    ' Park the cursor just past the table so the next block renders after it.
    sel.SetRange Start:=wtbl.Range.End, End:=wtbl.Range.End
    sel.Collapse 0                                ' wdCollapseEnd
End Sub

' True if every line begins (after optional leading space) with "|".
Private Function StartsWithPipe(ByVal line As String) As Boolean
    StartsWithPipe = (Left(LTrim(line), 1) = "|")
End Function

' A GFM separator row: only |, -, :, space, and at least one dash.
Private Function IsSeparatorRow(ByVal line As String) As Boolean
    Dim t As String: t = Trim(line)
    If InStr(t, "-") = 0 Then Exit Function
    Dim i As Long, ch As String
    For i = 1 To Len(t)
        ch = Mid(t, i, 1)
        If ch <> "|" And ch <> "-" And ch <> ":" And ch <> " " Then Exit Function
    Next i
    IsSeparatorRow = True
End Function

' Split a table line into trimmed cells, dropping the optional outer pipes.
Private Function SplitRow(ByVal line As String) As Variant
    Dim t As String: t = Trim(line)
    If Left(t, 1) = "|" Then t = Mid(t, 2)
    If Len(t) > 0 And Right(t, 1) = "|" Then t = Left(t, Len(t) - 1)
    Dim parts() As String: parts = Split(t, "|")
    Dim i As Long
    For i = 0 To UBound(parts): parts(i) = Trim(parts(i)): Next i
    SplitRow = parts
End Function

' "12. text" => content "text" (numbered list item). Plain digit run + ". ".
Private Function NumberedListContent(ByVal s As String, ByRef content As String) As Boolean
    Dim p As Long: p = 1
    Do While p <= Len(s)
        Dim d As String: d = Mid(s, p, 1)
        If d < "0" Or d > "9" Then Exit Do
        p = p + 1
    Loop
    If p > 1 And Mid(s, p, 2) = ". " Then
        content = Mid(s, p + 2)
        NumberedListContent = True
    End If
End Function

' Resolve a [[Step Name]] cross-reference to a clickable hyperlink that jumps to
' the target step's heading and displays "<current number> <name>" (e.g.
' "2.3.3 Configure the build") -- recomputed every run so it never goes stale.
' Unknown name -> leave the token verbatim as a visible placeholder.
Private Sub EmitLink(sel As Object, doc As Object, ByVal name As String, names As Object)
    If Len(name) > 0 And names.Exists(name) Then
        Dim info As Variant
        info = names(name)                       ' Array(displayText, bookmarkName)
        Dim hl As Object
        Set hl = doc.Hyperlinks.Add(Anchor:=sel.Range, Address:="", _
            SubAddress:=CStr(info(1)), TextToDisplay:=CStr(info(0)))
        sel.SetRange Start:=hl.Range.End, End:=hl.Range.End   ' caret past the link
        sel.Font.Underline = 0                   ' wdUnderlineNone -- don't bleed link
        sel.Font.Color = -16777216               ' wdColorAutomatic -- formatting onward
    Else
        sel.TypeText "[[" & name & "]]"
    End If
End Sub

' Resolve a [text](url) markdown link to a Word hyperlink: address = the URL
' (a web address or a file path), visible text = "text" (or the URL itself if
' the text is empty). External target -- contrast with [[name]], which links to
' a step's heading inside the document.
Private Sub EmitUrlLink(sel As Object, doc As Object, ByVal linkText As String, ByVal url As String)
    If Len(url) = 0 Then
        sel.TypeText "[" & linkText & "]()"          ' empty URL -> leave as typed
        Exit Sub
    End If
    Dim disp As String: disp = linkText
    If Len(disp) = 0 Then disp = url
    Dim hl As Object
    Set hl = doc.Hyperlinks.Add(Anchor:=sel.Range, Address:=url, TextToDisplay:=disp)
    sel.SetRange Start:=hl.Range.End, End:=hl.Range.End   ' caret past the link
    sel.Font.Underline = 0                            ' don't bleed link styling onward
    sel.Font.Color = -16777216                        ' wdColorAutomatic
End Sub

' Resolve a ![[key]] token: embed the named screenshot (with an optional
' numbered caption), or -- if the key resolves to nothing -- leave the token
' verbatim as a visible "missing image" placeholder.
Private Sub EmitImage(sel As Object, doc As Object, ByVal key As String, ByVal caption As String, images As Object)
    If Len(key) > 0 And images.Exists(key) Then
        PasteShape sel, doc, images(key)
        If Len(caption) > 0 Then
            sel.TypeParagraph
            sel.Style = "Caption"
            sel.TypeText "Figure "
            ' SEQ field -> auto-numbers the figure and feeds the List of Figures
            sel.Fields.Add Range:=sel.Range, Type:=-1, _
                Text:="SEQ Figure \* ARABIC", PreserveFormatting:=False
            sel.TypeText ": " & caption
            sel.TypeParagraph
            sel.Style = "Normal"
        End If
    Else
        sel.Style = "Normal"
        If Len(caption) > 0 Then
            sel.TypeText "![[" & key & "|" & caption & "]]"
        Else
            sel.TypeText "![[" & key & "]]"
        End If
    End If
End Sub

' Copy an Excel shape to the clipboard and paste it inline at the cursor,
' clamping its width to the page text area so big screenshots fit.
Private Sub PasteShape(sel As Object, doc As Object, shp As Object)
    Dim before As Long
    before = doc.InlineShapes.Count
    shp.Copy
    DoEvents
    sel.Paste
    DoEvents
    If doc.InlineShapes.Count > before Then
        Dim ils As Object
        Set ils = doc.InlineShapes(doc.InlineShapes.Count)
        Dim maxW As Single
        maxW = doc.PageSetup.PageWidth - doc.PageSetup.LeftMargin - doc.PageSetup.RightMargin
        If maxW > 0 And ils.Width > maxW Then
            ils.Height = ils.Height * (maxW / ils.Width)
            ils.Width = maxW
        End If
    End If
End Sub

' --- helpers ---------------------------------------------------------------

Private Function NormalizeHeader(v As Variant) As String
    NormalizeHeader = LCase(Trim(CStr(v & "")))
End Function

' The visible label for a heading or cross-ref: "<number> <name>" when
' SHOW_NUMBERS is True, else just "<name>". The No. column still drives depth
' and order regardless -- this governs display only.
Private Function HeadingLabel(ByVal noVal As String, ByVal disp As String) As String
    If SHOW_NUMBERS Then
        HeadingLabel = noVal & " " & disp
    Else
        HeadingLabel = disp
    End If
End Function

Private Function OutlineDepth(ByVal noVal As String) As Long
    Dim n As Long, p As Long
    n = 1
    For p = 1 To Len(noVal)
        If Mid(noVal, p, 1) = "." Then n = n + 1
    Next p
    OutlineDepth = n
End Function

Private Function ParentPrefix(ByVal noVal As String) As String
    Dim p As Long
    p = InStrRev(noVal, ".")
    If p > 0 Then ParentPrefix = Left(noVal, p - 1)
End Function

' Zero-pad each numeric segment so plain string compare gives outline order
' (so "1.10" sorts after "1.9", not before it).
Private Function OutlineSortKey(ByVal noVal As String) As String
    Dim parts() As String, i As Long, s As String
    parts = Split(noVal, ".")
    For i = 0 To UBound(parts)
        s = s & Right(String(8, "0") & Trim(parts(i)), 8) & "."
    Next i
    OutlineSortKey = s
End Function

Private Sub SortByOutline(arr() As Variant)
    Dim i As Long, j As Long, tmp As Variant
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If OutlineSortKey(arr(j)(0)) < OutlineSortKey(arr(i)(0)) Then
                tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            End If
        Next j
    Next i
End Sub

' Read a config-block value by its Property label (value is the cell to its right).
Private Function ConfigValue(data As Variant, nRows As Long, nCols As Long, label As String) As String
    Dim r As Long, c As Long
    For r = 1 To nRows
        For c = 1 To nCols - 1
            If NormalizeHeader(data(r, c)) = label Then
                ConfigValue = Trim(CStr(data(r, c + 1) & ""))
                Exit Function
            End If
        Next c
    Next r
End Function

' Read every config-block value for a repeatable Property label (e.g. one row
' per "Images Sheet"), in top-to-bottom, left-to-right order. Blank values are
' skipped. Returns a (possibly empty) Collection of trimmed strings.
Private Function ConfigValues(data As Variant, nRows As Long, nCols As Long, label As String) As Collection
    Dim out As Collection: Set out = New Collection
    Set ConfigValues = out
    Dim r As Long, c As Long
    For r = 1 To nRows
        For c = 1 To nCols - 1
            If NormalizeHeader(data(r, c)) = label Then
                Dim v As String
                v = Trim(CStr(data(r, c + 1) & ""))
                If Len(v) > 0 Then out.Add v
            End If
        Next c
    Next r
End Function

' Build a case-insensitive map of shape name -> Shape, pooling the shapes on
' every named Images Sheet (the "Images Sheet" config row is repeatable). Sheets
' are scanned in config-row order and the first shape to claim a name wins, so a
' key that collides across sheets resolves to the earlier-listed sheet. A missing
' / unnamed sheet is skipped (not an error) -- its ![[key]] tokens then fall to
' the leave-verbatim placeholder, so a typo'd or absent sheet is visible in the
' output rather than crashing the export.
Private Function CollectImages(wb As Object, sheetNames As Collection) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Set CollectImages = d
    If sheetNames Is Nothing Then Exit Function

    Dim nm As Variant
    For Each nm In sheetNames
        Dim sheetName As String
        sheetName = Trim(CStr(nm & ""))
        If Len(sheetName) > 0 Then
            Dim sh As Object
            Set sh = Nothing
            On Error Resume Next
            Set sh = wb.Worksheets(sheetName)
            On Error GoTo 0
            If Not sh Is Nothing Then
                Dim shp As Object
                For Each shp In sh.Shapes
                    If Not d.Exists(shp.name) Then d.Add shp.name, shp
                Next shp
            End If
        End If
    Next nm
End Function

' True if any ![[key|Caption]] in the notes both has a caption and resolves to
' a real image -- i.e. it will produce a figure caption worth listing.
Private Function HasResolvableCaption(ByVal notes As String, images As Object) As Boolean
    Dim pos As Long, s As Long, e As Long, inner As String, bar As Long, key As String
    pos = 1
    Do
        s = InStr(pos, notes, "![[")
        If s = 0 Then Exit Do
        e = InStr(s + 3, notes, "]]")
        If e = 0 Then Exit Do
        inner = Mid(notes, s + 3, e - (s + 3))
        bar = InStr(inner, "|")
        If bar > 0 Then
            key = Trim(Left(inner, bar - 1))
            If Len(key) > 0 And images.Exists(key) Then
                HasResolvableCaption = True
                Exit Function
            End If
        End If
        pos = e + 2
    Loop
End Function




