Attribute VB_Name = "WorkflowWordExport"
Option Explicit

' Start every top-level (Heading 1) section on a fresh page. Set via the
' Heading 1 style's "page break before" property (not manual breaks), so it
' reflows cleanly and leaves the TOC / cross-refs / bookmarks untouched. Set
' to False for a continuous document. Note: this also pushes the first section
' to its own page, leaving Title + TOC on page 1.
Private Const PAGE_BREAK_BEFORE_H1 As Boolean = True

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
' "Images Sheet | <sheet name>". Give each screenshot a stable key by
' selecting it and typing the key into the Name Box (top-left). Reference it
' from any Notes cell with ![[key]] (or ![[key|Caption text]]). See README.md.
'
' LINKS -- reference another step from a Notes cell with [[Exact Step Name]];
' on export it becomes a clickable cross-reference that jumps to that step's
' heading and shows its current number + name (e.g. "2.3.3 Configure the
' build"), recomputed every run so it can't go stale. An unresolved name is
' left verbatim. Word-export only -- the source keeps the durable [[name]] token.
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

            Dim disp As String
            If colLabel > 0 Then disp = Trim(CStr(data(r, colLabel) & ""))
            If Len(disp) = 0 Then disp = stepName

            Dim notes As String
            If colNotes > 0 Then notes = CStr(data(r, colNotes) & "")

            ' Each step gets a heading bookmark; a [[name]] cross-ref hyperlinks to
            ' it and displays "<number> <name>". Map: stepName -> Array(text, bookmark).
            Dim bmName As String
            bmName = "wfstep" & idx
            If Not names.Exists(stepName) Then _
                names.Add stepName, Array(noVal & " " & disp, bmName)

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
    If Len(title) = 0 Then title = ws.Name

    ' --- resolve images: shape-name -> Shape on the configured Images Sheet --
    Dim images As Object
    Set images = CollectImages(ws.Parent, ConfigValue(data, nRows, nCols, "images sheet"))

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
            Caption:="Figure", IncludePageNumbers:=True, _
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
        sel.TypeText arr(i)(0) & " " & arr(i)(1)   ' "1.2 Step name"
        ' Bookmark the heading text so [[name]] cross-refs can hyperlink to it.
        doc.Bookmarks.Add Name:=CStr(arr(i)(3)), Range:=doc.Range(hStart, sel.Range.Start)
        sel.TypeParagraph
        If Len(Trim(CStr(arr(i)(2)))) > 0 Then
            sel.Style = "Normal"
            RenderNotes sel, doc, CStr(arr(i)(2)), images, names
            sel.TypeParagraph
        End If
    Next i

    toc.Update
    If Not tof Is Nothing Then tof.Update
    doc.Fields.Update
    sel.HomeKey Unit:=6  ' back to the top
End Sub

' --- notes rendering -------------------------------------------------------

' Walk a Notes string, emitting literal text and resolving tokens at the exact
' spot they appear: ![[key]] / ![[key|Caption]] embed an image; [[name]]
' becomes a clickable cross-reference to that step's heading. (![[ ]] is
' detected by the "!" just before the brackets.)
Private Sub RenderNotes(sel As Object, doc As Object, ByVal notes As String, images As Object, names As Object)
    Dim pos As Long, p As Long, e As Long
    pos = 1
    Do
        p = InStr(pos, notes, "[[")
        If p = 0 Then
            EmitText sel, Mid(notes, pos)
            Exit Do
        End If
        e = InStr(p + 2, notes, "]]")
        If e = 0 Then
            EmitText sel, Mid(notes, pos)        ' unterminated -> literal
            Exit Do
        End If

        Dim isImage As Boolean, tokenStart As Long
        isImage = False
        If p > 1 Then
            If Mid(notes, p - 1, 1) = "!" Then isImage = True
        End If
        If isImage Then tokenStart = p - 1 Else tokenStart = p

        EmitText sel, Mid(notes, pos, tokenStart - pos)   ' text before the token

        Dim inner As String
        inner = Mid(notes, p + 2, e - (p + 2))

        If isImage Then
            Dim key As String, caption As String, bar As Long
            bar = InStr(inner, "|")
            If bar > 0 Then
                key = Trim(Left(inner, bar - 1))
                caption = Trim(Mid(inner, bar + 1))
            Else
                key = Trim(inner)
                caption = ""
            End If
            EmitImage sel, doc, key, caption, images
        Else
            EmitLink sel, doc, Trim(inner), names
        End If

        pos = e + 2
    Loop
End Sub

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

' Type literal text, turning in-cell line breaks into paragraphs.
Private Sub EmitText(sel As Object, ByVal s As String)
    If Len(s) = 0 Then Exit Sub
    sel.Style = "Normal"
    s = Replace(s, vbCrLf, vbLf)
    s = Replace(s, vbCr, vbLf)
    Dim parts() As String, i As Long
    parts = Split(s, vbLf)
    For i = 0 To UBound(parts)
        If i > 0 Then sel.TypeParagraph
        If Len(parts(i)) > 0 Then sel.TypeText parts(i)
    Next i
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

' Build a case-insensitive map of shape name -> Shape on the named Images
' Sheet. Missing / unnamed sheet -> empty map (every ![[key]] then falls to
' the leave-verbatim placeholder), so a typo'd or absent sheet is visible in
' the output rather than crashing the export.
Private Function CollectImages(wb As Object, ByVal sheetName As String) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Set CollectImages = d
    If Len(sheetName) = 0 Then Exit Function

    Dim sh As Object
    On Error Resume Next
    Set sh = wb.Worksheets(sheetName)
    On Error GoTo 0
    If sh Is Nothing Then Exit Function

    Dim shp As Object
    For Each shp In sh.Shapes
        If Not d.Exists(shp.Name) Then d.Add shp.Name, shp
    Next shp
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
