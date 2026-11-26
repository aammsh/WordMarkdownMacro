'==========================================================
'  MarkdownToWord  -  turns pasted Markdown into real Word formatting
'  Version 2.3
'
'  Usage: paste Markdown text into a Word document, then run the
'         macro MarkdownToWord. It works on the whole active document.
'
'  Supports: headings 1-6, bold, italic, bold+italic, strikethrough,
'            inline code, ``` code blocks, bulleted and numbered lists
'            with nesting, task lists, block quotes, horizontal rules,
'            links, and tables (with column alignment).
'
'  NOTE ON ENCODING - this file is deliberately pure ASCII.
'  The VBA editor stores module text in a single-byte code page that
'  depends on Tools > Options > Editor Format > Script. When that is not
'  set to Arabic, every Arabic character typed or pasted into a module
'  is replaced by "?" - inside comments AND inside string literals.
'  So all Arabic text this macro needs (the table style name, the
'  message boxes) is built at run time from Unicode code points via the
'  helper U() at the bottom of this module. Keep it that way.
'==========================================================
Option Explicit

'------------------ settings you may change ------------------
' Right-to-left document (Arabic)
Private Const RTL_DOC As Boolean = True

' Optional table style name. ASCII only - see the encoding note above.
' Leave "" to use the built-in candidate list in TrySetTableStyle,
' which already tries the Arabic and English names of "Grid Table 4 -
' Accent 1" and falls back to plain borders.
Private Const TABLE_STYLE As String = ""

' Font used for code (Latin text)
Private Const CODE_FONT As String = "Consolas"

' Font used for the Arabic part of code. Word keeps a separate font for
' complex-script text, and Consolas carries no Arabic glyphs at all, so
' Arabic inside code must get its own monospaced font or Word silently
' substitutes an arbitrary one. Courier New is monospaced, ships with both
' Windows and Mac Word, and does cover Arabic.
Private Const CODE_FONT_BI As String = "Courier New"

' Delete leftover blank paragraphs at the end
Private Const REMOVE_EMPTY_LINES As Boolean = True

' Turn manual line breaks (Shift+Enter, common after pasting) into paragraphs
Private Const CONVERT_LINE_BREAKS As Boolean = True

' Treat "two spaces at end of line" as a line break, per the Markdown spec
Private Const HONOR_TWO_SPACE_BREAK As Boolean = False
'-------------------------------------------------------------

' Private-use characters that temporarily shield text from conversion
Private PH_STAR As String, PH_TICK As String, PH_HASH As String
Private PH_PIPE As String, PH_TILDE As String, PH_BRACK As String
Private PH_GT As String, PH_DASH As String, PH_DOT As String
Private PH_UNDER As String, PH_PERIOD As String


'==========================================================
'  The only entry point
'==========================================================
Public Sub MarkdownToWord()
    Dim doc As Document
    Dim ur As Object

    If Documents.Count = 0 Then
        ' Arabic: "No document is open."
        MsgBox U("644 627 20 64A 648 62C 62F 20 645 633 62A 646 62F 20 645 641 62A 648 62D 2E"), _
               vbExclamation
        Exit Sub
    End If

    Set doc = ActiveDocument
    InitPlaceholders

    On Error Resume Next
    Set ur = Application.UndoRecord
    ' Arabic: "Markdown conversion" - the single undo step name
    If Not ur Is Nothing Then ur.StartCustomRecord U("62A 62D 648 64A 644 20 4D 61 72 6B 64 6F 77 6E")
    On Error GoTo Fail

    Application.ScreenUpdating = False

    NormalizeText doc          ' 1.  first pass cleanup + shield \* escapes
    ProtectCodeBlocks doc      ' 2.  ``` code blocks
    ConvertTables doc          ' 3.  tables (before dash lines are removed)
    ConvertHeadings doc        ' 4.  headings
    ConvertRules doc           ' 5.  horizontal rules ---
    ConvertQuotes doc          ' 6.  block quotes >
    ConvertLists doc           ' 7.  lists
    ConvertInline doc          ' 8.  bold / italic / strike / code
    ConvertLinks doc           ' 9.  links
    RestoreProtected doc       ' 10. put the shielded characters back
    TidyUp doc                 ' 11. final cleanup

    Application.ScreenUpdating = True
    On Error Resume Next
    If Not ur Is Nothing Then ur.EndCustomRecord
    On Error GoTo 0

    ' Arabic: "Markdown converted successfully."
    MsgBox U("62A 645 20 62A 62D 648 64A 644 20 4D 61 72 6B 64 6F 77 6E 20 628 646 62C 627 62D 2E"), _
           vbInformation, "Markdown to Word"
    Exit Sub

Fail:
    Application.ScreenUpdating = True
    On Error Resume Next
    If Not ur Is Nothing Then ur.EndCustomRecord
    ' Arabic: "Conversion stopped because of an error:" / "Press Ctrl+Z to undo everything."
    MsgBox U("62A 648 642 641 20 627 644 62A 62D 648 64A 644 20 628 633 628 628 20 62E 637 623 3A") & vbCrLf & _
           Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
           U("627 636 63A 637 20 43 74 72 6C 2B 5A 20 644 644 62A 631 627 62C 639 20 639 646 20 643 644 20 645 627 20 62A 645 2E"), _
           vbExclamation, "Markdown to Word"
End Sub


'==========================================================
'  1. First pass cleanup
'==========================================================
Private Sub NormalizeText(doc As Document)
    If CONVERT_LINE_BREAKS Then RepPlain doc, "^l", "^p"

    ' Do not use the {1,} form: it depends on the system list separator
    ' and is rejected by Arabic Word. "@" means "one or more" and always works.
    If HONOR_TWO_SPACE_BREAK Then RepWild doc, "  @^13", "^l"

    ' Strip trailing spaces at end of lines
    RepWild doc, " @^13", "^p"

    ' Shield escaped markers \* \_ ... so they are not treated as formatting
    RepPlain doc, "\*", PH_STAR
    RepPlain doc, "\_", PH_UNDER
    RepPlain doc, "\#", PH_HASH
    RepPlain doc, "\|", PH_PIPE
    RepPlain doc, "\~", PH_TILDE
    RepPlain doc, "\[", PH_BRACK
    RepPlain doc, "\>", PH_GT
    RepPlain doc, "\-", PH_DASH
    RepPlain doc, "\`", PH_TICK
End Sub


'==========================================================
'  2. ``` code blocks
'==========================================================
Private Sub ProtectCodeBlocks(doc As Document)
    Dim i As Long, openAt As Long, inCode As Boolean
    Dim t As String

    i = 1
    openAt = 0
    inCode = False

    Do While i <= doc.Paragraphs.Count
        t = StripIndent(LineText(doc.Paragraphs(i)))
        If Left$(t, 3) = "```" Then
            If Not inCode Then
                inCode = True
                openAt = i
            Else
                inCode = False
                ApplyCodeBlock doc, openAt, i
                i = i - 2                 ' the two fence lines were deleted
            End If
        End If
        i = i + 1
    Loop
End Sub

Private Sub ApplyCodeBlock(doc As Document, ByVal fenceTop As Long, ByVal fenceBottom As Long)
    Dim k As Long, r As Range, s As String

    For k = fenceTop + 1 To fenceBottom - 1
        Set r = doc.Paragraphs(k).Range
        If r.End - r.Start > 1 Then
            r.MoveEnd wdCharacter, -1
            s = r.Text
            s = Replace(s, "*", PH_STAR)
            s = Replace(s, "`", PH_TICK)
            s = Replace(s, "#", PH_HASH)
            s = Replace(s, "|", PH_PIPE)
            s = Replace(s, "~", PH_TILDE)
            s = Replace(s, "[", PH_BRACK)
            s = Replace(s, ">", PH_GT)
            s = Replace(s, "-", PH_DASH)
            s = Replace(s, ".", PH_DOT)
            s = Replace(s, "_", PH_UNDER)
            r.Text = s
        End If

        With doc.Paragraphs(k)
            .Range.Font.Name = CODE_FONT
            .Range.Font.NameBi = CODE_FONT_BI
            .Range.Font.Bold = False
            .Range.Font.BoldBi = False
            .Range.Font.Italic = False
            .Range.Font.ItalicBi = False
            .Range.Shading.BackgroundPatternColor = RGB(244, 244, 244)
            .SpaceAfter = 0
            .SpaceBefore = 0
            .Alignment = wdAlignParagraphLeft
            On Error Resume Next
            .Range.ParagraphFormat.ReadingOrder = wdReadingOrderLtr
            On Error GoTo 0
        End With
    Next k

    ' Delete the bottom fence first so the top index stays valid
    doc.Paragraphs(fenceBottom).Range.Delete
    doc.Paragraphs(fenceTop).Range.Delete
End Sub

Private Sub RestoreProtected(doc As Document)
    RepPlain doc, PH_STAR, "*"
    RepPlain doc, PH_TICK, "`"
    RepPlain doc, PH_HASH, "#"
    RepPlain doc, PH_PIPE, "|"
    RepPlain doc, PH_TILDE, "~"
    RepPlain doc, PH_BRACK, "["
    RepPlain doc, PH_GT, ">"
    RepPlain doc, PH_DASH, "-"
    RepPlain doc, PH_DOT, "."
    RepPlain doc, PH_UNDER, "_"
End Sub


'==========================================================
'  3. Tables
'==========================================================
Private Sub ConvertTables(doc As Document)
    Dim i As Long, startP As Long, endP As Long

    i = doc.Paragraphs.Count
    Do While i >= 1
        If IsPipeRow(doc.Paragraphs(i)) Then
            endP = i
            Do While i > 1
                If IsPipeRow(doc.Paragraphs(i - 1)) Then
                    i = i - 1
                Else
                    Exit Do
                End If
            Loop
            startP = i
            If endP - startP >= 1 Then BuildTable doc, startP, endP
        End If
        i = i - 1
    Loop
End Sub

Private Function IsPipeRow(p As Paragraph) As Boolean
    Dim t As String
    If p.Range.Information(wdWithInTable) Then Exit Function
    t = Trim$(LineText(p))
    If Len(t) < 3 Then Exit Function
    If Left$(t, 1) <> "|" Then Exit Function
    IsPipeRow = (InStr(2, t, "|") > 0)
End Function

Private Sub BuildTable(doc As Document, ByVal startP As Long, ByVal endP As Long)
    Dim n As Long, k As Long, r As Long, c As Long
    Dim lines() As String, cells() As String
    Dim align() As Long
    Dim nCols As Long, nRows As Long
    Dim cellText() As String
    Dim sepIdx As Long

    n = endP - startP + 1
    If n < 2 Then Exit Sub

    ReDim lines(1 To n)
    For k = 1 To n
        lines(k) = Trim$(LineText(doc.Paragraphs(startP + k - 1)))
    Next k

    ' In Markdown the separator row |---|---| sits right under the header row.
    ' Look for it inside the block: lines starting with | may precede the table.
    sepIdx = 0
    For k = 2 To n
        If IsSeparatorRow(lines(k)) Then
            sepIdx = k
            Exit For
        End If
    Next k
    If sepIdx = 0 Then Exit Sub

    If sepIdx > 2 Then
        startP = startP + sepIdx - 2
        n = endP - startP + 1
        ReDim lines(1 To n)
        For k = 1 To n
            lines(k) = Trim$(LineText(doc.Paragraphs(startP + k - 1)))
        Next k
    End If

    align = ParseAlign(lines(2))

    nCols = UBound(align) + 1
    For k = 1 To n
        If k <> 2 Then
            cells = SplitRow(lines(k))
            If (UBound(cells) + 1) > nCols Then nCols = UBound(cells) + 1
        End If
    Next k
    If nCols < 1 Then Exit Sub

    nRows = n - 1
    ReDim cellText(1 To nRows, 1 To nCols)

    r = 0
    For k = 1 To n
        If k <> 2 Then
            r = r + 1
            cells = SplitRow(lines(k))
            For c = 1 To nCols
                If (c - 1) <= UBound(cells) Then
                    cellText(r, c) = cells(c - 1)
                Else
                    cellText(r, c) = ""
                End If
            Next c
        End If
    Next k

    ' Replace the text lines with a real Word table
    Dim rngBlock As Range
    Set rngBlock = doc.Range(doc.Paragraphs(startP).Range.Start, _
                             doc.Paragraphs(endP).Range.End)
    rngBlock.Delete

    Dim tbl As Table
    Set tbl = doc.Tables.Add(Range:=rngBlock, NumRows:=nRows, NumColumns:=nCols)

    For r = 1 To nRows
        For c = 1 To nCols
            If Len(cellText(r, c)) > 0 Then
                tbl.Cell(r, c).Range.Text = cellText(r, c)
            End If
        Next c
    Next r

    FormatTable tbl, align, nRows, nCols
End Sub

Private Sub FormatTable(tbl As Table, align() As Long, ByVal nRows As Long, ByVal nCols As Long)
    Dim r As Long, c As Long, a As Long
    Dim o As Object

    ' Right-to-left table direction. Late bound so an older Word that has no
    ' TableDirection property fails at run time and is swallowed, not at compile time.
    If RTL_DOC Then
        On Error Resume Next
        Set o = tbl
        o.TableDirection = 0                        ' wdTableDirectionRtl
        tbl.Range.ParagraphFormat.ReadingOrder = wdReadingOrderRtl
        Err.Clear
        On Error GoTo 0
    End If

    If Not TrySetTableStyle(tbl) Then
        With tbl.Borders
            .InsideLineStyle = wdLineStyleSingle
            .OutsideLineStyle = wdLineStyleSingle
            .InsideLineWidth = wdLineWidth050pt
            .OutsideLineWidth = wdLineWidth075pt
        End With
        tbl.Rows(1).Shading.BackgroundPatternColor = RGB(235, 235, 235)
    End If

    ' Column alignment as declared by the separator row
    For r = 1 To nRows
        For c = 1 To nCols
            If (c - 1) <= UBound(align) Then
                a = align(c - 1)
                If a >= 0 Then tbl.Cell(r, c).Range.ParagraphFormat.Alignment = a
            End If
        Next c
    Next r

    ' Header row: bold, and repeated when the table spans pages
    tbl.Rows(1).Range.Font.Bold = True
    tbl.Rows(1).Range.Font.BoldBi = True      ' Arabic text uses the Bi property
    On Error Resume Next
    tbl.Rows(1).HeadingFormat = True
    On Error GoTo 0

    With tbl.Range.ParagraphFormat
        .SpaceBefore = 0
        .SpaceAfter = 0
    End With

    tbl.PreferredWidthType = wdPreferredWidthPercent
    tbl.PreferredWidth = 100
    On Error Resume Next
    tbl.AutoFitBehavior wdAutoFitWindow
    On Error GoTo 0
End Sub

' Table style names are localised, so try several in order and fall back to
' plain borders. The U(...) entries are the Arabic names of
' "Grid Table 4 - Accent 1" and "Grid Table 1 Light - Accent 1".
Private Function TrySetTableStyle(tbl As Table) As Boolean
    Dim names As Variant, k As Long

    names = Array(TABLE_STYLE, _
                  U("62C 62F 648 644 20 634 628 643 629 20 34 20 2D 20 62A 645 64A 64A 632 20 31"), _
                  "Grid Table 4 - Accent 1", _
                  U("62C 62F 648 644 20 634 628 643 629 20 31 20 641 627 62A 62D 20 2D 20 62A 645 64A 64A 632 20 31"), _
                  "Grid Table 1 Light - Accent 1")

    For k = LBound(names) To UBound(names)
        If Len(names(k)) > 0 Then
            On Error Resume Next
            Err.Clear
            tbl.Style = names(k)
            If Err.Number = 0 Then
                On Error GoTo 0
                TrySetTableStyle = True
                Exit Function
            End If
            Err.Clear
            On Error GoTo 0
        End If
    Next k
End Function

Private Function SplitRow(ByVal t As String) As String()
    Dim s As String, parts() As String, k As Long

    s = Trim$(t)
    If Left$(s, 1) = "|" Then s = Mid$(s, 2)
    If Right$(s, 1) = "|" Then s = Left$(s, Len(s) - 1)

    ' Escaped \| was already swapped for a shielded character in NormalizeText,
    ' so the cell is not split there.
    parts = Split(s, "|")
    For k = LBound(parts) To UBound(parts)
        parts(k) = Trim$(parts(k))
        parts(k) = Replace(parts(k), "<br>", Chr(11))
        parts(k) = Replace(parts(k), "<br/>", Chr(11))
        parts(k) = Replace(parts(k), "<br />", Chr(11))
    Next k
    SplitRow = parts
End Function

Private Function IsSeparatorRow(ByVal t As String) As Boolean
    Dim cells() As String, k As Long, i As Long
    Dim c As String, ch As String, dashSeen As Boolean

    cells = SplitRow(t)
    If UBound(cells) < LBound(cells) Then Exit Function

    For k = LBound(cells) To UBound(cells)
        c = Trim$(cells(k))
        If Len(c) = 0 Then Exit Function
        dashSeen = False
        For i = 1 To Len(c)
            ch = Mid$(c, i, 1)
            If ch = "-" Then
                dashSeen = True
            ElseIf ch <> ":" And ch <> " " Then
                Exit Function
            End If
        Next i
        If Not dashSeen Then Exit Function
    Next k

    IsSeparatorRow = True
End Function

Private Function ParseAlign(ByVal t As String) As Long()
    Dim cells() As String, res() As Long, k As Long, c As String

    cells = SplitRow(t)
    ReDim res(0 To UBound(cells) - LBound(cells))

    For k = LBound(cells) To UBound(cells)
        c = Trim$(cells(k))
        If Left$(c, 1) = ":" And Right$(c, 1) = ":" Then
            res(k - LBound(cells)) = wdAlignParagraphCenter
        ElseIf Right$(c, 1) = ":" Then
            res(k - LBound(cells)) = wdAlignParagraphRight
        ElseIf Left$(c, 1) = ":" Then
            res(k - LBound(cells)) = wdAlignParagraphLeft
        Else
            res(k - LBound(cells)) = -1              ' no explicit alignment
        End If
    Next k

    ParseAlign = res
End Function


'==========================================================
'  4. Headings
'==========================================================
Private Sub ConvertHeadings(doc As Document)
    Dim i As Long, lvl As Long
    Dim t As String, body As String
    Dim r As Range

    For i = 1 To doc.Paragraphs.Count
        If Not doc.Paragraphs(i).Range.Information(wdWithInTable) Then
            t = LineText(doc.Paragraphs(i))
            lvl = HashLevel(t)
            If lvl > 0 Then
                body = Trim$(Mid$(StripIndent(t), lvl + 1))
                ' drop trailing hashes, as in: ## Title ##
                Do While Right$(body, 1) = "#"
                    body = RTrim$(Left$(body, Len(body) - 1))
                Loop

                Set r = doc.Paragraphs(i).Range
                r.MoveEnd wdCharacter, -1
                r.Text = body

                ' Built-in style ids, not names, so this works in any UI language
                doc.Paragraphs(i).Range.Font.Reset
                doc.Paragraphs(i).Style = doc.Styles(wdStyleHeading1 - (lvl - 1))
            End If
        End If
    Next i
End Sub

Private Function HashLevel(ByVal t As String) As Long
    Dim s As String, k As Long
    s = StripIndent(t)
    k = 0
    Do While k < Len(s)
        If Mid$(s, k + 1, 1) <> "#" Then Exit Do
        k = k + 1
    Loop
    If k < 1 Or k > 6 Then Exit Function
    If Len(s) = k Then Exit Function                  ' a line of hashes only
    If Mid$(s, k + 1, 1) <> " " Then Exit Function    ' a space must follow
    HashLevel = k
End Function


'==========================================================
'  5. Horizontal rules  ---  ***  ___
'==========================================================
Private Sub ConvertRules(doc As Document)
    Dim i As Long, t As String, r As Range

    For i = 1 To doc.Paragraphs.Count
        If Not doc.Paragraphs(i).Range.Information(wdWithInTable) Then
            t = Trim$(LineText(doc.Paragraphs(i)))
            If IsRuleLine(t) Then
                Set r = doc.Paragraphs(i).Range
                r.MoveEnd wdCharacter, -1
                If r.End > r.Start Then r.Delete
                With doc.Paragraphs(i)
                    .Range.Font.Reset
                    .SpaceBefore = 6
                    .SpaceAfter = 6
                    .Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
                    .Borders(wdBorderBottom).LineWidth = wdLineWidth075pt
                End With
            End If
        End If
    Next i
End Sub

Private Function IsRuleLine(ByVal t As String) As Boolean
    Dim ch As String, i As Long
    If Len(t) < 3 Then Exit Function
    ch = Left$(t, 1)
    If ch <> "-" And ch <> "*" And ch <> "_" Then Exit Function
    For i = 1 To Len(t)
        If Mid$(t, i, 1) <> ch Then Exit Function
    Next i
    IsRuleLine = True
End Function


'==========================================================
'  6. Block quotes  >
'==========================================================
Private Sub ConvertQuotes(doc As Document)
    Dim i As Long, t As String, s As String
    Dim wsChars As Long, wsWidth As Long, cut As Long

    For i = 1 To doc.Paragraphs.Count
        If Not doc.Paragraphs(i).Range.Information(wdWithInTable) Then
            t = LineText(doc.Paragraphs(i))
            MeasureIndent t, wsChars, wsWidth
            s = Mid$(t, wsChars + 1)
            If Left$(s, 1) = ">" Then
                cut = wsChars + 1
                If Mid$(s, 2, 1) = " " Then cut = cut + 1
                StripChars doc, i, cut

                With doc.Paragraphs(i)
                    .Range.Font.Italic = True
                    .Range.Font.Color = RGB(80, 80, 80)
                    If RTL_DOC Then
                        .RightIndent = 24
                        .Borders(wdBorderRight).LineStyle = wdLineStyleSingle
                        .Borders(wdBorderRight).LineWidth = wdLineWidth225pt
                        .Borders(wdBorderRight).Color = RGB(180, 180, 180)
                    Else
                        .LeftIndent = 24
                        .Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
                        .Borders(wdBorderLeft).LineWidth = wdLineWidth225pt
                        .Borders(wdBorderLeft).Color = RGB(180, 180, 180)
                    End If
                End With
            End If
        End If
    Next i
End Sub


'==========================================================
'  7. Lists
'==========================================================
Private Sub ConvertLists(doc As Document)
    Dim i As Long, t As String, s As String, body As String
    Dim wsChars As Long, wsWidth As Long, lvl As Long, mkLen As Long
    Dim prevWasBullet As Boolean, prevWasNumber As Boolean
    Dim boxCh As String, unit As Long

    ' Some sources indent nested items by two spaces, some by four
    unit = DetectListIndentUnit(doc)

    For i = 1 To doc.Paragraphs.Count
        If doc.Paragraphs(i).Range.Information(wdWithInTable) Then
            prevWasBullet = False
            prevWasNumber = False
        Else
            t = LineText(doc.Paragraphs(i))
            MeasureIndent t, wsChars, wsWidth
            s = Mid$(t, wsChars + 1)
            lvl = wsWidth \ unit
            If lvl > 4 Then lvl = 4

            mkLen = BulletMarkerLen(s)
            If mkLen > 0 Then
                StripChars doc, i, wsChars + mkLen
                body = LineText(doc.Paragraphs(i))
                If Left$(body, 4) = "[ ] " Or Left$(body, 4) = "[x] " Or Left$(body, 4) = "[X] " Then
                    If Mid$(body, 2, 1) = " " Then
                        boxCh = ChrW(&H2610)          ' empty ballot box
                    Else
                        boxCh = ChrW(&H2611)          ' ticked ballot box
                    End If
                    ReplaceHead doc, i, 4, boxCh & " "
                    If RTL_DOC Then
                        doc.Paragraphs(i).RightIndent = 18 * (lvl + 1)
                    Else
                        doc.Paragraphs(i).LeftIndent = 18 * (lvl + 1)
                    End If
                Else
                    ApplyList doc, i, True, lvl, prevWasBullet
                End If
                prevWasBullet = True
                prevWasNumber = False
            Else
                mkLen = NumberMarkerLen(s)
                If mkLen > 0 Then
                    StripChars doc, i, wsChars + mkLen
                    ApplyList doc, i, False, lvl, prevWasNumber
                    prevWasNumber = True
                    prevWasBullet = False
                Else
                    prevWasBullet = False
                    prevWasNumber = False
                End If
            End If
        End If
    Next i
End Sub

' The smallest indent found in front of a list item counts as one nesting level
Private Function DetectListIndentUnit(doc As Document) As Long
    Dim i As Long, t As String, s As String
    Dim wsChars As Long, wsWidth As Long, best As Long

    best = 0
    For i = 1 To doc.Paragraphs.Count
        If Not doc.Paragraphs(i).Range.Information(wdWithInTable) Then
            t = LineText(doc.Paragraphs(i))
            MeasureIndent t, wsChars, wsWidth
            If wsWidth > 0 Then
                s = Mid$(t, wsChars + 1)
                If BulletMarkerLen(s) > 0 Or NumberMarkerLen(s) > 0 Then
                    If best = 0 Or wsWidth < best Then best = wsWidth
                End If
            End If
        End If
    Next i

    If best < 1 Then best = 2
    DetectListIndentUnit = best
End Function

Private Sub ApplyList(doc As Document, ByVal i As Long, ByVal bulleted As Boolean, _
                      ByVal lvl As Long, ByVal continuePrev As Boolean)
    Dim lt As ListTemplate

    On Error Resume Next
    If bulleted Then
        Set lt = ListGalleries(wdBulletGallery).ListTemplates(1)
    Else
        Set lt = ListGalleries(wdNumberGallery).ListTemplates(1)
    End If

    doc.Paragraphs(i).Range.ListFormat.ApplyListTemplateWithLevel _
        ListTemplate:=lt, _
        ContinuePreviousList:=continuePrev, _
        ApplyTo:=wdListApplyToWholeList, _
        DefaultListBehavior:=wdWord10ListBehavior

    If lvl > 0 Then doc.Paragraphs(i).Range.ListFormat.ListLevelNumber = lvl + 1
    Err.Clear
    On Error GoTo 0
End Sub

Private Function BulletMarkerLen(ByVal s As String) As Long
    Dim ch As String
    If Len(s) < 2 Then Exit Function
    ch = Left$(s, 1)
    If ch <> "-" And ch <> "*" And ch <> "+" Then Exit Function
    If Mid$(s, 2, 1) <> " " Then Exit Function
    BulletMarkerLen = 2
End Function

Private Function NumberMarkerLen(ByVal s As String) As Long
    Dim k As Long, ch As String
    k = 0
    Do While k < Len(s)
        ch = Mid$(s, k + 1, 1)
        If ch < "0" Or ch > "9" Then Exit Do
        k = k + 1
    Loop
    If k = 0 Then Exit Function
    If Len(s) < k + 2 Then Exit Function
    ch = Mid$(s, k + 1, 1)
    If ch <> "." And ch <> ")" Then Exit Function
    If Mid$(s, k + 2, 1) <> " " Then Exit Function
    NumberMarkerLen = k + 2
End Function


'==========================================================
'  8. Inline formatting
'==========================================================
Private Sub ConvertInline(doc As Document)
    ' Word's wildcard engine refuses to let "[!x^13]@" span a full stop, so
    ' **version 1.2** or **jumla. Next** never matched and kept its markers.
    ' Hiding every "." behind a placeholder for the duration of this step is
    ' the only form that keeps the match bounded: "(*)" is greedy and runs
    ' past the closing marker, and "{1,}" is rejected outright by Arabic Word.
    RepPlain doc, ".", PH_PERIOD

    ' Order matters: longest marker first.
    ' [!*^13] means any character except an asterisk and except a paragraph
    ' mark, so a match cannot run across paragraphs. "@" means shortest match.
    RepWild doc, "\*\*\*([!*^13]@)\*\*\*", "\1", True, True             ' bold + italic
    RepWild doc, "\*\*([!*^13]@)\*\*", "\1", True, False                ' bold
    RepWild doc, "__([!_^13]@)__", "\1", True, False                    ' bold, underscores
    RepWild doc, "\*([!*^13]@)\*", "\1", False, True                    ' italic
    RepWild doc, "~~([!~^13]@)~~", "\1", False, False, True             ' strikethrough
    RepWild doc, "`([!`^13]@)`", "\1", False, False, False, _
            CODE_FONT, CODE_FONT_BI                                     ' inline code

    RepPlain doc, PH_PERIOD, "."

    ' HTML style line breaks
    RepPlain doc, "<br>", "^l"
    RepPlain doc, "<br/>", "^l"
    RepPlain doc, "<br />", "^l"
End Sub


'==========================================================
'  9. Links  [text](url)
'==========================================================
' No wildcards here: Word cannot exclude "]" inside a character class, so a
' correct link pattern is not expressible. Instead find "](" and expand
' outwards within the same paragraph.
Private Sub ConvertLinks(doc As Document)
    Dim r As Range, para As Range, lnk As Range
    Dim pText As String, ch As String
    Dim disp As String, addr As String
    Dim relPos As Long, ob As Long, cb As Long, k As Long
    Dim guard As Long, searchFrom As Long

    searchFrom = 0
    guard = 0

    Do
        guard = guard + 1
        If guard > 2000 Then Exit Do
        If searchFrom >= doc.Content.End Then Exit Do

        Set r = doc.Range(searchFrom, doc.Content.End)
        With r.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = "]("
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindStop
            .Format = False
            .MatchWildcards = False
            If Not .Execute Then Exit Do
        End With

        Set para = r.Paragraphs(1).Range
        pText = para.Text
        relPos = r.Start - para.Start + 1        ' position of "]" in the paragraph

        ob = 0
        For k = relPos - 1 To 1 Step -1
            ch = Mid$(pText, k, 1)
            If ch = "[" Then
                ob = k
                Exit For
            ElseIf ch = "]" Then
                Exit For
            End If
        Next k

        cb = 0
        For k = relPos + 2 To Len(pText)
            ch = Mid$(pText, k, 1)
            If ch = ")" Then
                cb = k
                Exit For
            ElseIf ch = "(" Or ch = vbCr Then
                Exit For
            End If
        Next k

        If ob > 0 And cb > 0 And relPos - ob - 1 > 0 Then
            disp = Mid$(pText, ob + 1, relPos - ob - 1)
            addr = Mid$(pText, relPos + 2, cb - relPos - 2)

            ' Image syntax ![alt](url): swallow the "!" and make it a text link
            If ob > 1 Then
                If Mid$(pText, ob - 1, 1) = "!" Then ob = ob - 1
            End If

            Set lnk = doc.Range(para.Start + ob - 1, para.Start + cb)

            On Error Resume Next
            doc.Hyperlinks.Add Anchor:=lnk, Address:=addr, TextToDisplay:=disp
            If Err.Number <> 0 Then
                Err.Clear
                lnk.Text = disp                  ' bad url: keep the text only
            End If
            On Error GoTo 0

            searchFrom = para.Start + ob - 1 + Len(disp)
        Else
            searchFrom = r.End
        End If
    Loop
End Sub


'==========================================================
'  11. Final cleanup
'==========================================================
Private Sub TidyUp(doc As Document)
    Dim i As Long

    If Not REMOVE_EMPTY_LINES Then Exit Sub

    For i = doc.Paragraphs.Count - 1 To 2 Step -1
        If Not doc.Paragraphs(i).Range.Information(wdWithInTable) Then
            If Len(Trim$(LineText(doc.Paragraphs(i)))) = 0 Then
                ' keep blank lines inside code blocks
                If doc.Paragraphs(i).Range.Font.Name <> CODE_FONT Then
                    ' and keep the separating paragraph that follows a table
                    If Not doc.Paragraphs(i - 1).Range.Information(wdWithInTable) Then
                        doc.Paragraphs(i).Range.Delete
                    End If
                End If
            End If
        End If
    Next i
End Sub


'==========================================================
'  Helpers
'==========================================================
' Builds a string from space separated Unicode code points in hex.
' Needed because the VBA editor cannot store non-Latin characters unless
' Tools > Options > Editor Format > Script matches them - see the note at
' the top of this module.
Private Function U(ByVal hexCodes As String) As String
    Dim parts() As String, i As Long, s As String
    parts = Split(hexCodes, " ")
    For i = LBound(parts) To UBound(parts)
        If Len(parts(i)) > 0 Then s = s & ChrW$(CLng("&H" & parts(i)))
    Next i
    U = s
End Function

Private Sub InitPlaceholders()
    PH_STAR = ChrW(&HE001)
    PH_TICK = ChrW(&HE002)
    PH_HASH = ChrW(&HE003)
    PH_PIPE = ChrW(&HE004)
    PH_TILDE = ChrW(&HE005)
    PH_BRACK = ChrW(&HE006)
    PH_GT = ChrW(&HE007)
    PH_DASH = ChrW(&HE008)
    PH_DOT = ChrW(&HE009)
    PH_UNDER = ChrW(&HE00A)
    PH_PERIOD = ChrW(&HE00B)
End Sub

' Paragraph text without the paragraph mark or the end-of-cell mark
Private Function LineText(p As Paragraph) As String
    Dim s As String, ch As String
    s = p.Range.Text
    Do While Len(s) > 0
        ch = Right$(s, 1)
        If ch = vbCr Or ch = vbLf Or ch = Chr(7) Then
            s = Left$(s, Len(s) - 1)
        Else
            Exit Do
        End If
    Loop
    LineText = s
End Function

' Measures leading whitespace: how many characters, and how wide (tab = 4)
Private Sub MeasureIndent(ByVal t As String, ByRef wsChars As Long, ByRef wsWidth As Long)
    Dim i As Long, ch As String
    wsChars = 0
    wsWidth = 0
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        If ch = " " Then
            wsChars = wsChars + 1
            wsWidth = wsWidth + 1
        ElseIf ch = vbTab Then
            wsChars = wsChars + 1
            wsWidth = wsWidth + 4
        Else
            Exit For
        End If
    Next i
End Sub

Private Function StripIndent(ByVal t As String) As String
    Dim a As Long, b As Long
    MeasureIndent t, a, b
    StripIndent = Mid$(t, a + 1)
End Function

' Deletes the first n characters of paragraph i
Private Sub StripChars(doc As Document, ByVal i As Long, ByVal n As Long)
    Dim r As Range, maxLen As Long
    If n <= 0 Then Exit Sub
    Set r = doc.Paragraphs(i).Range
    maxLen = r.End - r.Start - 1
    If n > maxLen Then n = maxLen
    If n <= 0 Then Exit Sub
    r.End = r.Start + n
    r.Delete
End Sub

' Replaces the first n characters of paragraph i with new text
Private Sub ReplaceHead(doc As Document, ByVal i As Long, ByVal n As Long, ByVal newText As String)
    Dim r As Range
    Set r = doc.Paragraphs(i).Range
    If r.End - r.Start - 1 < n Then Exit Sub
    r.End = r.Start + n
    r.Text = newText
End Sub

' Plain find and replace over the whole document
Private Sub RepPlain(doc As Document, ByVal findText As String, ByVal repText As String)
    Dim r As Range
    Set r = doc.Content
    With r.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = findText
        .Replacement.Text = repText
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchWildcards = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        On Error Resume Next
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        Err.Clear
        On Error GoTo 0
        .Execute Replace:=wdReplaceAll
    End With
End Sub

' Wildcard find and replace, with optional formatting for the result
Private Sub RepWild(doc As Document, ByVal pat As String, ByVal rep As String, _
                    Optional ByVal makeBold As Boolean = False, _
                    Optional ByVal makeItalic As Boolean = False, _
                    Optional ByVal makeStrike As Boolean = False, _
                    Optional ByVal fntName As String = "", _
                    Optional ByVal fntNameBi As String = "")
    Dim r As Range
    Set r = doc.Content
    With r.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = pat
        .Replacement.Text = rep
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchWildcards = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        On Error Resume Next
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        Err.Clear
        On Error GoTo 0

        ' Arabic (complex script) text ignores Bold/Italic and only shows
        ' BoldBi/ItalicBi, so both must be set or Arabic stays plain.
        If makeBold Then
            .Replacement.Font.Bold = True
            .Replacement.Font.BoldBi = True
            .Format = True
        End If
        If makeItalic Then
            .Replacement.Font.Italic = True
            .Replacement.Font.ItalicBi = True
            .Format = True
        End If
        If makeStrike Then
            .Replacement.Font.StrikeThrough = True
            .Format = True
        End If
        ' Latin and complex-script fonts are two separate properties, so an
        ' Arabic run keeps its old font unless NameBi is set as well.
        If Len(fntName) > 0 Then
            .Replacement.Font.Name = fntName
            .Format = True
        End If
        If Len(fntNameBi) > 0 Then
            .Replacement.Font.NameBi = fntNameBi
            .Format = True
        End If

        .Execute Replace:=wdReplaceAll
    End With
End Sub
