Attribute VB_Name = "AcceptanceRunner"
Option Explicit

Private passed As Long
Private failed As Long
Private lines As String
Private oldBlue As Long
Private nearBlue As Long
Private purple As Long

Public Sub RunAcceptanceTests()
    Dim deck As Presentation, page As Slide, second As Slide, third As Slide
    Dim tableShape As Shape, chartShape As Shape, report As Shape, codes As String
    Dim count As Long, changed As Long, expected As Long, valid As Boolean
    On Error GoTo Unexpected
    Set deck = Application.ActivePresentation
    If InStr(1, deck.Name, "AcceptanceHarness", vbTextCompare) = 0 Then
        MsgBox "请只在 AcceptanceHarness 测试稿中运行验收。", vbExclamation
        Exit Sub
    End If
    passed = 0: failed = 0: lines = vbNullString
    oldBlue = RGB(32, 79, 146): nearBlue = RGB(38, 80, 146): purple = RGB(170, 51, 204)
    BuildFixtures deck
    Set page = deck.Slides(2): Set second = deck.Slides(3): Set third = deck.Slides(4)
    Set tableShape = page.Shapes("TABLE")
    Set chartShape = page.Shapes("CHART")
    Verify ColorReplaceInPlace.HexToLong("112233") = RGB(17, 34, 51), "HEX_112233"
    Verify ColorReplaceInPlace.HexToLong("#aabbcc") = RGB(170, 187, 204), "HEX_HASH_LOWERCASE"
    Verify InvalidHex("12G"), "INVALID_HEX"
    Verify InvalidHex(""), "EMPTY_HEX"
    Verify InvalidHex("12345"), "INCOMPLETE_HEX"

    ColorReplaceInPlace.PrepareScan deck, oldBlue, 0, "2"
    Verify ColorReplaceInPlace.CandidateCount = 1, "EXACT_SCAN"
    Verify ColorReplaceInPlace.CandidateUses(1) >= 6, "SCAN_MULTIPLE_OBJECT_TYPES"
    changed = ColorReplaceInPlace.ApplySelection(deck, "204F92", "112233", "2")
    Verify changed >= 6, "APPLY_COUNT"
    Verify page.Shapes("FILL").Fill.ForeColor.RGB = RGB(17, 34, 51), "SHAPE_FILL"
    Verify page.Shapes("FILL").Line.ForeColor.RGB = nearBlue, "EXACT_PRESERVES_NEAR_LINE"
    Verify page.Shapes("TEXT").TextFrame.TextRange.Font.Color.RGB = RGB(17, 34, 51), "TEXT_COLOR"
    Verify page.Shapes("GROUP").GroupItems(1).Fill.ForeColor.RGB = RGB(17, 34, 51), "GROUP_FILL"
    Verify tableShape.Table.Cell(1, 1).Shape.Fill.ForeColor.RGB = RGB(17, 34, 51), "TABLE_FILL"
    Verify tableShape.Table.Cell(1, 2).Shape.TextFrame.TextRange.Font.Color.RGB = RGB(17, 34, 51), "TABLE_TEXT"
    Verify chartShape.Chart.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(17, 34, 51), "CHART_SERIES_FILL"
    Verify chartShape.Chart.SeriesCollection(1).Points(1).Format.Fill.ForeColor.RGB = RGB(17, 34, 51), "CHART_POINT_FILL"
    Verify chartShape.Chart.SeriesCollection(1).Points(2).Format.Fill.ForeColor.RGB = purple, "CHART_UNRELATED_POINT"
    Verify second.Shapes("FILL").Fill.ForeColor.RGB = oldBlue, "UNSELECTED_PAGE_3"
    Verify third.Shapes("FILL").Fill.ForeColor.RGB = oldBlue, "UNSELECTED_PAGE_4"
    Verify page.Shapes("PURPLE").Fill.ForeColor.RGB = purple, "UNRELATED_PURPLE"
    ColorReplaceInPlace.UndoLastColorChange Nothing
    Verify page.Shapes("FILL").Fill.ForeColor.RGB = oldBlue, "UNDO_SHAPE"
    Verify tableShape.Table.Cell(1, 1).Shape.Fill.ForeColor.RGB = oldBlue, "UNDO_TABLE_FILL"
    Verify chartShape.Chart.SeriesCollection(1).Format.Fill.ForeColor.RGB = oldBlue, "UNDO_CHART_SERIES"
    Verify chartShape.Chart.SeriesCollection(1).Points(1).Format.Fill.ForeColor.RGB = oldBlue, "UNDO_CHART_POINT"
    Verify chartShape.Chart.SeriesCollection(1).Points(2).Format.Fill.ForeColor.RGB = purple, "UNDO_CHART_UNRELATED_POINT"

    ColorReplaceInPlace.PrepareScan deck, oldBlue, 40, "2"
    codes = ScanCodes()
    Verify InStr(codes, "265092") > 0, "NEARBY_SCAN_LINE"
    Verify InStr(codes, "215092") > 0, "NEARBY_SCAN_FILL"
    changed = ColorReplaceInPlace.ApplySelection(deck, codes, "T2", "2")
    expected = deck.Designs(1).SlideMaster.Theme.ThemeColorScheme(msoThemeDark2).RGB
    Verify page.Shapes("FILL").Fill.ForeColor.ObjectThemeColor = msoThemeColorText2, "THEME_SLOT_4_CODE"
    Verify page.Shapes("FILL").Fill.ForeColor.RGB = expected, "THEME_SLOT_4_RGB"
    Verify page.Shapes("FILL").Line.ForeColor.ObjectThemeColor = msoThemeColorText2, "THEME_LINE"
    Verify tableShape.Table.Cell(1, 1).Borders(ppBorderTop).ForeColor.ObjectThemeColor = msoThemeColorText2, "TABLE_BORDER_THEME"
    lines = lines & "INFO | CHART_THEME_INDEX=" & CStr(chartShape.Chart.SeriesCollection(1).Format.Fill.ForeColor.ObjectThemeColor) & " | RGB=" & ColorReplaceInPlace.LongToHex(chartShape.Chart.SeriesCollection(1).Format.Fill.ForeColor.RGB) & " | EXPECTED=" & ColorReplaceInPlace.LongToHex(expected) & vbCrLf
    Verify chartShape.Chart.SeriesCollection(1).Format.Fill.ForeColor.RGB = expected, "CHART_THEME_RGB"
    lines = lines & "INFO | CHART_COLOR_TYPE=" & CStr(chartShape.Chart.SeriesCollection(1).Format.Fill.ForeColor.Type) & vbCrLf
    ColorReplaceInPlace.UndoLastColorChange Nothing
    Verify page.Shapes("FILL").Line.ForeColor.RGB = nearBlue, "UNDO_LINE"
    Verify tableShape.Table.Cell(1, 1).Borders(ppBorderTop).ForeColor.RGB = nearBlue, "UNDO_TABLE_BORDER"
    Verify page.Shapes("FILL").Fill.ForeColor.ObjectThemeColor = msoNotThemeColor, "UNDO_RETURNS_FIXED_RGB"
    Verify chartShape.Chart.SeriesCollection(1).Format.Fill.ForeColor.RGB = oldBlue, "UNDO_THEME_CHART"

    ColorReplaceInPlace.PrepareScan deck, oldBlue, 0, ""
    changed = ColorReplaceInPlace.ApplySelection(deck, "204F92", "112233", "")
    Verify second.Shapes("FILL").Fill.ForeColor.RGB = RGB(17, 34, 51), "ALL_PAGES_3"
    Verify third.Shapes("FILL").Fill.ForeColor.RGB = RGB(17, 34, 51), "ALL_PAGES_4"
    ColorReplaceInPlace.UndoLastColorChange Nothing
    Verify second.Shapes("FILL").Fill.ForeColor.RGB = oldBlue, "UNDO_ALL_PAGES"

    Verify InvalidPages(deck, "0"), "INVALID_PAGE_ZERO"
    Verify InvalidPages(deck, "99"), "INVALID_PAGE_RANGE"
    Verify InvalidPages(deck, "3-2"), "INVALID_PAGE_REVERSED"
    Verify InvalidRadius(deck, -1), "INVALID_NEGATIVE_RADIUS"
    Verify InvalidRadius(deck, 443), "INVALID_LARGE_RADIUS"
    ColorReplaceInPlace.PrepareScan deck, oldBlue, 0, "2"
    On Error Resume Next
    Err.Clear
    changed = ColorReplaceInPlace.ApplySelection(deck, "204F92", "12G", "2")
    valid = (Err.Number <> 0)
    Err.Clear
    On Error GoTo Unexpected
    Verify valid, "INVALID_TARGET_BLOCKED"
    Verify page.Shapes("FILL").Fill.ForeColor.RGB = oldBlue, "INVALID_TARGET_NO_MUTATION"
    RunNewFeatureChecks deck
    FinishReport deck
    MsgBox "验收完成：通过 " & CStr(passed) & "，失败 " & CStr(failed) & "。结果已写入测试稿第 1 页。", vbInformation, "Mac 批量改色·测试版验收"
    Exit Sub
Unexpected:
    failed = failed + 1
    lines = lines & "FAIL | UNEXPECTED | " & CStr(Err.Number) & " | " & Err.Description & vbCrLf
    On Error Resume Next
    FinishReport deck
    MsgBox "验收中断，结果已写入测试稿第 1 页：" & Err.Description, vbExclamation
End Sub

Private Sub Verify(ByVal condition As Boolean, ByVal name As String)
    If condition Then
        passed = passed + 1: lines = lines & "PASS | " & name & vbCrLf
    Else
        failed = failed + 1: lines = lines & "FAIL | " & name & vbCrLf
    End If
End Sub

Private Function InvalidHex(ByVal value As String) As Boolean
    Dim result As Long
    On Error GoTo Rejected
    result = ColorReplaceInPlace.HexToLong(value)
    Exit Function
Rejected:
    InvalidHex = True
End Function

Private Function InvalidPages(ByVal deck As Presentation, ByVal value As String) As Boolean
    On Error GoTo Rejected
    ColorReplaceInPlace.PrepareScan deck, oldBlue, 0, value
    Exit Function
Rejected:
    InvalidPages = True
End Function

Private Function InvalidRadius(ByVal deck As Presentation, ByVal value As Double) As Boolean
    On Error GoTo Rejected
    ColorReplaceInPlace.PrepareScan deck, oldBlue, value, "2"
    Exit Function
Rejected:
    InvalidRadius = True
End Function

Private Function ScanCodes() As String
    Dim i As Long, value As String
    For i = 1 To ColorReplaceInPlace.CandidateCount
        If Len(value) > 0 Then value = value & ","
        value = value & ColorReplaceInPlace.LongToHex(ColorReplaceInPlace.CandidateColor(i))
    Next i
    ScanCodes = value
End Function

Private Function AddBox(ByVal page As Slide, ByVal name As String, ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single, ByVal fill As Long, ByVal label As String) As Shape
    Dim item As Shape
    Set item = page.Shapes.AddShape(msoShapeRectangle, x, y, w, h)
    item.Name = name: item.Fill.Solid: item.Fill.ForeColor.RGB = fill
    item.Line.ForeColor.RGB = nearBlue
    item.TextFrame.TextRange.Text = label
    item.TextFrame.TextRange.Font.Name = "Microsoft YaHei"
    item.TextFrame.TextRange.Font.Size = 18
    item.TextFrame.TextRange.Font.Color.RGB = vbWhite
    Set AddBox = item
End Function

Private Sub BuildFixtures(ByVal deck As Presentation)
    Dim i As Long, j As Long, page As Slide, item As Shape, pair As Shape
    Dim group As Shape, tableShape As Shape, chartShape As Object
    For i = deck.Slides.Count To 1 Step -1
        If i <> 2 Then deck.Slides(i).Delete
    Next i
    For j = deck.Slides(1).Shapes.Count To 1 Step -1
        If deck.Slides(1).Shapes(j).Name <> "CHART" Then deck.Slides(1).Shapes(j).Delete
    Next j
    Set page = deck.Slides.Add(1, ppLayoutBlank)
    Set page = deck.Slides.Add(3, ppLayoutBlank)
    Set page = deck.Slides.Add(4, ppLayoutBlank)
    For i = 2 To 4
        Set page = deck.Slides(i)
        Set item = page.Shapes.AddTextbox(msoTextOrientationHorizontal, 30, 15, 880, 40)
        item.TextFrame.TextRange.Text = "Mac PowerPoint 批量改色 | 自制示例 " & CStr(i - 1)
        item.TextFrame.TextRange.Font.Name = "Microsoft YaHei"
        item.TextFrame.TextRange.Font.Size = 24
        Set item = AddBox(page, "FILL", 30, 75, 220, 75, oldBlue, "形状填充和线条")
        Set item = AddBox(page, "NEAR", 270, 75, 210, 75, RGB(33, 80, 146), "相近的蓝色")
        Set item = AddBox(page, "PURPLE", 500, 75, 210, 75, purple, "保留的其他颜色")
        Set item = page.Shapes.AddTextbox(msoTextOrientationHorizontal, 30, 165, 400, 35)
        item.Name = "TEXT": item.TextFrame.TextRange.Text = "文字颜色也可以一起替换"
        item.TextFrame.TextRange.Font.Name = "Microsoft YaHei"
        item.TextFrame.TextRange.Font.Size = 20
        item.TextFrame.TextRange.Font.Color.RGB = oldBlue
        If i = 2 Then
            Set item = AddBox(page, "GROUP_A", 30, 220, 95, 50, oldBlue, "组合 A")
            Set pair = AddBox(page, "GROUP_B", 135, 220, 95, 50, purple, "组合 B")
            Set group = page.Shapes.Range(Array("GROUP_A", "GROUP_B")).Group
            group.Name = "GROUP"
            Set tableShape = page.Shapes.AddTable(1, 2, 270, 220, 440, 70)
            tableShape.Name = "TABLE"
            For j = 1 To 2
                tableShape.Table.Cell(1, j).Shape.Fill.ForeColor.RGB = IIf(j = 1, oldBlue, vbWhite)
                tableShape.Table.Cell(1, j).Shape.TextFrame.TextRange.Text = "表格单元格 " & CStr(j)
                tableShape.Table.Cell(1, j).Shape.TextFrame.TextRange.Font.Color.RGB = IIf(j = 1, vbWhite, oldBlue)
                tableShape.Table.Cell(1, j).Borders(ppBorderTop).ForeColor.RGB = nearBlue
                tableShape.Table.Cell(1, j).Borders(ppBorderBottom).ForeColor.RGB = nearBlue
                tableShape.Table.Cell(1, j).Borders(ppBorderLeft).ForeColor.RGB = nearBlue
                tableShape.Table.Cell(1, j).Borders(ppBorderRight).ForeColor.RGB = nearBlue
            Next j
            Set chartShape = page.Shapes("CHART")
            chartShape.Name = "CHART"
            chartShape.Chart.HasTitle = True
            chartShape.Chart.ChartTitle.Text = "虚构示例：图表系列与数据点"
            chartShape.Chart.SeriesCollection(1).Format.Fill.ForeColor.RGB = oldBlue
            chartShape.Chart.SeriesCollection(1).Format.Line.ForeColor.RGB = nearBlue
            chartShape.Chart.SeriesCollection(1).Points(1).Format.Fill.ForeColor.RGB = oldBlue
            chartShape.Chart.SeriesCollection(1).Points(2).Format.Fill.ForeColor.RGB = purple
        End If
        Set item = page.Shapes.AddTextbox(msoTextOrientationHorizontal, 520, 330, 370, 75)
        item.Name = "CLIPBOARD": item.TextFrame.TextRange.Text = "112233"
        item.TextFrame.TextRange.Font.Name = "华文琥珀"
        item.TextFrame.TextRange.Font.Size = 36
        Set item = page.Shapes.AddTextbox(msoTextOrientationHorizontal, 520, 405, 370, 80)
        item.TextFrame.TextRange.Text = "自制测试数据，无客户资料。" & vbCrLf & "复制上方色号可测试纯文本粘贴。"
        item.TextFrame.TextRange.Font.Name = "Microsoft YaHei"
        item.TextFrame.TextRange.Font.Size = 15
    Next i
End Sub

Private Sub RunNewFeatureChecks(ByVal deck As Presentation)
    Dim page As Slide, item As Shape, tinted As Shape, fixed As Shape, text As Shape
    Dim inner As Shape, outer As Shape, count As Long, original As Long, valid As Boolean
    Dim i As Long, codes As Variant, theme As Long, mixed As Shape, tableShape As Shape
    Set page = deck.Slides(3)
    Set item = AddBox(page, "DIRECT", 30, 220, 140, 45, RGB(7, 8, 9), "不扫描替换")
    count = ColorReplaceInPlace.ApplyExact(deck, "070809", "987654", "3")
    Verify count >= 1, "DIRECT_WITHOUT_SCAN"
    Verify item.Fill.ForeColor.RGB = RGB(152, 118, 84), "DIRECT_EXACT_COLOR"
    Verify page.Shapes("NEAR").Fill.ForeColor.RGB = RGB(33, 80, 146), "DIRECT_IGNORES_NEAR_RADIUS"
    ColorReplaceInPlace.UndoLastColorChange Nothing
    Verify item.Fill.ForeColor.RGB = RGB(7, 8, 9), "UNDO_DIRECT"

    Set page = deck.Slides(4)
    Set item = AddBox(page, "NESTED_A", 30, 220, 95, 45, RGB(17, 34, 68), "内 A")
    Set item = AddBox(page, "NESTED_B", 135, 220, 95, 45, RGB(68, 85, 102), "内 B")
    item.TextFrame.TextRange.Font.Color.RGB = RGB(7, 8, 9)
    Set inner = page.Shapes.Range(Array("NESTED_A", "NESTED_B")).Group
    inner.Name = "INNER"
    Set item = AddBox(page, "NESTED_C", 245, 220, 95, 45, RGB(153, 170, 187), "外 C")
    Set outer = page.Shapes.Range(Array("INNER", "NESTED_C")).Group
    outer.Name = "NESTED"
    deck.Windows(1).View.GotoSlide 4
    lines = lines & "INFO | NESTED_BUILT" & vbCrLf
    FindGroupMember(outer, "NESTED_B").Select
    lines = lines & "INFO | NESTED_SELECTED" & vbCrLf
    Verify ColorReplaceInPlace.GetSelectedObjectColor("F") = RGB(68, 85, 102), "NESTED_CHILD_FILL_PICK"
    Verify ColorReplaceInPlace.GetSelectedObjectColor("L") = nearBlue, "NESTED_CHILD_LINE_PICK"
    FindGroupMember(outer, "NESTED_B").TextFrame.TextRange.Select
    Verify ColorReplaceInPlace.GetSelectedObjectColor("T") = RGB(7, 8, 9), "NESTED_CHILD_TEXT_PICK"
    Verify ColorReplaceInPlace.SelectedPagesSpec() = "4", "CURRENT_PAGE_TEXT_SELECTION"
    outer.Select
    On Error Resume Next
    Err.Clear
    original = ColorReplaceInPlace.GetSelectedObjectColor("F")
    valid = (Err.Number <> 0)
    Err.Clear
    On Error GoTo 0
    Verify valid, "WHOLE_GROUP_NOT_FALSE_COLOR"
    deck.Windows(1).ViewType = ppViewSlideSorter
    deck.Slides.Range(Array(2, 4)).Select
    lines = lines & "INFO | SELECTED_PAGES=" & ColorReplaceInPlace.SelectedPagesSpec() & vbCrLf
    Verify ColorReplaceInPlace.SelectedPagesSpec() = "2,4", "MULTI_SELECTED_PAGES"
    deck.Windows(1).ViewType = ppViewNormal
    deck.Windows(1).View.GotoSlide 3
    deck.Slides(3).Shapes("DIRECT").Select
    Verify ColorReplaceInPlace.SelectedPagesSpec() = "3", "CURRENT_PAGE_OVERRIDES_STALE_MULTI"

    Set page = deck.Slides(3)
    original = ColorReplaceInPlace.ColorSpecToRGB(deck, "A1", 3)
    Set item = AddBox(page, "THEME_BASE", 30, 285, 150, 45, oldBlue, "主题 A1")
    item.Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent1
    item.Line.ForeColor.ObjectThemeColor = msoThemeColorAccent1
    Set tinted = AddBox(page, "THEME_TINT", 195, 285, 150, 45, oldBlue, "主题浅色")
    tinted.Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent1
    tinted.Fill.ForeColor.TintAndShade = 0.35
    Set fixed = AddBox(page, "FIXED_SAME_RGB", 360, 285, 150, 45, original, "同色 RGB")
    Set text = page.Shapes.AddTextbox(msoTextOrientationHorizontal, 30, 350, 300, 30)
    text.Name = "THEME_TEXT": text.TextFrame.TextRange.Text = "主题文字"
    text.TextFrame.TextRange.Font.Color.ObjectThemeColor = msoThemeColorAccent1
    Set mixed = page.Shapes.AddTextbox(msoTextOrientationHorizontal, 350, 350, 200, 30)
    mixed.Name = "MIXED_TEXT": mixed.TextFrame.TextRange.Text = "ABCDEF"
    mixed.TextFrame.TextRange.Characters(1, 3).Font.Color.ObjectThemeColor = msoThemeColorAccent1
    mixed.TextFrame.TextRange.Characters(4, 3).Font.Color.ObjectThemeColor = msoThemeColorAccent2
    Set tableShape = page.Shapes.AddTable(1, 1, 30, 400, 280, 55)
    tableShape.Name = "THEME_TABLE"
    tableShape.Table.Cell(1, 1).Shape.Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent1
    tableShape.Table.Cell(1, 1).Shape.TextFrame.TextRange.Text = "ABCDEF"
    tableShape.Table.Cell(1, 1).Shape.TextFrame.TextRange.Characters(1, 3).Font.Color.ObjectThemeColor = msoThemeColorAccent1
    tableShape.Table.Cell(1, 1).Shape.TextFrame.TextRange.Characters(4, 3).Font.Color.ObjectThemeColor = msoThemeColorAccent2
    tableShape.Table.Cell(1, 1).Borders(ppBorderTop).ForeColor.ObjectThemeColor = msoThemeColorAccent1
    deck.Slides(2).Shapes("GROUP").GroupItems(1).Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent1
    ColorReplaceInPlace.PrepareScan deck, original, 0, "3"
    Verify ColorReplaceInPlace.CandidateCount >= 1, "SCAN_INCLUDES_THEME_COLORS"
    count = ColorReplaceInPlace.ApplyExact(deck, "A1", "A2", "2,3")
    Verify item.Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent2, "THEME_TO_THEME_FILL"
    Verify item.Line.ForeColor.ObjectThemeColor = msoThemeColorAccent2, "THEME_TO_THEME_LINE"
    Verify text.TextFrame.TextRange.Font.Color.ObjectThemeColor = msoThemeColorAccent2, "THEME_TO_THEME_TEXT"
    Verify Abs(tinted.Fill.ForeColor.TintAndShade - 0.35) < 0.001, "THEME_MIGRATION_PRESERVES_TINT"
    Verify tinted.Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent2, "THEME_TINT_SLOT"
    Verify fixed.Fill.ForeColor.ObjectThemeColor = msoNotThemeColor And fixed.Fill.ForeColor.RGB = original, "THEME_SLOT_IGNORES_FIXED_SAME_RGB"
    Verify deck.Slides(2).Shapes("GROUP").GroupItems(1).Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent2, "THEME_GROUP_CHILD"
    Verify deck.Slides(2).Shapes("GROUP").GroupItems(2).Fill.ForeColor.RGB = purple, "THEME_GROUP_PRESERVES_OTHER_CHILD"
    Verify tableShape.Table.Cell(1, 1).Shape.Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent2, "THEME_TABLE_FILL"
    Verify tableShape.Table.Cell(1, 1).Borders(ppBorderTop).ForeColor.ObjectThemeColor = msoThemeColorAccent2, "THEME_TABLE_BORDER"
    ColorReplaceInPlace.UndoLastColorChange Nothing
    Verify item.Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent1, "UNDO_RESTORES_THEME_SLOT"
    Verify tinted.Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent1 And Abs(tinted.Fill.ForeColor.TintAndShade - 0.35) < 0.001, "UNDO_RESTORES_THEME_TINT"
    Verify mixed.TextFrame.TextRange.Characters(1, 3).Font.Color.ObjectThemeColor = msoThemeColorAccent1, "UNDO_MERGED_TEXT_FIRST_RUN"
    Verify mixed.TextFrame.TextRange.Characters(4, 3).Font.Color.ObjectThemeColor = msoThemeColorAccent2, "UNDO_MERGED_TEXT_OTHER_RUN"
    Verify tableShape.Table.Cell(1, 1).Shape.TextFrame.TextRange.Characters(1, 3).Font.Color.ObjectThemeColor = msoThemeColorAccent1, "UNDO_MERGED_TABLE_TEXT"
    count = ColorReplaceInPlace.ApplyExact(deck, "A1", "112233", "3")
    Verify item.Fill.ForeColor.ObjectThemeColor = msoNotThemeColor And item.Fill.ForeColor.RGB = RGB(17, 34, 51), "THEME_TO_FIXED_RGB"
    Verify Abs(tinted.Fill.ForeColor.TintAndShade) < 0.001 And tinted.Fill.ForeColor.RGB = RGB(17, 34, 51), "FIXED_TARGET_CLEARS_TINT"
    ColorReplaceInPlace.UndoLastColorChange Nothing
    Verify item.Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent1, "UNDO_THEME_TO_RGB"
    count = ColorReplaceInPlace.ApplyExact(deck, ColorReplaceInPlace.LongToHex(original), "112233", "3")
    Verify item.Fill.ForeColor.RGB = RGB(17, 34, 51), "HEX_SOURCE_INCLUDES_THEME"
    Verify fixed.Fill.ForeColor.RGB = RGB(17, 34, 51), "HEX_SOURCE_INCLUDES_FIXED"
    ColorReplaceInPlace.UndoLastColorChange Nothing
    Verify item.Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent1, "UNDO_HEX_RESTORES_THEME"

    codes = Array("B1", "T1", "B2", "T2", "A1", "A2", "A3", "A4", "A5", "A6")
    Set page = deck.Slides(4)
    For i = 0 To 9
        Set item = AddBox(page, "SLOT_" & CStr(i), 30 + (i Mod 5) * 150, 300 + (i \ 5) * 55, 140, 45, oldBlue, CStr(codes(i)))
        theme = ColorReplaceInPlace.ThemeIndexFromCode(CStr(codes(i)))
        item.Fill.ForeColor.ObjectThemeColor = theme
        count = ColorReplaceInPlace.ApplyExact(deck, CStr(codes(i)), "A6", "4")
        Verify item.Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent6, "SOURCE_SLOT_" & CStr(codes(i))
        ColorReplaceInPlace.UndoLastColorChange Nothing
        Verify item.Fill.ForeColor.ObjectThemeColor = theme, "UNDO_SOURCE_SLOT_" & CStr(codes(i))
    Next i
End Sub

Private Function FindGroupMember(ByVal parent As Shape, ByVal objectName As String) As Shape
    Dim i As Long, child As Shape, found As Shape
    For i = 1 To parent.GroupItems.Count
        Set child = parent.GroupItems(i)
        If child.Name = objectName Then
            Set FindGroupMember = child
            Exit Function
        End If
        If child.Type = msoGroup Then
            Set found = FindGroupMember(child, objectName)
            If Not found Is Nothing Then
                Set FindGroupMember = found
                Exit Function
            End If
        End If
    Next i
End Function

Private Sub FinishReport(ByVal deck As Presentation)
    Dim item As Shape, entries As Variant, i As Long, mid As Long, leftText As String, rightText As String
    Set item = deck.Slides(1).Shapes.AddTextbox(msoTextOrientationHorizontal, 25, 12, 900, 45)
    item.Name = "ACCEPTANCE_SUMMARY"
    item.TextFrame.TextRange.Text = "RESULT | PASS=" & CStr(passed) & " | FAIL=" & CStr(failed)
    item.TextFrame.TextRange.Font.Size = 22
    entries = Split(lines, vbCrLf): mid = (UBound(entries) + 1) \ 2
    For i = 0 To UBound(entries)
        If i < mid Then leftText = leftText & entries(i) & vbCrLf Else rightText = rightText & entries(i) & vbCrLf
    Next i
    Set item = deck.Slides(1).Shapes.AddTextbox(msoTextOrientationHorizontal, 25, 65, 445, 460)
    item.Name = "ACCEPTANCE_DETAILS_A"
    item.TextFrame.TextRange.Text = leftText
    item.TextFrame.TextRange.Font.Name = "Menlo": item.TextFrame.TextRange.Font.Size = 8
    Set item = deck.Slides(1).Shapes.AddTextbox(msoTextOrientationHorizontal, 480, 65, 445, 460)
    item.Name = "ACCEPTANCE_DETAILS_B"
    item.TextFrame.TextRange.Text = rightText
    item.TextFrame.TextRange.Font.Name = "Menlo": item.TextFrame.TextRange.Font.Size = 8
    deck.Windows(1).View.GotoSlide 1
End Sub
