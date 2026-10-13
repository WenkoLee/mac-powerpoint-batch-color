' SPDX-License-Identifier: MIT
' Mac PowerPoint Batch Color Replacer — 0.1.0-beta.1
Attribute VB_Name = "ColorReplaceInPlace"
Option Explicit

Private Type ColorUndoItem
    Address As String
    OldRGB As Long
    OldTheme As Long
    OldTint As Single
End Type

Private foundColors() As Long
Private foundCounts() As Long
Private foundN As Long
Private pickedColors() As Long
Private pickedN As Long
Private pageMask() As Boolean
Private allPages As Boolean
Private baseColor As Long
Private radius As Double
Private doReplace As Boolean
Private useTheme As Boolean
Private newColor As Long
Private newTheme As Long
Private changedN As Long
Private undoItems() As ColorUndoItem
Private undoN As Long
Private undoDeckName As String

Public Sub ShowColorPanel(ByVal control As IRibbonControl)
    On Error GoTo Failed
    If Application.Presentations.Count = 0 Then
        MsgBox "请先打开一份演示文稿。", vbInformation, "Mac 批量改色（测试版）"
        Exit Sub
    End If
    ColorPickerForm.OpenForDeck Application.ActivePresentation
    Exit Sub
Failed:
    MsgBox "无法打开选色面板：" & Err.Description, vbExclamation, "Mac 批量改色（测试版）"
End Sub

Public Sub UndoLastColorChange(ByVal control As IRibbonControl)
    Dim deck As Presentation, color As ColorFormat, i As Long, restored As Long, expected As Long
    On Error GoTo Failed
    If undoN = 0 Then
        MsgBox "当前没有可撤销的插件改色操作。", vbInformation, "Mac 批量改色（测试版）"
        Exit Sub
    End If
    Set deck = Application.ActivePresentation
    If deck.FullName <> undoDeckName Then
        MsgBox "请先切回上次改色的演示文稿。", vbInformation, "Mac 批量改色（测试版）"
        Exit Sub
    End If
    expected = undoN
    For i = undoN To 1 Step -1
        Set color = ResolveColor(deck, undoItems(i).Address)
        If Not color Is Nothing Then
            On Error Resume Next
            RestoreColor color, undoItems(i).OldRGB, undoItems(i).OldTheme, undoItems(i).OldTint
            If Err.Number = 0 Then restored = restored + 1
            Err.Clear
            On Error GoTo Failed
        End If
    Next i
    If restored = expected Then
        undoN = 0
        Erase undoItems
        undoDeckName = vbNullString
        MsgBox "已恢复 " & CStr(restored) & " 处颜色（包括可访问的图表颜色）。", vbInformation, "Mac 批量改色（测试版）"
    Else
        MsgBox "仅恢复 " & CStr(restored) & "/" & CStr(expected) & " 处颜色。请检查页面，暂勿保存。", vbExclamation, "Mac 批量改色（测试版）"
    End If
    Exit Sub
Failed:
    MsgBox "撤销未完成：" & Err.Description, vbExclamation, "Mac 批量改色（测试版）"
End Sub

Public Function GetSelectedObjectColor(ByVal kind As String) As Long
    Dim sel As Selection, sh As Shape, color As ColorFormat
    Set sel = Application.ActiveWindow.Selection
    If sel.Type = ppSelectionText And UCase$(Trim$(kind)) = "T" Then
        Set color = sel.TextRange.Font.Color
    ElseIf sel.Type = ppSelectionShapes Or sel.Type = ppSelectionText Then
        If sel.ShapeRange.Count <> 1 Then Err.Raise 5, , "请只选中一个形状。"
        Set sh = sel.ShapeRange(1)
        Select Case UCase$(Trim$(kind))
            Case "F": Set color = sh.Fill.ForeColor
            Case "L": Set color = sh.Line.ForeColor
            Case "T"
                If Not sh.HasTextFrame Then Err.Raise 5, , "所选形状没有文字。"
                If Not sh.TextFrame.HasText Then Err.Raise 5, , "所选形状没有文字。"
                Set color = sh.TextFrame.TextRange.Runs(1).Font.Color
            Case Else: Err.Raise 5, , "请选择填充、线条或文字取色。"
        End Select
    Else
        Err.Raise 5, , "请先在页面中选中一个形状或一段文字。"
    End If
    GetSelectedObjectColor = color.RGB
End Function

Public Sub PrepareScan(ByVal deck As Presentation, ByVal oldColor As Long, _
                       ByVal distance As Double, ByVal pages As String)
    If distance < 0 Or distance > 442 Then Err.Raise 5, , "相近色范围请输入 0 到 442。"
    ParsePages deck, pages
    baseColor = oldColor
    radius = distance
    foundN = 0
    ReDim foundColors(1 To 16)
    ReDim foundCounts(1 To 16)
    doReplace = False
    WalkDeck deck
    SortFound
End Sub

Public Function CandidateCount() As Long
    CandidateCount = foundN
End Function

Public Function CandidateColor(ByVal index As Long) As Long
    If index < 1 Or index > foundN Then Err.Raise 5, , "候选颜色编号超出范围。"
    CandidateColor = foundColors(index)
End Function

Public Function CandidateUses(ByVal index As Long) As Long
    If index < 1 Or index > foundN Then Err.Raise 5, , "候选颜色编号超出范围。"
    CandidateUses = foundCounts(index)
End Function

Public Function ApplySelection(ByVal deck As Presentation, ByVal selectedHex As String, _
                               ByVal target As String, ByVal pages As String) As Long
    ParsePages deck, pages
    ParsePicked selectedHex
    ParseTarget target
    doReplace = True
    changedN = 0
    undoN = 0
    ReDim undoItems(1 To 64)
    undoDeckName = deck.FullName
    WalkDeck deck
    ApplySelection = changedN
End Function

Private Sub ParsePages(ByVal deck As Presentation, ByVal spec As String)
    Dim terms() As String, term As Variant, ends() As String
    Dim a As Long, b As Long, i As Long, n As Long
    n = deck.Slides.Count
    If n = 0 Then Err.Raise 5, , "演示文稿没有页面。"
    ReDim pageMask(1 To n)
    allPages = Len(Trim$(spec)) = 0
    If allPages Then
        For i = 1 To n: pageMask(i) = True: Next i
        Exit Sub
    End If
    terms = Split(Replace$(Trim$(spec), "，", ","), ",")
    For Each term In terms
        ends = Split(Trim$(CStr(term)), "-")
        If UBound(ends) > 1 Then Err.Raise 5, , "页码格式应为 1,3,5-7。"
        If Not IsNumeric(Trim$(ends(0))) Then Err.Raise 5, , "页码格式应为 1,3,5-7。"
        a = CLng(Trim$(ends(0)))
        b = a
        If UBound(ends) = 1 Then
            If Not IsNumeric(Trim$(ends(1))) Then Err.Raise 5, , "页码格式应为 1,3,5-7。"
            b = CLng(Trim$(ends(1)))
        End If
        If a < 1 Or b > n Or a > b Then Err.Raise 5, , "页码超出文件范围。"
        For i = a To b: pageMask(i) = True: Next i
    Next term
End Sub

Private Sub WalkDeck(ByVal deck As Presentation)
    Dim i As Long, j As Long, k As Long, d As Design, layout As CustomLayout
    For i = 1 To deck.Slides.Count
        If pageMask(i) Then
            For j = 1 To deck.Slides(i).Shapes.Count
                WalkShape deck.Slides(i).Shapes(j), "S|" & CStr(deck.Slides(i).SlideID) & "|" & CStr(j)
            Next j
            On Error Resume Next
            Touch deck.Slides(i).Background.Fill.ForeColor, "S|" & CStr(deck.Slides(i).SlideID) & "||BG"
            On Error GoTo 0
        End If
    Next i
    If Not allPages Then Exit Sub
    For i = 1 To deck.Designs.Count
        Set d = deck.Designs(i)
        For j = 1 To d.SlideMaster.Shapes.Count
            WalkShape d.SlideMaster.Shapes(j), "M|" & CStr(i) & "|" & CStr(j)
        Next j
        For k = 1 To d.SlideMaster.CustomLayouts.Count
            Set layout = d.SlideMaster.CustomLayouts(k)
            For j = 1 To layout.Shapes.Count
                WalkShape layout.Shapes(j), "L|" & CStr(i) & ":" & CStr(k) & "|" & CStr(j)
            Next j
        Next k
    Next i
End Sub

Private Sub WalkShape(ByVal sh As Shape, ByVal address As String)
    Dim i As Long, r As Long, c As Long, b As Long, cellShape As Shape
    On Error Resume Next
    Touch sh.Fill.ForeColor, address & "|F"
    Touch sh.Line.ForeColor, address & "|L"
    If sh.Type = msoGroup Then
        For i = 1 To sh.GroupItems.Count
            WalkShape sh.GroupItems(i), address & "." & CStr(i)
        Next i
    End If
    If sh.HasTable Then
        For r = 1 To sh.Table.Rows.Count
            For c = 1 To sh.Table.Columns.Count
                Set cellShape = sh.Table.Cell(r, c).Shape
                Touch cellShape.Fill.ForeColor, address & "|C:" & CStr(r) & ":" & CStr(c) & ":F"
                Touch cellShape.Line.ForeColor, address & "|C:" & CStr(r) & ":" & CStr(c) & ":L"
                If cellShape.HasTextFrame Then
                    If cellShape.TextFrame.HasText Then _
                        WalkText cellShape.TextFrame.TextRange, address & "|C:" & CStr(r) & ":" & CStr(c) & ":X"
                End If
                For b = 1 To 6
                    Touch sh.Table.Cell(r, c).Borders(b).ForeColor, address & "|B:" & CStr(r) & ":" & CStr(c) & ":" & CStr(b)
                Next b
            Next c
        Next r
    End If
    If sh.HasTextFrame Then
        If sh.TextFrame.HasText Then WalkText sh.TextFrame.TextRange, address & "|X"
    End If
    If sh.HasChart Then WalkChart sh.Chart, address
    On Error GoTo 0
End Sub

Private Sub WalkText(ByVal t As TextRange, ByVal address As String)
    Dim i As Long
    On Error Resume Next
    For i = 1 To t.Runs.Count
        Touch t.Runs(i).Font.Color, address & ":" & CStr(i)
    Next i
    On Error GoTo 0
End Sub

Private Sub WalkChart(ByVal chart As Object, ByVal address As String)
    Dim i As Long, j As Long, k As Long, series As Object, point As Object
    Dim before() As ColorUndoItem, savedN As Long, savedTheme As Long
    Dim preserveFill As Boolean, preserveLine As Boolean, prop As String, fields As Variant
    On Error Resume Next
    savedTheme = newTheme
    If useTheme Then newTheme = ChartThemeIndex(newTheme)
    For i = 1 To chart.SeriesCollection.Count
        Set series = chart.SeriesCollection(i)
        preserveFill = False: preserveLine = False: savedN = 0
        If doReplace Then
            preserveFill = WillReplace(series.Format.Fill.ForeColor)
            preserveLine = WillReplace(series.Format.Line.ForeColor)
        End If
        ReDim before(1 To series.Points.Count * 2 + 1)
        If preserveFill Or preserveLine Then
            ' Changing a series resets point overrides on Mac. Capture points
            ' first so reverse undo restores the series before its points.
            For j = 1 To series.Points.Count
                Set point = series.Points(j)
                If preserveFill Then SnapshotPoint point.Format.Fill.ForeColor, address & "|Q:" & CStr(i) & ":" & CStr(j) & ":F", before, savedN
                If preserveLine Then SnapshotPoint point.Format.Line.ForeColor, address & "|Q:" & CStr(i) & ":" & CStr(j) & ":L", before, savedN
            Next j
        End If
        Touch series.Format.Fill.ForeColor, address & "|Q:" & CStr(i) & ":0:F"
        Touch series.Format.Line.ForeColor, address & "|Q:" & CStr(i) & ":0:L"
        For k = 1 To savedN
            fields = Split(before(k).Address, ":")
            prop = fields(3)
            j = CLng(fields(2))
            Set point = series.Points(j)
            If prop = "F" Then
                RestoreColor point.Format.Fill.ForeColor, before(k).OldRGB, before(k).OldTheme, before(k).OldTint
            Else
                RestoreColor point.Format.Line.ForeColor, before(k).OldRGB, before(k).OldTheme, before(k).OldTint
            End If
        Next k
        For j = 1 To series.Points.Count
            Set point = series.Points(j)
            Touch point.Format.Fill.ForeColor, address & "|Q:" & CStr(i) & ":" & CStr(j) & ":F"
            Touch point.Format.Line.ForeColor, address & "|Q:" & CStr(i) & ":" & CStr(j) & ":L"
        Next j
    Next i
    newTheme = savedTheme
    On Error GoTo 0
End Sub

Private Function ChartThemeIndex(ByVal value As Long) As Long
    Select Case value
        Case msoThemeColorBackground1: ChartThemeIndex = msoThemeColorLight1
        Case msoThemeColorText1: ChartThemeIndex = msoThemeColorDark1
        Case msoThemeColorBackground2: ChartThemeIndex = msoThemeColorLight2
        Case msoThemeColorText2: ChartThemeIndex = msoThemeColorDark2
        Case Else: ChartThemeIndex = value
    End Select
End Function

Private Function WillReplace(ByVal color As Object) As Boolean
    Dim i As Long
    On Error GoTo Done
    If color.ObjectThemeColor <> msoNotThemeColor Then Exit Function
    For i = 1 To pickedN
        If color.RGB = pickedColors(i) Then WillReplace = True: Exit Function
    Next i
Done:
End Function

Private Sub SnapshotPoint(ByVal color As Object, ByVal address As String, ByRef items() As ColorUndoItem, ByRef n As Long)
    Dim value As Long, theme As Long, tint As Single
    On Error GoTo Done
    value = color.RGB: theme = color.ObjectThemeColor: tint = color.TintAndShade
    n = n + 1
    items(n).Address = address: items(n).OldRGB = value
    items(n).OldTheme = theme: items(n).OldTint = tint
    RecordUndo address, value, theme, tint
Done:
End Sub

Private Sub RestoreColor(ByVal color As Object, ByVal value As Long, ByVal theme As Long, ByVal tint As Single)
    If theme = msoNotThemeColor Then
        color.RGB = value
    Else
        color.ObjectThemeColor = theme
    End If
    color.TintAndShade = tint
End Sub

Private Sub Touch(ByVal color As ColorFormat, ByVal address As String)
    Dim value As Long, i As Long
    On Error GoTo SkipColor
    If color.ObjectThemeColor <> msoNotThemeColor Then Exit Sub
    value = color.RGB
    If Not doReplace Then
        If ColorDistance(value, baseColor) <= radius Then AddFound value
    Else
        For i = 1 To pickedN
            If value = pickedColors(i) Then
                If useTheme Then
                    color.ObjectThemeColor = newTheme
                Else
                    color.RGB = newColor
                End If
                RecordUndo address, value
                changedN = changedN + 1
                Exit For
            End If
        Next i
    End If
SkipColor:
End Sub

Private Sub RecordUndo(ByVal address As String, ByVal oldValue As Long, Optional ByVal oldTheme As Long = 0, Optional ByVal oldTint As Single = 0)
    undoN = undoN + 1
    If undoN > UBound(undoItems) Then ReDim Preserve undoItems(1 To undoN * 2)
    undoItems(undoN).Address = address
    undoItems(undoN).OldRGB = oldValue
    undoItems(undoN).OldTheme = oldTheme
    undoItems(undoN).OldTint = oldTint
End Sub

Private Function ResolveColor(ByVal deck As Presentation, ByVal address As String) As ColorFormat
    Dim parts() As String, path() As String, prop() As String
    Dim sh As Shape, shapes As Shapes, i As Long, series As Object
    On Error GoTo NotFound
    parts = Split(address, "|")
    Select Case parts(0)
        Case "S"
            If parts(3) = "BG" Then
                Set ResolveColor = deck.Slides.FindBySlideID(CLng(parts(1))).Background.Fill.ForeColor
                Exit Function
            End If
            Set shapes = deck.Slides.FindBySlideID(CLng(parts(1))).Shapes
        Case "M": Set shapes = deck.Designs(CLng(parts(1))).SlideMaster.Shapes
        Case "L"
            path = Split(parts(1), ":")
            Set shapes = deck.Designs(CLng(path(0))).SlideMaster.CustomLayouts(CLng(path(1))).Shapes
    End Select
    path = Split(parts(2), ".")
    Set sh = shapes(CLng(path(0)))
    For i = 1 To UBound(path)
        Set sh = sh.GroupItems(CLng(path(i)))
    Next i
    Select Case parts(3)
        Case "F": Set ResolveColor = sh.Fill.ForeColor
        Case "L": Set ResolveColor = sh.Line.ForeColor
        Case Else
            prop = Split(parts(3), ":")
            Select Case prop(0)
                Case "X": Set ResolveColor = sh.TextFrame.TextRange.Runs(CLng(prop(1))).Font.Color
                Case "B": Set ResolveColor = sh.Table.Cell(CLng(prop(1)), CLng(prop(2))).Borders(CLng(prop(3))).ForeColor
                Case "C"
                    Select Case prop(3)
                        Case "F": Set ResolveColor = sh.Table.Cell(CLng(prop(1)), CLng(prop(2))).Shape.Fill.ForeColor
                        Case "L": Set ResolveColor = sh.Table.Cell(CLng(prop(1)), CLng(prop(2))).Shape.Line.ForeColor
                        Case "X": Set ResolveColor = sh.Table.Cell(CLng(prop(1)), CLng(prop(2))).Shape.TextFrame.TextRange.Runs(CLng(prop(4))).Font.Color
                    End Select
                Case "Q"
                    Set series = sh.Chart.SeriesCollection(CLng(prop(1)))
                    If CLng(prop(2)) > 0 Then Set series = series.Points(CLng(prop(2)))
                    Select Case prop(3)
                        Case "F": Set ResolveColor = series.Format.Fill.ForeColor
                        Case "L": Set ResolveColor = series.Format.Line.ForeColor
                    End Select
            End Select
    End Select
    Exit Function
NotFound:
    Set ResolveColor = Nothing
End Function

Private Sub AddFound(ByVal value As Long)
    Dim i As Long
    For i = 1 To foundN
        If foundColors(i) = value Then
            foundCounts(i) = foundCounts(i) + 1
            Exit Sub
        End If
    Next i
    foundN = foundN + 1
    If foundN > UBound(foundColors) Then
        ReDim Preserve foundColors(1 To foundN * 2)
        ReDim Preserve foundCounts(1 To foundN * 2)
    End If
    foundColors(foundN) = value
    foundCounts(foundN) = 1
End Sub

Private Sub SortFound()
    Dim i As Long, j As Long, temp As Long
    For i = 1 To foundN - 1
        For j = i + 1 To foundN
            If ColorDistance(foundColors(j), baseColor) < ColorDistance(foundColors(i), baseColor) Then
                temp = foundColors(i): foundColors(i) = foundColors(j): foundColors(j) = temp
                temp = foundCounts(i): foundCounts(i) = foundCounts(j): foundCounts(j) = temp
            End If
        Next j
    Next i
End Sub

Private Sub ParsePicked(ByVal value As String)
    Dim items() As String, item As Variant, c As Long, i As Long, exists As Boolean
    items = Split(value, ",")
    pickedN = 0
    ReDim pickedColors(1 To UBound(items) + 1)
    For Each item In items
        c = HexToLong(CStr(item))
        exists = False
        For i = 1 To foundN
            If foundColors(i) = c Then exists = True
        Next i
        If Not exists Then Err.Raise 5, , "颜色 " & LongToHex(c) & " 不在扫描结果中。"
        exists = False
        For i = 1 To pickedN
            If pickedColors(i) = c Then exists = True
        Next i
        If Not exists Then
            pickedN = pickedN + 1
            pickedColors(pickedN) = c
        End If
    Next item
    If pickedN = 0 Then Err.Raise 5, , "请至少选择一种旧颜色。"
End Sub

Private Sub ParseTarget(ByVal value As String)
    Select Case UCase$(Trim$(value))
        Case "B1": useTheme = True: newTheme = msoThemeColorBackground1
        Case "T1": useTheme = True: newTheme = msoThemeColorText1
        Case "B2": useTheme = True: newTheme = msoThemeColorBackground2
        Case "T2": useTheme = True: newTheme = msoThemeColorText2
        Case "A1": useTheme = True: newTheme = msoThemeColorAccent1
        Case "A2": useTheme = True: newTheme = msoThemeColorAccent2
        Case "A3": useTheme = True: newTheme = msoThemeColorAccent3
        Case "A4": useTheme = True: newTheme = msoThemeColorAccent4
        Case "A5": useTheme = True: newTheme = msoThemeColorAccent5
        Case "A6": useTheme = True: newTheme = msoThemeColorAccent6
        Case Else: useTheme = False: newColor = HexToLong(value)
    End Select
End Sub

Public Function HexToLong(ByVal value As String) As Long
    Dim h As String
    h = UCase$(Replace$(Trim$(value), "#", ""))
    If Len(h) <> 6 Or h Like "*[!0-9A-F]*" Then Err.Raise 5, , "请输入六位十六进制色号。"
    HexToLong = RGB(CLng("&H" & Left$(h, 2)), CLng("&H" & Mid$(h, 3, 2)), CLng("&H" & Right$(h, 2)))
End Function

Public Function LongToHex(ByVal value As Long) As String
    LongToHex = Right$("0" & Hex$(value Mod 256), 2) & Right$("0" & Hex$((value \ 256) Mod 256), 2) & _
        Right$("0" & Hex$((value \ 65536) Mod 256), 2)
End Function

Private Function ColorDistance(ByVal a As Long, ByVal b As Long) As Double
    Dim dr As Double, dg As Double, db As Double
    dr = (a Mod 256) - (b Mod 256)
    dg = ((a \ 256) Mod 256) - ((b \ 256) Mod 256)
    db = ((a \ 65536) Mod 256) - ((b \ 65536) Mod 256)
    ColorDistance = Sqr(dr * dr + dg * dg + db * db)
End Function
