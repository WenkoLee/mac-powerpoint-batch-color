' SPDX-License-Identifier: MIT
' Mac PowerPoint Batch Color Replacer — 0.1.0-beta.2
Option Explicit

#If Mac Then
Private Declare PtrSafe Function CurrentKeyModifiers Lib "/System/Library/Frameworks/Carbon.framework/Carbon" Alias "GetCurrentKeyModifiers" () As Long
#End If
#If Mac Then
Private Declare PtrSafe Function PBCreate Lib "/System/Library/Frameworks/ApplicationServices.framework/ApplicationServices" Alias "PasteboardCreate" (ByVal name As LongPtr, ByRef board As LongPtr) As Long
Private Declare PtrSafe Function PBSync Lib "/System/Library/Frameworks/ApplicationServices.framework/ApplicationServices" Alias "PasteboardSynchronize" (ByVal board As LongPtr) As Long
Private Declare PtrSafe Function PBCount Lib "/System/Library/Frameworks/ApplicationServices.framework/ApplicationServices" Alias "PasteboardGetItemCount" (ByVal board As LongPtr, ByRef count As Long) As Long
Private Declare PtrSafe Function PBItem Lib "/System/Library/Frameworks/ApplicationServices.framework/ApplicationServices" Alias "PasteboardGetItemIdentifier" (ByVal board As LongPtr, ByVal index As LongPtr, ByRef item As LongPtr) As Long
Private Declare PtrSafe Function PBData Lib "/System/Library/Frameworks/ApplicationServices.framework/ApplicationServices" Alias "PasteboardCopyItemFlavorData" (ByVal board As LongPtr, ByVal item As LongPtr, ByVal flavor As LongPtr, ByRef data As LongPtr) As Long
Private Declare PtrSafe Function CFAscii Lib "/System/Library/Frameworks/CoreFoundation.framework/CoreFoundation" Alias "CFStringCreateWithCString" (ByVal allocator As LongPtr, ByVal value As String, ByVal encoding As Long) As LongPtr
Private Declare PtrSafe Function CFDecode Lib "/System/Library/Frameworks/CoreFoundation.framework/CoreFoundation" Alias "CFStringCreateFromExternalRepresentation" (ByVal allocator As LongPtr, ByVal data As LongPtr, ByVal encoding As Long) As LongPtr
Private Declare PtrSafe Function CFLength Lib "/System/Library/Frameworks/CoreFoundation.framework/CoreFoundation" Alias "CFStringGetLength" (ByVal value As LongPtr) As LongPtr
Private Declare PtrSafe Function CFChar Lib "/System/Library/Frameworks/CoreFoundation.framework/CoreFoundation" Alias "CFStringGetCharacterAtIndex" (ByVal value As LongPtr, ByVal index As LongPtr) As Integer
Private Declare PtrSafe Function CFDataLength Lib "/System/Library/Frameworks/CoreFoundation.framework/CoreFoundation" Alias "CFDataGetLength" (ByVal data As LongPtr) As LongPtr
Private Declare PtrSafe Sub CFDrop Lib "/System/Library/Frameworks/CoreFoundation.framework/CoreFoundation" Alias "CFRelease" (ByVal value As LongPtr)
#End If

Private mTextUndo(1 To 4) As Collection
Private mTextRedo(1 To 4) As Collection
Private mTextLast(1 To 4) As String
Private mHistoryReady As Boolean
Private mTextRestore As Boolean

Private mDeckName As String
Private mReady As Boolean
Private mScanPages As String
Private mScanOldSpec As String
Private mPaletteMode As String
Private WithEvents mApp As Application
Private mPickKind As String
Private mBusy As Boolean
Private mOld As Long
Private mTarget As String
Private mCount As Long
Private mStart As Long
Private mPick() As Boolean
Private mScope As String
Private mSelectedSpec As String
Private mSelectedCount As Long

Public Sub OpenForDeck(ByVal deck As Presentation)
    If mReady And mDeckName = deck.FullName Then
        CaptureSelectedSlides
        ClearCandidates
        SyncLabels
        If Not Me.Visible Then Me.Show vbModeless
        Exit Sub
    End If
    mDeckName = deck.FullName
    mSelectedSpec = vbNullString
    mSelectedCount = 0
    CaptureSelectedSlides
    BuildPanel
    mReady = True
    mPickKind = "F"
    Set mApp = Application
    mApp_WindowSelectionChange Application.ActiveWindow.Selection
    Me.Show vbModeless
End Sub

Public Sub resetForm()
    BuildPanel
End Sub

Public Sub SetInitColor(ByVal value As Long)
    mOld = value
    mBusy = True
    HexBox.Value = ColorReplaceInPlace.LongToHex(value)
    mTextLast(1) = ColorReplaceInPlace.LongToHex(value)
    ColorLabel.BackColor = value
    mBusy = False
End Sub

Public Function GetSelectedColor() As Long
    GetSelectedColor = mOld
End Function

Private Function GetDeck() As Presentation
    Dim deck As Presentation
    If Application.Presentations.Count = 0 Then Err.Raise 5, , "演示文稿已关闭，请关闭面板后重新打开。"
    Set deck = Application.ActivePresentation
    If deck.FullName <> mDeckName Then Err.Raise 5, , "请切回打开此面板的文件，或在当前文件重新打开面板。"
    Set GetDeck = deck
End Function

Private Sub CaptureSelectedSlides()
    Dim spec As String, n As Long
    On Error GoTo Finished
    spec = ColorReplaceInPlace.SelectedPagesSpec()
    n = UBound(Split(spec, ",")) + 1
    If spec <> mSelectedSpec Then
        mSelectedSpec = spec
        mSelectedCount = n
        If mReady And mScope = "SELECTED" Then ClearCandidates
    End If
Finished:
    If mReady Then SetCaption ScopeSelected, IIf(mScope = "SELECTED", "● ", "○ ") & "已选页 (" & CStr(mSelectedCount) & ")"
End Sub

Private Sub UserForm_Activate()
    If Not mReady Then Exit Sub
    CaptureSelectedSlides
    SyncLabels
End Sub

Private Function ScopeSpec() As String
    Select Case mScope
        Case "ALL": ScopeSpec = vbNullString
        Case "SELECTED"
            CaptureSelectedSlides
            If mSelectedCount = 0 Then Err.Raise 5, , "请切回普通视图中的一张幻灯片。"
            ScopeSpec = mSelectedSpec
        Case "CUSTOM"
            If Len(Trim$(RangeBox.Value)) = 0 Then Err.Raise 5, , "请输入页码，例如 2,4-6。"
            ScopeSpec = Trim$(RangeBox.Value)
    End Select
End Function

Private Sub ChooseScope(ByVal scope As String)
    If scope = "SELECTED" And mSelectedCount = 0 Then
        MsgBox "请切回普通视图中的一张幻灯片。", vbInformation, "Mac 批量改色（测试版）"
        Exit Sub
    End If
    mScope = scope
    SetCaption ScopeAll, IIf(scope = "ALL", "● 所有页", "○ 所有页")
    SetCaption ScopeSelected, IIf(scope = "SELECTED", "● 已选页 (" & CStr(mSelectedCount) & ")", "○ 已选页 (" & CStr(mSelectedCount) & ")")
    SetCaption ScopeCustom, IIf(scope = "CUSTOM", "● 指定页码", "○ 指定页码")
    ScopeAll.ForeColor = IIf(scope = "ALL", RGB(0, 102, 204), RGB(31, 41, 55))
    ScopeSelected.ForeColor = IIf(scope = "SELECTED", RGB(0, 102, 204), RGB(31, 41, 55))
    ScopeCustom.ForeColor = IIf(scope = "CUSTOM", RGB(0, 102, 204), RGB(31, 41, 55))
    ClearCandidates
    SyncLabels
End Sub

Private Sub ScopeAll_Click()
    ChooseScope "ALL"
End Sub

Private Sub ScopeSelected_Click()
    CaptureSelectedSlides
    ChooseScope "SELECTED"
End Sub

Private Sub ScopeCustom_Click()
    ChooseScope "CUSTOM"
End Sub

Private Sub RangeBox_Change()
    On Error Resume Next
    If mBusy Then Exit Sub
    ChooseScope "CUSTOM"
End Sub

Private Sub BuildPanel()
    Dim ctl As Object, i As Long, item As Object
    mBusy = True
    mHistoryReady = False
    Me.Caption = ""
    Me.Width = 680
    Me.Height = 418
    Me.BackColor = RGB(246, 247, 250)
    Me.Font.Name = "Microsoft YaHei"
    For Each ctl In Me.Controls
        On Error Resume Next
        ctl.Visible = False
        On Error GoTo 0
    Next ctl

    ' Keep informational labels transparent above the white cards.
    For Each ctl In Me.Controls
        Select Case ctl.Name
            Case "ThemeColor11", "ThemeColor12", _
                "RadiusHint", "CandidatesHint", "TargetTitle", "TargetHint", "ThemeHint"
                On Error Resume Next
                ctl.BackStyle = 0
                On Error GoTo 0
        End Select
    Next ctl

    Place LeftCard, 16, 14, 314, 276
    Place RightCard, 342, 14, 322, 276
    LeftCard.Caption = "": RightCard.Caption = ""
    LeftCard.BackColor = vbWhite: RightCard.BackColor = vbWhite
    LeftCard.BorderStyle = 1: RightCard.BorderStyle = 1
    LeftCard.BorderColor = RGB(229, 232, 237)
    RightCard.BorderColor = RGB(229, 232, 237)


    Place ThemeColor11, 28, 30, 285, 22
    SetCaption ThemeColor11, ChrW$(26087) & ChrW$(39068) & ChrW$(33394)
    ThemeColor11.Font.Size = 14
    ThemeColor11.Font.Bold = True
    ThemeColor11.ForeColor = RGB(31, 41, 55)
    Place ThemeColor12, 28, 57, 285, 18
    SetCaption ThemeColor12, "HEX / 主题编号，或从已选对象取色"
    ThemeColor12.Font.Size = 10.5
    ThemeColor12.ForeColor = RGB(105, 113, 124)

    Place HexBox, 28, 82, 130, 27
    Place ColorLabel, 166, 82, 56, 27
    SetCaption ColorLabel, ChrW$(39044) & ChrW$(35272)
    ColorLabel.TextAlign = 2
    For i = 1 To 3
        Select Case i
            Case 1: Set item = PickFill: SetCaption item, ChrW$(22635) & ChrW$(20805)
            Case 2: Set item = PickLine: SetCaption item, ChrW$(32447) & ChrW$(26465)
            Case 3: Set item = PickText: SetCaption item, ChrW$(25991) & ChrW$(23383)
        End Select
        Place item, 226 + (i - 1) * 32, 82, 28, 27
        item.BackColor = RGB(232, 241, 255)
        item.ForeColor = RGB(0, 102, 204)
        item.TextAlign = 2
        item.BorderStyle = 1
        item.BorderColor = RGB(214, 228, 247)
    Next i

    Place RedBox, 28, 121, 48, 26
    RedBox.Value = "40"
    Place RadiusHint, 82, 124, 100, 20
    SetCaption RadiusHint, ChrW$(30456) & ChrW$(36817) & ChrW$(33539) & ChrW$(22260)
    RadiusHint.ForeColor = RGB(105, 113, 124)
    Place CommandButtonEsc, 184, 119, 130, 29
    SetCaption CommandButtonEsc, ChrW$(25195) & ChrW$(25551) & ChrW$(30456) & ChrW$(36817) & ChrW$(33394)

    Place CandidatesHint, 28, 156, 286, 17
    SetCaption CandidatesHint, "直接替换精确旧色 · 扫描相近色可选"
    CandidatesHint.Font.Size = 10.5
    CandidatesHint.ForeColor = RGB(105, 113, 124)
    For i = 2 To 12
        Set item = Me.Controls("StandardColor" & CStr(i))
        If i <= 9 Then
            Place item, 28 + ((i - 2) Mod 2) * 150, 182 + ((i - 2) \ 2) * 25, 136, 23
            item.TextAlign = 2
        End If
        item.Visible = False
    Next i
    Place BlueBar, 315, 182, 11, 98
    BlueBar.Visible = False

    Place TargetTitle, 354, 30, 292, 22
    SetCaption TargetTitle, ChrW$(30446) & ChrW$(26631) & ChrW$(39068) & ChrW$(33394)
    TargetTitle.ForeColor = RGB(31, 41, 55)
    TargetTitle.Font.Size = 14
    TargetTitle.Font.Bold = True
    Place TargetHint, 354, 57, 292, 18
    SetCaption TargetHint, ChrW$(36755) & ChrW$(20837) & " HEX" & ChrW$(65292) & ChrW$(25110) & ChrW$(28857) & ChrW$(20987) & ChrW$(19979) & ChrW$(26041) & ChrW$(20027) & ChrW$(39064) & ChrW$(33394)
    TargetHint.Font.Size = 10.5
    TargetHint.ForeColor = RGB(105, 113, 124)
    Place LongBox, 354, 82, 132, 27
    Place StandardColor1, 494, 82, 155, 27
    SetCaption StandardColor1, ChrW$(39044) & ChrW$(35272)
    StandardColor1.TextAlign = 2

    Place ThemeHint, 354, 126, 56, 17
    SetCaption ThemeHint, "主题色"
    Place PaletteOld, 412, 122, 104, 24
    Place PaletteTarget, 522, 122, 126, 24
    mPaletteMode = "TARGET"
    SetPaletteMode "TARGET"
    ThemeHint.Font.Size = 10.5
    ThemeHint.ForeColor = RGB(105, 113, 124)
    For i = 1 To 10
        Set item = Me.Controls("ThemeColor" & CStr(i))
        Place item, 354 + ((i - 1) Mod 5) * 58, 154 + ((i - 1) \ 5) * 52, 50, 42
        SetCaption item, ThemeCode(i)
        item.BackColor = ThemeColorValue(i)
        item.ForeColor = ContrastColor(item.BackColor)
        item.TextAlign = 2
        item.Font.Bold = True
    Next i

    Place ScopeTitle, 28, 298, 70, 23
    SetCaption ScopeTitle, "处理页面"
    ScopeTitle.ForeColor = RGB(31, 41, 55)
    ScopeTitle.Font.Bold = True
    Place ScopeAll, 96, 298, 88, 23
    Place ScopeSelected, 188, 298, 115, 23
    Place ScopeCustom, 309, 298, 102, 23
    Place RangeBox, 414, 297, 135, 25
    RangeBox.Value = ""
    Place RangeHint, 554, 300, 112, 20
    SetCaption RangeHint, "例：2,4-6"
    RangeHint.ForeColor = RGB(105, 113, 124)
    mScope = "SELECTED"
    SetCaption ScopeAll, "○ 所有页"
    SetCaption ScopeSelected, "● 已选页 (" & CStr(mSelectedCount) & ")"
    SetCaption ScopeCustom, "○ 指定页码"
    ScopeAll.ForeColor = RGB(31, 41, 55)
    ScopeSelected.ForeColor = RGB(0, 102, 204)
    ScopeCustom.ForeColor = RGB(31, 41, 55)

    Place OKButton, 424, 342, 104, 30
    SetCaption OKButton, "替换选定范围"
    Place CancelButton, 540, 342, 112, 30
    SetCaption CancelButton, ChrW$(21462) & ChrW$(28040)
    OKButton.ForeColor = RGB(31, 41, 55)
    CommandButtonEsc.ForeColor = RGB(31, 41, 55)
    OKButton.Font.Bold = True
    CommandButtonEsc.Default = False
    OKButton.Default = True
    CancelButton.Cancel = True

    mOld = ColorReplaceInPlace.HexToLong("204F92")
    HexBox.Value = "204F92"
    ColorLabel.BackColor = mOld
    ColorLabel.ForeColor = vbWhite
    mTarget = "T2"
    LongBox.Value = "T2"
    StandardColor1.BackColor = ThemeColorValue(4)
    StandardColor1.ForeColor = ContrastColor(StandardColor1.BackColor)
    mCount = 0
    mStart = 1
    ReDim mPick(1 To 1)
    ScopeTitle.TextAlign = 2
    ScopeAll.TextAlign = 2: ScopeSelected.TextAlign = 2: ScopeCustom.TextAlign = 2
    RangeHint.TextAlign = 2
    ResetTextHistory
    SyncLabels
    mBusy = False
End Sub

Private Sub Place(ByVal item As Object, ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single)
    Dim frame As Object, lineHeight As Single
    item.Left = x: item.Top = y: item.Width = w: item.Height = h
    item.Visible = True
    On Error Resume Next
    item.Font.Name = "Microsoft YaHei"
    item.Font.Size = 11.5
    If TypeName(item) = "TextBox" Then
        Set frame = Me.Controls(item.Name & "Frame")
        frame.Left = x: frame.Top = y: frame.Width = w: frame.Height = h
        frame.Caption = "": frame.BackColor = vbWhite
        frame.BorderStyle = 1: frame.BorderColor = RGB(207, 213, 222)
        frame.Visible = True
        lineHeight = 21
        item.Left = x + 5: item.Width = w - 10
        item.Top = y + (h - lineHeight) / 2: item.Height = lineHeight
        item.BorderStyle = 0: item.SpecialEffect = 0
        item.BackColor = vbWhite: item.MultiLine = False
        item.Locked = False: item.Enabled = True
        item.HideSelection = False: item.TabStop = True
        item.ZOrder 0
    End If
    On Error GoTo 0
End Sub

Private Sub SetCaption(ByVal item As Object, ByVal value As String)
    item.Tag = value
    item.Caption = ""
    SyncLabel item
End Sub

Private Sub SyncLabel(ByVal item As Object)
    Dim text As Object, lineHeight As Single
    On Error Resume Next
    Set text = Me.Controls(item.Name & "Text")
    If text Is Nothing Then Exit Sub
    lineHeight = item.Font.Size * 1.4 + 2
    If lineHeight > item.Height Then lineHeight = item.Height
    text.Left = item.Left: text.Width = item.Width
    text.Top = item.Top + (item.Height - lineHeight) / 2
    text.Height = lineHeight
    text.Caption = item.Tag: text.BackStyle = 0
    text.Font.Name = "Microsoft YaHei": text.Font.Size = item.Font.Size
    text.Font.Bold = item.Font.Bold
    text.ForeColor = item.ForeColor
    text.TextAlign = 2
    If TypeName(item) <> "CommandButton" Then text.TextAlign = item.TextAlign
    text.Visible = item.Visible
    text.ZOrder 0
    On Error GoTo 0
End Sub

Private Sub SyncLabels()
    Dim name As Variant
    For Each name In Split("PaletteOld,PaletteTarget,PickFill,PickLine,PickText,RadiusHint,CandidatesHint,TargetTitle,TargetHint,ThemeHint,ScopeTitle,ScopeAll,ScopeSelected,ScopeCustom,RangeHint,ColorLabel,ThemeColor11,ThemeColor12,StandardColor1,CommandButtonEsc,OKButton,CancelButton,ThemeColor1,ThemeColor2,ThemeColor3,ThemeColor4,ThemeColor5,ThemeColor6,ThemeColor7,ThemeColor8,ThemeColor9,ThemeColor10,StandardColor2,StandardColor3,StandardColor4,StandardColor5,StandardColor6,StandardColor7,StandardColor8,StandardColor9,StandardColor10,StandardColor11,StandardColor12", ",")
        SyncLabel Me.Controls(CStr(name))
    Next name
End Sub

Private Function ThemeCode(ByVal index As Long) As String
    Select Case index
        Case 1: ThemeCode = "B1"
        Case 2: ThemeCode = "T1"
        Case 3: ThemeCode = "B2"
        Case 4: ThemeCode = "T2"
        Case Else: ThemeCode = "A" & CStr(index - 4)
    End Select
End Function

Private Function ThemeSchemeIndex(ByVal index As Long) As Long
    Select Case index
        Case 1: ThemeSchemeIndex = msoThemeLight1
        Case 2: ThemeSchemeIndex = msoThemeDark1
        Case 3: ThemeSchemeIndex = msoThemeLight2
        Case 4: ThemeSchemeIndex = msoThemeDark2
        Case Else: ThemeSchemeIndex = msoThemeAccent1 + index - 5
    End Select
End Function

Private Function ThemeColorValue(ByVal index As Long) As Long
    Dim slide As Slide, deck As Presentation
    On Error GoTo Fallback
    Set deck = GetDeck
    Set slide = Application.ActiveWindow.View.Slide
    ThemeColorValue = slide.Design.SlideMaster.Theme.ThemeColorScheme(ThemeSchemeIndex(index)).RGB
    Exit Function
Fallback:
    Err.Clear
    On Error GoTo Unavailable
    ThemeColorValue = deck.Designs(1).SlideMaster.Theme.ThemeColorScheme(ThemeSchemeIndex(index)).RGB
    Exit Function
Unavailable:
    ThemeColorValue = RGB(180, 180, 180)
End Function

Private Function ContrastColor(ByVal value As Long) As Long
    Dim lum As Double
    lum = (value Mod 256) * 0.299 + ((value \ 256) Mod 256) * 0.587 + ((value \ 65536) Mod 256) * 0.114
    If lum < 150 Then
        ContrastColor = vbWhite
    Else
        ContrastColor = vbBlack
    End If
End Function

Private Sub RefreshOldColor()
    Dim value As Long
    On Error GoTo ChangeFinished
    TrackText HexBox
    If mBusy Then Exit Sub
    On Error GoTo Invalid
    value = ColorReplaceInPlace.ColorSpecToRGB(GetDeck(), CStr(HexBox.Value), ColorReplaceInPlace.CurrentSlideIndex())
    If value <> mOld Then ClearCandidates
    mOld = value
    ColorLabel.BackColor = value
    ColorLabel.ForeColor = ContrastColor(value)
    SyncLabels
    Exit Sub
Invalid:
ChangeFinished:
End Sub

Private Sub RedBox_Change()
    On Error Resume Next
    If mBusy Then Exit Sub
    ClearCandidates
End Sub

Private Sub ClearCandidates()
    Dim i As Long
    mCount = 0
    mScanPages = vbNullString
    mScanOldSpec = vbNullString
    SetCaption CandidatesHint, "直接替换精确旧色 · 扫描相近色可选"
    For i = 2 To 12
        Me.Controls("StandardColor" & CStr(i)).Visible = False
    Next i
    BlueBar.Visible = False
    SyncLabels
End Sub

Private Sub RefreshTargetColor()
    Dim value As Long, spec As String, i As Long
    On Error GoTo ChangeFinished
    TrackText LongBox
    If mBusy Then Exit Sub
    spec = UCase$(Trim$(LongBox.Value))
    On Error GoTo Invalid
    For i = 1 To 10
        If spec = ThemeCode(i) Then
            mTarget = spec
            value = ThemeColorValue(i)
            GoTo Preview
        End If
    Next i
    value = ColorReplaceInPlace.HexToLong(spec)
    mTarget = spec
Preview:
    StandardColor1.BackColor = value
    StandardColor1.ForeColor = ContrastColor(value)
    SyncLabels
    Exit Sub
Invalid:
ChangeFinished:
End Sub

Private Sub SetPaletteMode(ByVal mode As String)
    Dim item As Object, controlName As Variant
    mPaletteMode = mode
    SetCaption PaletteOld, "选择旧色"
    SetCaption PaletteTarget, "选择目标色"
    For Each controlName In Array("PaletteOld", "PaletteTarget")
        Set item = Me.Controls(CStr(controlName))
        item.TextAlign = 2
        item.BorderStyle = 1
        item.BorderColor = RGB(214, 228, 247)
        item.BackColor = RGB(246, 247, 250)
        item.ForeColor = RGB(105, 113, 124)
    Next controlName
    If mode = "OLD" Then Set item = PaletteOld Else Set item = PaletteTarget
    item.BackColor = RGB(232, 241, 255)
    item.ForeColor = RGB(0, 102, 204)
    SyncLabels
End Sub

Private Sub PaletteOld_Click()
    SetPaletteMode "OLD"
End Sub

Private Sub PaletteTarget_Click()
    SetPaletteMode "TARGET"
End Sub

Private Sub PaletteOldText_Click()
    PaletteOld_Click
End Sub

Private Sub PaletteTargetText_Click()
    PaletteTarget_Click
End Sub

Private Sub HexBox_Enter()
    If mReady Then SetPaletteMode "OLD"
End Sub

Private Sub LongBox_Enter()
    If mReady Then SetPaletteMode "TARGET"
End Sub

Private Sub ChooseTheme(ByVal index As Long)
    If mPaletteMode = "OLD" Then
        ClearCandidates
        mBusy = True
        mOld = ThemeColorValue(index)
        HexBox.Value = ThemeCode(index)
        mTextLast(1) = ThemeCode(index)
        ColorLabel.BackColor = mOld
        ColorLabel.ForeColor = ContrastColor(mOld)
        mBusy = False
        RefreshInput HexBox
    Else
        mBusy = True
        mTarget = ThemeCode(index)
        LongBox.Value = mTarget
        mTextLast(2) = mTarget
        StandardColor1.BackColor = ThemeColorValue(index)
        StandardColor1.ForeColor = ContrastColor(StandardColor1.BackColor)
        mBusy = False
        RefreshInput LongBox
    End If
    SyncLabels
End Sub

Private Sub ReadSelected(ByVal kind As String)
    Dim deck As Presentation
    On Error GoTo Failed
    mPickKind = kind
    Set deck = GetDeck
    Set deck = Nothing
    mOld = ColorReplaceInPlace.GetSelectedObjectColor(kind)
    ClearCandidates
    mBusy = True
    HexBox.Value = ColorReplaceInPlace.LongToHex(mOld)
    mTextLast(1) = ColorReplaceInPlace.LongToHex(mOld)
    ColorLabel.BackColor = mOld
    ColorLabel.ForeColor = ContrastColor(mOld)
    SyncLabels
    mBusy = False
    Exit Sub
Failed:
    mBusy = False
    MsgBox Err.Description, vbInformation, ChrW$(39029) & ChrW$(38754) & ChrW$(21462) & ChrW$(33394)
End Sub

Private Sub ScanNearby()
    Dim distance As Double, i As Long, deck As Presentation
    On Error GoTo Failed
    Set deck = GetDeck
    mOld = ColorReplaceInPlace.ColorSpecToRGB(deck, CStr(HexBox.Value), ColorReplaceInPlace.CurrentSlideIndex())
    If Not IsNumeric(RedBox.Value) Then Err.Raise 5, , ChrW$(30456) & ChrW$(36817) & ChrW$(33539) & ChrW$(22260) & ChrW$(24517) & ChrW$(39035) & ChrW$(26159) & ChrW$(25968) & ChrW$(23383) & ChrW$(12290)
    distance = CDbl(RedBox.Value)
    ClearCandidates
    Set deck = GetDeck
    mScanPages = ScopeSpec()
    mScanOldSpec = UCase$(Trim$(CStr(HexBox.Value)))
    ColorReplaceInPlace.PrepareScan deck, mOld, distance, mScanPages
    Set deck = Nothing
    mCount = ColorReplaceInPlace.CandidateCount()
    If mCount = 0 Then
        MsgBox "所选范围没有找到接近的颜色。也可以不扫描，直接精确替换旧色。", vbInformation, "Mac 批量改色（测试版）"
        Exit Sub
    End If
    ReDim mPick(1 To mCount)
    For i = 1 To mCount
        mPick(i) = True
    Next i
    mStart = 1
    mBusy = True
    If mCount > 8 Then
        BlueBar.Min = 1
        BlueBar.Max = mCount - 7
        BlueBar.SmallChange = 1
        BlueBar.LargeChange = 8
        BlueBar.Value = 1
        BlueBar.Visible = True
    Else
        BlueBar.Visible = False
    End If
    mBusy = False
    ShowCandidates
    Exit Sub
Failed:
    MsgBox ChrW$(25195) & ChrW$(25551) & ChrW$(22833) & ChrW$(36133) & ChrW$(65306) & Err.Description, vbExclamation, ChrW$(25209) & ChrW$(37327) & ChrW$(25913) & ChrW$(33394)
End Sub

Private Sub ShowCandidates()
    Dim slot As Long, index As Long, value As Long, item As Object, marker As String
    SetCaption CandidatesHint, "匹配颜色 · 点击可取消勾选"
    For slot = 1 To 8
        Set item = Me.Controls("StandardColor" & CStr(slot + 1))
        index = mStart + slot - 1
        If index <= mCount Then
            value = ColorReplaceInPlace.CandidateColor(index)
            If mPick(index) Then marker = "[x]" Else marker = "[ ]"
            SetCaption item, marker & " #" & ColorReplaceInPlace.LongToHex(value) & "  (" & _
                CStr(ColorReplaceInPlace.CandidateUses(index)) & ")"
            item.BackColor = value
            item.ForeColor = ContrastColor(value)
            item.Visible = True
        Else
            item.Visible = False
        End If
    Next slot
    SyncLabels
End Sub

Private Sub ToggleCandidate(ByVal slot As Long)
    Dim index As Long
    index = mStart + slot - 1
    If index < 1 Or index > mCount Then Exit Sub
    mPick(index) = Not mPick(index)
    ShowCandidates
End Sub

Private Sub BlueBar_Change()
    If mBusy Then Exit Sub
    mStart = BlueBar.Value
    ShowCandidates
End Sub

Private Sub OKButton_Click()
    Dim selected As String, i As Long, n As Long, pages As String, deck As Presentation, prompt As String
    On Error GoTo Failed
    ValidateColorInput
    Set deck = GetDeck
    pages = ScopeSpec()
    If mCount > 0 Then
        If pages <> mScanPages Or UCase$(Trim$(CStr(HexBox.Value))) <> mScanOldSpec Then ClearCandidates
    End If
    If mCount = 0 Then
        selected = CStr(HexBox.Value)
        prompt = "将所选范围的旧色 " & selected & " 精确替换为 " & CStr(LongBox.Value) & "？"
    Else
        For i = 1 To mCount
            If mPick(i) Then
                If Len(selected) > 0 Then selected = selected & ","
                selected = selected & ColorReplaceInPlace.LongToHex(ColorReplaceInPlace.CandidateColor(i))
                n = n + 1
            End If
        Next i
        If n = 0 Then Err.Raise 5, , "请至少勾选一种旧颜色。"
        prompt = "将所选范围内勾选的 " & CStr(n) & " 种颜色替换为 " & CStr(LongBox.Value) & "？"
    End If
    If MsgBox(prompt, vbYesNo + vbQuestion, "确认替换") <> vbYes Then Exit Sub
    If mCount = 0 Then
        n = ColorReplaceInPlace.ApplyExact(deck, selected, CStr(LongBox.Value), pages)
    Else
        n = ColorReplaceInPlace.ApplySelection(deck, selected, CStr(LongBox.Value), pages)
    End If
    If n = 0 Then
        MsgBox "所选范围没有匹配的颜色，未作更改。", vbInformation, "Mac 批量改色（测试版）"
        Exit Sub
    End If
    mDeckName = vbNullString
    mReady = False
    Set mApp = Nothing
    Unload Me
    MsgBox "已替换 " & CStr(n) & " 处颜色。需要回退时，请点插件的“撤销上次改色”。", vbInformation, "Mac 批量改色（测试版）"
    Exit Sub
Failed:
    MsgBox "替换未完成：" & Err.Description, vbExclamation, "Mac 批量改色（测试版）"
End Sub

Private Sub CancelButton_Click()
    mDeckName = vbNullString
    mReady = False
    Set mApp = Nothing
    Unload Me
End Sub

Private Sub CommandButtonEsc_Click()
    ScanNearby
End Sub

Private Sub ThemeColor11_Click()
End Sub

Private Sub ThemeColor12_Click()
End Sub

Private Sub PickFill_Click()
    ReadSelected "F"
End Sub

Private Sub PickLine_Click()
    ReadSelected "L"
End Sub

Private Sub PickText_Click()
    ReadSelected "T"
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    mDeckName = vbNullString
    mReady = False
    Set mApp = Nothing
End Sub

Private Sub ThemeColor1_Click()
    ChooseTheme 1
End Sub

Private Sub ThemeColor2_Click()
    ChooseTheme 2
End Sub

Private Sub ThemeColor3_Click()
    ChooseTheme 3
End Sub

Private Sub ThemeColor4_Click()
    ChooseTheme 4
End Sub

Private Sub ThemeColor5_Click()
    ChooseTheme 5
End Sub

Private Sub ThemeColor6_Click()
    ChooseTheme 6
End Sub

Private Sub ThemeColor7_Click()
    ChooseTheme 7
End Sub

Private Sub ThemeColor8_Click()
    ChooseTheme 8
End Sub

Private Sub ThemeColor9_Click()
    ChooseTheme 9
End Sub

Private Sub ThemeColor10_Click()
    ChooseTheme 10
End Sub

Private Sub StandardColor2_Click()
    ToggleCandidate 1
End Sub

Private Sub StandardColor3_Click()
    ToggleCandidate 2
End Sub

Private Sub StandardColor4_Click()
    ToggleCandidate 3
End Sub

Private Sub StandardColor5_Click()
    ToggleCandidate 4
End Sub

Private Sub StandardColor6_Click()
    ToggleCandidate 5
End Sub

Private Sub StandardColor7_Click()
    ToggleCandidate 6
End Sub

Private Sub StandardColor8_Click()
    ToggleCandidate 7
End Sub

Private Sub StandardColor9_Click()
    ToggleCandidate 8
End Sub

Private Sub StandardColor10_Click()
    ToggleCandidate 9
End Sub

Private Sub StandardColor11_Click()
    ToggleCandidate 10
End Sub

Private Sub StandardColor12_Click()
    ToggleCandidate 11
End Sub


Private Sub mApp_WindowSelectionChange(ByVal Sel As Selection)
    Dim deck As Presentation, value As Long
    On Error GoTo Finished
    If Not mReady Or mBusy Then Exit Sub
    Set deck = GetDeck
    Set deck = Nothing
    CaptureSelectedSlides
    If Sel.Type = ppSelectionShapes Or Sel.Type = ppSelectionText Then
        value = ColorReplaceInPlace.GetSelectedObjectColor(mPickKind)
        If UCase$(Trim$(CStr(HexBox.Value))) <> ColorReplaceInPlace.LongToHex(value) Then ClearCandidates
        mBusy = True
        mOld = value
        HexBox.Value = ColorReplaceInPlace.LongToHex(value)
    mTextLast(1) = ColorReplaceInPlace.LongToHex(value)
        ColorLabel.BackColor = value
        ColorLabel.ForeColor = ContrastColor(value)
    End If
    SyncLabels
Finished:
    mBusy = False
End Sub

Private Sub mApp_PresentationClose(ByVal Pres As Presentation)
    On Error Resume Next
    If Not mReady Then Exit Sub
    If Pres.FullName <> mDeckName Then Exit Sub
    mReady = False
    mDeckName = vbNullString
    Set mApp = Nothing
    Unload Me
End Sub

Private Sub PickFillText_Click()
    PickFill_Click
End Sub

Private Sub PickLineText_Click()
    PickLine_Click
End Sub

Private Sub PickTextText_Click()
    PickText_Click
End Sub

Private Sub ScopeAllText_Click()
    ScopeAll_Click
End Sub

Private Sub ScopeSelectedText_Click()
    ScopeSelected_Click
End Sub

Private Sub ScopeCustomText_Click()
    ScopeCustom_Click
End Sub

Private Sub ThemeColor11Text_Click()
    ThemeColor11_Click
End Sub

Private Sub ThemeColor12Text_Click()
    ThemeColor12_Click
End Sub

Private Sub ThemeColor1Text_Click()
    ThemeColor1_Click
End Sub

Private Sub ThemeColor2Text_Click()
    ThemeColor2_Click
End Sub

Private Sub ThemeColor3Text_Click()
    ThemeColor3_Click
End Sub

Private Sub ThemeColor4Text_Click()
    ThemeColor4_Click
End Sub

Private Sub ThemeColor5Text_Click()
    ThemeColor5_Click
End Sub

Private Sub ThemeColor6Text_Click()
    ThemeColor6_Click
End Sub

Private Sub ThemeColor7Text_Click()
    ThemeColor7_Click
End Sub

Private Sub ThemeColor8Text_Click()
    ThemeColor8_Click
End Sub

Private Sub ThemeColor9Text_Click()
    ThemeColor9_Click
End Sub

Private Sub ThemeColor10Text_Click()
    ThemeColor10_Click
End Sub

Private Sub StandardColor2Text_Click()
    StandardColor2_Click
End Sub

Private Sub StandardColor3Text_Click()
    StandardColor3_Click
End Sub

Private Sub StandardColor4Text_Click()
    StandardColor4_Click
End Sub

Private Sub StandardColor5Text_Click()
    StandardColor5_Click
End Sub

Private Sub StandardColor6Text_Click()
    StandardColor6_Click
End Sub

Private Sub StandardColor7Text_Click()
    StandardColor7_Click
End Sub

Private Sub StandardColor8Text_Click()
    StandardColor8_Click
End Sub

Private Sub StandardColor9Text_Click()
    StandardColor9_Click
End Sub

Private Sub StandardColor10Text_Click()
    StandardColor10_Click
End Sub

Private Sub StandardColor11Text_Click()
    StandardColor11_Click
End Sub

Private Sub StandardColor12Text_Click()
    StandardColor12_Click
End Sub

Private Function TextIndex(ByVal box As Object) As Long
    Select Case box.Name
        Case "HexBox": TextIndex = 1
        Case "LongBox": TextIndex = 2
        Case "RedBox": TextIndex = 3
        Case "RangeBox": TextIndex = 4
    End Select
End Function

Private Sub ResetTextHistory()
    Dim i As Long, box As Object
    For i = 1 To 4
        Set mTextUndo(i) = New Collection
        Set mTextRedo(i) = New Collection
    Next i
    For Each box In Me.Controls
        If TextIndex(box) > 0 Then mTextLast(TextIndex(box)) = box.Value
    Next box
    mHistoryReady = True
End Sub

Private Sub TrackText(ByVal box As Object)
    Dim i As Long
    On Error GoTo Finished
    If Not mHistoryReady Or mTextRestore Then Exit Sub
    i = TextIndex(box)
    If box.Value = mTextLast(i) Then Exit Sub
    If mBusy Then
        Set mTextUndo(i) = New Collection
    Else
        mTextUndo(i).Add mTextLast(i)
        If mTextUndo(i).Count > 100 Then mTextUndo(i).Remove 1
    End If
    Set mTextRedo(i) = New Collection
    mTextLast(i) = box.Value
Finished:
End Sub

Private Sub UndoText(ByVal box As Object, ByVal redo As Boolean)
    Dim i As Long, source As Collection, destination As Collection, value As String
    i = TextIndex(box)
    If redo Then
        Set source = mTextRedo(i): Set destination = mTextUndo(i)
    Else
        Set source = mTextUndo(i): Set destination = mTextRedo(i)
    End If
    If source.Count = 0 Then Exit Sub
    destination.Add box.Value
    value = CStr(source(source.Count))
    source.Remove source.Count
    mTextRestore = True
    box.Value = value
    box.SelStart = Len(value): box.SelLength = 0
    mTextRestore = False
    mTextLast(i) = value
    RefreshInput box
End Sub

Private Function EditModifier(ByVal Shift As Integer) As Boolean
    EditModifier = ((Shift And 2) <> 0)
#If Mac Then
    On Error Resume Next
    EditModifier = EditModifier Or ((CurrentKeyModifiers() And 256) <> 0)
#End If
End Function

Private Function CleanInput(ByVal value As String, ByVal color As Boolean) As String
    value = Replace$(Replace$(value, vbCr, ""), vbLf, "")
    If color Then value = UCase$(Replace$(Replace$(Replace$(value, "#", ""), " ", ""), vbTab, ""))
    CleanInput = value
End Function



Private Sub EditKey(ByVal box As Object, ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim key As Long
    If Not EditModifier(Shift) Then Exit Sub
    key = CLng(KeyCode)
    If key >= 97 And key <= 122 Then key = key - 32
    Select Case key
        Case vbKeyA, vbKeyC, vbKeyX, vbKeyV, vbKeyZ, vbKeyY
            KeyCode = 0
            TextAction box, key, Shift
    End Select
End Sub

Private Sub TextAction(ByVal box As Object, ByVal key As Long, ByVal Shift As Integer)
    On Error GoTo Failed
    Select Case key
        Case vbKeyA
            box.SelStart = 0: box.SelLength = Len(mTextLast(TextIndex(box)))
        Case vbKeyC
            box.Copy
        Case vbKeyX
            box.Cut
        Case vbKeyV
            PasteText box
        Case vbKeyZ
            UndoText box, EditShift(Shift)
        Case vbKeyY
            UndoText box, True
    End Select
    Exit Sub
Failed:
    mTextRestore = False
    If key = vbKeyV Then
        PasteNotice box, Err.Description
    Else
        MsgBox "文本操作未完成，请重试：" & Err.Description, vbInformation, "色号编辑"
    End If
End Sub

Private Sub HexBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    EditKey HexBox, KeyCode, Shift
End Sub

Private Sub LongBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    EditKey LongBox, KeyCode, Shift
End Sub

Private Sub RedBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    EditKey RedBox, KeyCode, Shift
End Sub

Private Sub RangeBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    EditKey RangeBox, KeyCode, Shift
End Sub

Private Sub HexBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    TextCharacter HexBox, KeyAscii
End Sub

Private Sub LongBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    TextCharacter LongBox, KeyAscii
End Sub

Private Sub RedBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    TextCharacter RedBox, KeyAscii
End Sub

Private Sub RangeBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    TextCharacter RangeBox, KeyAscii
End Sub

Private Sub TextCharacter(ByVal box As Object, ByVal KeyAscii As MSForms.ReturnInteger)
    Dim key As Long
    Select Case KeyAscii
        Case 1: key = vbKeyA
        Case 3: key = vbKeyC
        Case 24: key = vbKeyX
        Case 22: key = vbKeyV
        Case 26: key = vbKeyZ
        Case Else: Exit Sub
    End Select
    KeyAscii = 0
    TextAction box, key, 2
End Sub

Private Sub CommandButtonEscText_Click()
    CommandButtonEsc_Click
End Sub

Private Sub OKButtonText_Click()
    OKButton_Click
End Sub

Private Sub CancelButtonText_Click()
    CancelButton_Click
End Sub

Private Sub HexBox_Change()
    On Error Resume Next
    If Not mBusy Then ClearCandidates
End Sub

Private Sub LongBox_Change()
    ' Preview is refreshed after the keystroke or edit operation completes.
End Sub

Private Sub RefreshInput(ByVal box As Object)
    Dim spec As String, valid As Boolean, hint As Object
    On Error GoTo Finished
    If mBusy Or Not mHistoryReady Then Exit Sub
    TrackText box
    If box.Name <> "HexBox" And box.Name <> "LongBox" Then Exit Sub
    spec = UCase$(Replace$(Trim$(CStr(box.Value)), "#", ""))
    valid = (Len(spec) = 6 And Not spec Like "*[!0-9A-F]*")
    If ColorReplaceInPlace.ThemeIndexFromCode(spec) > 0 Then valid = True
    If box.Name = "LongBox" Then
        Set hint = TargetHint
    Else
        Set hint = ThemeColor12
    End If
    If valid Then
        If box.Name = "LongBox" Then
            SetCaption hint, "输入 HEX，或点击下方主题色"
            RefreshTargetColor
        Else
            SetCaption hint, "HEX / 主题编号，或从已选对象取色"
            RefreshOldColor
        End If
        hint.ForeColor = RGB(105, 113, 124)
        Me.Controls(box.Name & "Frame").BorderColor = RGB(207, 213, 222)
    Else
        SetCaption hint, "请输入 HEX 或 B1/T1/B2/T2/A1-A6"
        hint.ForeColor = RGB(170, 77, 55)
        Me.Controls(box.Name & "Frame").BorderColor = RGB(205, 130, 110)
    End If
    SyncLabels
Finished:
End Sub

Private Sub HexBox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    RefreshInput HexBox
End Sub

Private Sub HexBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    RefreshInput HexBox
End Sub

Private Sub LongBox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    RefreshInput LongBox
End Sub

Private Sub LongBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    RefreshInput LongBox
End Sub

Private Sub RedBox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    RefreshInput RedBox
End Sub

Private Sub RedBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    RefreshInput RedBox
End Sub

Private Sub RangeBox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    RefreshInput RangeBox
End Sub

Private Sub RangeBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    RefreshInput RangeBox
End Sub

Private Sub ValidateColorInput()
    Dim value As Long, spec As String
    value = ColorReplaceInPlace.ColorSpecToRGB(GetDeck(), CStr(HexBox.Value), ColorReplaceInPlace.CurrentSlideIndex())
    spec = UCase$(Trim$(CStr(LongBox.Value)))
    Select Case spec
        Case "B1", "T1", "B2", "T2", "A1", "A2", "A3", "A4", "A5", "A6"
        Case Else: value = ColorReplaceInPlace.HexToLong(spec)
    End Select
End Sub

Private Function EditShift(ByVal Shift As Integer) As Boolean
    EditShift = ((Shift And 1) <> 0)
#If Mac Then
    On Error Resume Next
    EditShift = EditShift Or ((CurrentKeyModifiers() And 512) <> 0)
#End If
End Function



Private Function ClipboardPlainText() As String
#If Mac Then
    Dim name As LongPtr, board As LongPtr, item As LongPtr, flavor As LongPtr
    Dim data As LongPtr, value As LongPtr, length As LongPtr, index As LongPtr
    Dim count As Long, status As Long, kind As Variant, encoding As Long
    Dim result As String, issue As String
    On Error GoTo Failed
    name = CFAscii(0, "com.apple.pasteboard.clipboard", &H8000100)
    If name = 0 Then Err.Raise 5
    status = PBCreate(name, board)
    If status <> 0 Or board = 0 Then Err.Raise 5
    status = PBSync(board)
    If PBCount(board, count) <> 0 Or count < 1 Then Err.Raise 5
    If PBItem(board, 1, item) <> 0 Then Err.Raise 5
    For Each kind In Array("public.utf8-plain-text", "public.utf16-external-plain-text", "public.utf16-plain-text")
        flavor = CFAscii(0, CStr(kind), &H8000100)
        If flavor = 0 Then Err.Raise 5
        status = PBData(board, item, flavor, data)
        CFDrop flavor: flavor = 0
        If status = 0 And data <> 0 Then
            If CFDataLength(data) > 32768 Then
                issue = "剪贴板文字过长，请只复制色号。"
                GoTo Cleanup
            End If
            If kind = "public.utf8-plain-text" Then encoding = &H8000100 Else encoding = &H100
            value = CFDecode(0, data, encoding)
            CFDrop data: data = 0
            If value <> 0 Then Exit For
        End If
        If data <> 0 Then CFDrop data: data = 0
    Next kind
    If value = 0 Then Err.Raise 5
    length = CFLength(value)
    If length > 16384 Then Err.Raise 5
    For index = 0 To length - 1
        result = result & ChrW$(CFChar(value, index))
    Next index
    GoTo Cleanup
Failed:
    issue = "剪贴板没有可用的纯文本，请复制色号文字。"
Cleanup:
    On Error Resume Next
    If value <> 0 Then CFDrop value
    If data <> 0 Then CFDrop data
    If flavor <> 0 Then CFDrop flavor
    If board <> 0 Then CFDrop board
    If name <> 0 Then CFDrop name
    On Error GoTo 0
    If Len(issue) > 0 Then Err.Raise 5, , issue
    ClipboardPlainText = result
#Else
    Err.Raise 5, , "此版本的纯文本粘贴适用于 Mac。"
#End If
End Function

Private Function CleanPlainInput(ByVal value As String, ByVal color As Boolean) As String
    Dim i As Long, code As Long, ch As String, result As String
    For i = 1 To Len(value)
        ch = Mid$(value, i, 1): code = AscW(ch)
        If code < 0 Then code = code + 65536
        Select Case code
            Case 0, 9, 10, 13, 32, 160, 5760, 8192 To 8207, 8232, 8233, 8239, 8287, 8288, 12288, 65279
                ch = vbNullString
            Case 65296 To 65305, 65313 To 65318, 65345 To 65350
                ch = ChrW$(code - 65248)
        End Select
        result = result & ch
    Next i
    If color Then
        result = UCase$(result)
        If Left$(result, 1) = "#" Then result = Mid$(result, 2)
    End If
    CleanPlainInput = result
End Function

Private Sub PasteText(ByVal box As Object)
    Dim textId As Long, previous As String, incoming As String, value As String
    Dim cursor As Long, selected As Long, color As Boolean
    textId = TextIndex(box): previous = mTextLast(textId)
    incoming = ClipboardPlainText()
    color = (box.Name = "HexBox" Or box.Name = "LongBox")
    incoming = CleanPlainInput(incoming, color)
    If color Then
        If ColorReplaceInPlace.ThemeIndexFromCode(incoming) = 0 And incoming Like "*[!0-9A-F]*" Then Err.Raise 5, , "剪贴板包含无效字符，请只复制 HEX 色号或主题色编号。"
    End If
    If Len(incoming) = 0 Then Err.Raise 5, , "剪贴板中没有色号文字。"
    cursor = box.SelStart: selected = box.SelLength
    If cursor > Len(previous) Then cursor = Len(previous)
    If selected > Len(previous) - cursor Then selected = Len(previous) - cursor
    value = Left$(previous, cursor) & incoming & Mid$(previous, cursor + selected + 1)
    If Len(value) > 1024 Then Err.Raise 5, , "输入内容过长。"
    mTextRestore = True
    box.Value = value
    box.Font.Name = "Microsoft YaHei"
    box.SelStart = cursor + Len(incoming): box.SelLength = 0
    mTextRestore = False
    If value <> previous Then
        mTextUndo(textId).Add previous
        If mTextUndo(textId).Count > 100 Then mTextUndo(textId).Remove 1
        Set mTextRedo(textId) = New Collection
    End If
    mTextLast(textId) = value
    RefreshInput box
End Sub

Private Sub PasteNotice(ByVal box As Object, ByVal message As String)
    Dim hint As Object
    On Error Resume Next
    If box.Name = "HexBox" Then Set hint = ThemeColor12 Else Set hint = TargetHint
    SetCaption hint, message
    hint.ForeColor = RGB(170, 77, 55)
End Sub
