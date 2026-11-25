Attribute VB_Name = "modLists"
Option Explicit
' Dung style + ListTemplate da cap cho mot ho so (idempotent: goi nhieu lan khong tao them).
' Moi buoc ghi vao mStg de loi tra ve cho biet dang o dau.

Private mStg As String

Private Function NumStyleOf(ByVal n As String) As Long
    Select Case LCase$(n)
        Case "arabic": NumStyleOf = 0
        Case "upperroman": NumStyleOf = 1
        Case "lowerroman": NumStyleOf = 2
        Case "upperletter": NumStyleOf = 3
        Case "lowerletter": NumStyleOf = 4
        Case "bullet": NumStyleOf = 23
        Case Else: NumStyleOf = 0
    End Select
End Function

Private Function AlignOf(ByVal a As String) As Long
    Select Case LCase$(a)
        Case "center": AlignOf = 1
        Case "right": AlignOf = 2
        Case "justify": AlignOf = 3
        Case Else: AlignOf = 0
    End Select
End Function

Private Function FindListTemplate(ByVal doc As Document, ByVal nm As String) As ListTemplate
    Dim lt As ListTemplate
    For Each lt In doc.ListTemplates
        If lt.Name = nm Then
            Set FindListTemplate = lt
            Exit Function
        End If
    Next lt
End Function

Private Sub EnsureStyle(ByVal doc As Document, ByVal prof As Object, ByVal level As Long)
    Dim nm As String, st As Style, ol As Long
    nm = modProfile.LevelStyle(prof, level)
    If Len(nm) = 0 Then Exit Sub
    mStg = "style.get " & nm
    On Error Resume Next
    Set st = doc.Styles(nm)
    On Error GoTo 0
    If st Is Nothing Then
        mStg = "style.add " & nm
        Set st = doc.Styles.Add(nm, 1)
        ' Styles.Add chep dinh dang cua doan dang chon (vd ngat trang, in dam): dua ve Normal
        mStg = "style.reset " & nm
        st.BaseStyle = doc.Styles(-1)
        st.Font = doc.Styles(-1).Font
        st.ParagraphFormat = doc.Styles(-1).ParagraphFormat
    End If
    mStg = "style.pagebreak " & nm
    st.ParagraphFormat.PageBreakBefore = (Val(modProfile.PGet(prof, modProfile.LevelKey(level, "pagebreak"), "0")) <> 0)
    st.ParagraphFormat.KeepWithNext = False
    mStg = "style.next " & nm
    st.NextParagraphStyle = nm
    mStg = "style.bold " & nm
    st.Font.Bold = (Val(modProfile.PGet(prof, modProfile.LevelKey(level, "bold"), "0")) <> 0)
    mStg = "style.align " & nm
    st.ParagraphFormat.Alignment = AlignOf(modProfile.PGet(prof, modProfile.LevelKey(level, "align"), "left"))
    mStg = "style.outline " & nm
    ol = CLng(Val(modProfile.PGet(prof, modProfile.LevelKey(level, "outline"), "0")))
    If ol >= 1 And ol <= 9 Then st.ParagraphFormat.OutlineLevel = ol Else st.ParagraphFormat.OutlineLevel = 10
End Sub

Private Sub ConfigLevel(ByVal lt As ListTemplate, ByVal prof As Object, ByVal level As Long)
    Dim tr As String, rh As String, np As Single, tp As Single
    np = CentimetersToPoints(Val(modProfile.PGet(prof, modProfile.LevelKey(level, "numpos"), "0")))
    tp = CentimetersToPoints(Val(modProfile.PGet(prof, modProfile.LevelKey(level, "textpos"), "0")))
    tr = LCase$(modProfile.PGet(prof, modProfile.LevelKey(level, "trailing"), "tab"))
    With lt.ListLevels(level)
        mStg = "lvl" & CStr(level) & ".format"
        .NumberFormat = modProfile.PGet(prof, modProfile.LevelKey(level, "format"), "%" & CStr(level) & ".")
        mStg = "lvl" & CStr(level) & ".numstyle"
        .NumberStyle = NumStyleOf(modProfile.PGet(prof, modProfile.LevelKey(level, "numstyle"), "arabic"))
        mStg = "lvl" & CStr(level) & ".format2"
        .NumberFormat = modProfile.PGet(prof, modProfile.LevelKey(level, "format"), "%" & CStr(level) & ".")
        mStg = "lvl" & CStr(level) & ".trailing"
        If tr = "space" Then
            .TrailingCharacter = 1
        ElseIf tr = "none" Then
            .TrailingCharacter = 2
        Else
            .TrailingCharacter = 0
        End If
        mStg = "lvl" & CStr(level) & ".align"
        .Alignment = 0
        mStg = "lvl" & CStr(level) & ".pos"
        .NumberPosition = np
        .TextPosition = tp
        .TabPosition = tp
        mStg = "lvl" & CStr(level) & ".startat"
        .StartAt = 1
        mStg = "lvl" & CStr(level) & ".reset"
        rh = modProfile.PGet(prof, modProfile.LevelKey(level, "resetonhigher"), "")
        If Len(rh) = 0 Then
            .ResetOnHigher = level - 1
        Else
            .ResetOnHigher = CLng(Val(rh))
        End If
        mStg = "lvl" & CStr(level) & ".link"
        If Len(modProfile.LevelStyle(prof, level)) > 0 Then .LinkedStyle = modProfile.LevelStyle(prof, level)
    End With
End Sub

Public Sub EnsureProfile(ByVal doc As Document, ByVal prof As Object)
    Dim n As Long, i As Long, lt As ListTemplate, nm As String, en As Long, ed As String
    On Error GoTo EH
    n = modProfile.LevelCount(prof)
    nm = "STC_" & modProfile.PGet(prof, "profile.name")
    mStg = "lt.find"
    Set lt = FindListTemplate(doc, nm)
    If lt Is Nothing Then
        mStg = "lt.add"
        Set lt = doc.ListTemplates.Add(OutlineNumbered:=True)
        mStg = "lt.name"
        lt.Name = nm
    End If
    For i = 1 To n
        EnsureStyle doc, prof, i
        ConfigLevel lt, prof, i
    Next i
    Exit Sub
EH:
    en = Err.Number: ed = Err.Description
    Err.Raise en, "modLists", "[" & mStg & "] " & ed
End Sub

Public Function ActiveProfileName(ByVal doc As Document) As String
    Dim v As Variable
    For Each v In doc.Variables
        If v.Name = "STC_Profile" Then ActiveProfileName = v.Value
    Next v
End Function

Public Sub SetActiveProfile(ByVal doc As Document, ByVal profileName As String)
    Dim cur As String
    cur = ActiveProfileName(doc)
    If cur = profileName Then Exit Sub
    If Len(cur) = 0 Then
        doc.Variables.Add "STC_Profile", profileName
    Else
        doc.Variables("STC_Profile").Value = profileName
    End If
End Sub
