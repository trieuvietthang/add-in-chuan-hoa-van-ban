Attribute VB_Name = "modRibbon"
Option Explicit
' Callback ribbon (customUI14.xml). Selection chi duoc DOC vi tri con tro; moi ghi di qua Range (luat 3).

Private mRibbon As IRibbonUI
Private mProfile As String
Private mStd As String
Private mIssues() As String
Private mIssueCount As Long
Private mIssueIdx As Long
Private mSpell() As String
Private mSpellCount As Long
Private mSpellIdx As Long

Public Function Ping() As String
    Ping = "ok"
End Function

Public Sub OnRibbonLoad(ByVal ribbon As IRibbonUI)
    Set mRibbon = ribbon
End Sub

Private Function CurrentProfile() As String
    If Len(mProfile) = 0 Then
        On Error Resume Next
        mProfile = modLists.ActiveProfileName(ActiveDocument)
        On Error GoTo 0
        If Len(mProfile) = 0 Then mProfile = "HopDong"
    End If
    CurrentProfile = mProfile
End Function

Private Sub Report(ByVal where As String)
    MsgBox modStrings.T("err.generic") & vbCrLf & "(" & where & ": " & CStr(Err.Number) & " " & Err.Description & ")", vbExclamation
End Sub

' Vung lam viec: vung chon neu co, nguoc lai doan chua con tro
Private Function WorkRange() As Range
    Set WorkRange = Selection.Range
End Function

' ---- Ho so ----
Public Sub GetProfileId(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = CurrentProfile()
End Sub

Public Sub OnProfileSelect(ByVal control As IRibbonControl, ByVal id As String, ByVal index As Long)
    mProfile = id
    If Not mRibbon Is Nothing Then mRibbon.Invalidate
End Sub

' ---- Chuan the thuc (ND30 / ISO9001) + chuan hoa ----
Public Sub GetStdId(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    If Len(mStd) = 0 Then mStd = "ND30"
    returnedVal = mStd
End Sub

Public Sub OnStdSelect(ByVal control As IRibbonControl, ByVal id As String, ByVal index As Long)
    mStd = id
End Sub

Public Sub OnNormalize(ByVal control As IRibbonControl)
    On Error GoTo EH
    Dim pv As String, lines() As String, shown As String, i As Long, kv() As String, n As Long
    If Len(mStd) = 0 Then mStd = "ND30"
    pv = modNormalize.PreviewStandard(ActiveDocument, CurrentProfile(), mStd)
    If Len(pv) = 0 Then
        MsgBox modStrings.T("norm.nothing"), vbInformation
        Exit Sub
    End If
    lines = Split(pv, vbLf)
    For i = 0 To UBound(lines)
        If Len(lines(i)) > 0 Then
            kv = Split(lines(i), "|")
            shown = shown & modStrings.T(kv(0)) & ": " & kv(1) & vbCrLf
        End If
    Next i
    If MsgBox(modStrings.T("norm.confirm") & vbCrLf & vbCrLf & shown, vbOKCancel + vbQuestion) <> vbOK Then Exit Sub
    n = modNormalize.ApplyStandard(ActiveDocument, CurrentProfile(), mStd)
    MsgBox modStrings.T("norm.done") & " (" & CStr(n) & ")", vbInformation
    Exit Sub
EH:
    Report "OnNormalize"
End Sub

' ---- Dan chieu ----
Private Function Msg1(ByVal key As String, ByVal a As String) As String
    Msg1 = Replace(modStrings.T(key), "{0}", a)
End Function

Private Sub GoIssue(ByVal i As Long)
    Dim kv() As String
    kv = Split(mIssues(i), "|")
    ActiveDocument.Range(CLng(kv(1)), CLng(kv(2))).Select
    MsgBox "(" & CStr(i + 1) & "/" & CStr(mIssueCount) & ") " & Msg1("ref." & kv(0), kv(3)), vbInformation
End Sub

Public Sub OnCheckRefs(ByVal control As IRibbonControl)
    On Error GoTo EH
    Dim lines() As String, i As Long, kv() As String, nErr As Long, nTyped As Long
    lines = Split(modRefs.CheckRefs(ActiveDocument, CurrentProfile()), vbLf)
    ReDim mIssues(0 To UBound(lines) + 1)
    mIssueCount = 0: mIssueIdx = 0
    For i = 0 To UBound(lines)
        If Len(lines(i)) > 0 Then
            kv = Split(lines(i), "|")
            If kv(0) = "R4" Then
                nTyped = nTyped + 1
            Else
                mIssues(mIssueCount) = lines(i)
                mIssueCount = mIssueCount + 1
            End If
        End If
    Next i
    nErr = mIssueCount
    MsgBox Replace(Replace(Replace(modStrings.T("ref.summary"), "{0}", CStr(nErr)), "{1}", CStr(nTyped)), "{2}", vbCrLf), vbInformation
    If nErr > 0 Then GoIssue 0
    Exit Sub
EH:
    Report "OnCheckRefs"
End Sub

Public Sub OnNextRef(ByVal control As IRibbonControl)
    On Error GoTo EH
    If mIssueCount = 0 Then
        MsgBox modStrings.T("ref.nolist"), vbInformation
        Exit Sub
    End If
    mIssueIdx = (mIssueIdx + 1) Mod mIssueCount
    GoIssue mIssueIdx
    Exit Sub
EH:
    Report "OnNextRef"
End Sub

Public Sub OnConvertRefs(ByVal control As IRibbonControl)
    On Error GoTo EH
    Dim n As Long
    If MsgBox(modStrings.T("ref.convertask"), vbOKCancel + vbQuestion) <> vbOK Then Exit Sub
    n = modRefs.ConvertRefs(ActiveDocument, CurrentProfile())
    MsgBox Msg1("ref.converted", CStr(n)), vbInformation
    Exit Sub
EH:
    Report "OnConvertRefs"
End Sub

Public Sub OnUpdateRefs(ByVal control As IRibbonControl)
    On Error GoTo EH
    Dim n As Long
    n = modRefs.UpdateRefs(ActiveDocument)
    MsgBox Msg1("ref.updated", CStr(n)), vbInformation
    Exit Sub
EH:
    Report "OnUpdateRefs"
End Sub

Public Sub OnInsertRef(ByVal control As IRibbonControl)
    On Error GoTo EH
    Dim s As String, r As String
    s = InputBox(modStrings.T("ref.askno"), "Chen dan chieu")
    If Len(s) = 0 Then Exit Sub
    If Not IsNumeric(s) Then Exit Sub
    r = modRefs.InsertRef(ActiveDocument, WorkRange().Start, CLng(s), CurrentProfile())
    If Len(r) = 0 Then MsgBox Msg1("ref.R1", s), vbExclamation
    Exit Sub
EH:
    Report "OnInsertRef"
End Sub

' ---- Chinh ta ----
Private Sub GoSpell(ByVal i As Long)
    Dim kv() As String, det As String
    kv = Split(mSpell(i), "|")
    ActiveDocument.Range(CLng(kv(1)), CLng(kv(2))).Select
    If UBound(kv) >= 3 Then det = modSpell.Plain(kv(3))
    MsgBox "(" & CStr(i + 1) & "/" & CStr(mSpellCount) & ") " & Replace(modStrings.T("spell.msg." & kv(0)), "{0}", det), vbInformation
End Sub

Public Sub OnCheckSpell(ByVal control As IRibbonControl)
    On Error GoTo EH
    Dim lines() As String, i As Long, kv() As String, rules As Variant, r As Variant, cnt As Object, shown As String
    Set cnt = CreateObject("Scripting.Dictionary")
    lines = Split(modSpell.CheckSpell(ActiveDocument, CurrentProfile()), vbLf)
    ReDim mSpell(0 To UBound(lines) + 1)
    mSpellCount = 0: mSpellIdx = 0
    For i = 0 To UBound(lines)
        If Len(lines(i)) > 0 Then
            mSpell(mSpellCount) = lines(i)
            mSpellCount = mSpellCount + 1
            kv = Split(lines(i), "|")
            cnt(kv(0)) = 1 + CLng(IIf(cnt.Exists(kv(0)), cnt(kv(0)), 0))
        End If
    Next i
    If mSpellCount = 0 Then
        MsgBox modStrings.T("spell.none") & vbCrLf & vbCrLf & modStrings.T("spell.note"), vbInformation
        Exit Sub
    End If
    For Each r In Array("S1", "S2", "S3", "S4", "S5", "S6", "W1", "W2", "W3", "W4", "W5")
        If cnt.Exists(CStr(r)) Then shown = shown & modStrings.T("spell.lbl." & CStr(r)) & ": " & CStr(cnt(CStr(r))) & vbCrLf
    Next r
    MsgBox modStrings.T("spell.title") & vbCrLf & vbCrLf & shown & vbCrLf & modStrings.T("spell.note"), vbInformation
    GoSpell 0
    Exit Sub
EH:
    Report "OnCheckSpell"
End Sub

Public Sub OnNextSpell(ByVal control As IRibbonControl)
    On Error GoTo EH
    If mSpellCount = 0 Then
        MsgBox modStrings.T("spell.nolist"), vbInformation
        Exit Sub
    End If
    mSpellIdx = (mSpellIdx + 1) Mod mSpellCount
    GoSpell mSpellIdx
    Exit Sub
EH:
    Report "OnNextSpell"
End Sub

' ---- Sua ho so ----
Public Sub OnEditProfile(ByVal control As IRibbonControl)
    On Error GoTo EH
    frmProfile.Show
    modSpell.ResetConfig
    Exit Sub
EH:
    Report "OnEditProfile"
End Sub

' ---- Muc luc ----
Public Sub OnInsertToc(ByVal control As IRibbonControl)
    On Error GoTo EH
    Dim doc As Document, idx As Long, ans As VbMsgBoxResult
    Set doc = ActiveDocument
    idx = doc.Range(0, WorkRange().Paragraphs(1).Range.End).Paragraphs.Count
    If modToc.CountSources(doc, CurrentProfile()) = 0 Then
        If Len(modConvert.PreviewTyped(doc.Content, CurrentProfile())) > 0 Then
            ans = MsgBox(modStrings.T("toc.convertask"), vbYesNoCancel + vbQuestion)
            If ans = vbCancel Then Exit Sub
            If ans = vbYes Then modConvert.ConvertTyped doc.Content, CurrentProfile()
        End If
    End If
    modToc.InsertToc doc, idx, CurrentProfile()
    If modToc.CountSources(doc, CurrentProfile()) = 0 Then MsgBox modStrings.T("toc.nosource"), vbInformation
    Exit Sub
EH:
    Report "OnInsertToc"
End Sub

Public Sub OnUpdateToc(ByVal control As IRibbonControl)
    On Error GoTo EH
    If modToc.UpdateToc(ActiveDocument) = 0 Then MsgBox modStrings.T("toc.none"), vbInformation
    Exit Sub
EH:
    Report "OnUpdateToc"
End Sub

' ---- Nut cap (tag = so cap) ----
Public Sub GetLevelLabel(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    On Error GoTo EH
    Dim prof As Object
    Set prof = modProfile.LoadProfile(CurrentProfile())
    returnedVal = modProfile.PGet(prof, modProfile.LevelKey(CLng(control.Tag), "label"), "Cap " & control.Tag)
    Exit Sub
EH:
    returnedVal = "Cap " & control.Tag
End Sub

Public Sub GetLevelVisible(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    On Error GoTo EH
    Dim prof As Object
    Set prof = modProfile.LoadProfile(CurrentProfile())
    returnedVal = (Len(modProfile.LevelStyle(prof, CLng(control.Tag))) > 0)
    Exit Sub
EH:
    returnedVal = False
End Sub

Public Sub OnLevel(ByVal control As IRibbonControl)
    On Error GoTo EH
    modApply.ApplyLevelToRange WorkRange(), CLng(control.Tag), CurrentProfile()
    Exit Sub
EH:
    Report "OnLevel"
End Sub

' ---- Tang / giam cap (tag = +1 / -1) ----
Public Sub OnChangeLevel(ByVal control As IRibbonControl)
    On Error GoTo EH
    modApply.ChangeLevelRange WorkRange(), CLng(control.Tag), CurrentProfile()
    Exit Sub
EH:
    Report "OnChangeLevel"
End Sub

' ---- Chuyen so go tay (co xem truoc) ----
Public Sub OnConvertTyped(ByVal control As IRibbonControl)
    On Error GoTo EH
    Dim rng As Range, pv As String, n As Long, lines() As String, shown As String, i As Long, kv() As String
    Set rng = WorkRange()
    If rng.Start = rng.End Then Set rng = ActiveDocument.Content
    pv = modConvert.PreviewTyped(rng, CurrentProfile())
    If Len(pv) = 0 Then
        MsgBox modStrings.T("convert.nothing"), vbInformation
        Exit Sub
    End If
    lines = Split(pv, vbLf)
    For i = 0 To UBound(lines)
        If Len(lines(i)) > 0 And i < 15 Then
            kv = Split(lines(i), "|")
            shown = shown & Replace(Replace(modStrings.T("convert.line"), "{0}", kv(0)), "{1}", kv(1)) & vbCrLf
        End If
    Next i
    If MsgBox(modStrings.T("convert.confirm") & vbCrLf & vbCrLf & shown, vbOKCancel + vbQuestion) <> vbOK Then Exit Sub
    n = modConvert.ConvertTyped(rng, CurrentProfile())
    MsgBox modStrings.T("convert.done") & " (" & CStr(n) & ")", vbInformation
    Exit Sub
EH:
    Report "OnConvertTyped"
End Sub
