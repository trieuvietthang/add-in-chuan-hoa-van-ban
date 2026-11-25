Attribute VB_Name = "modApply"
Option Explicit
' Ap cap danh so (style) cho doan; tang/giam cap. Moi lenh = 1 UndoRecord. Khong dung Selection.

Private Function ResolveProfile(ByVal doc As Document, ByVal profileName As String) As String
    If Len(profileName) = 0 Then profileName = modLists.ActiveProfileName(doc)
    If Len(profileName) = 0 Then Err.Raise vbObjectError + 1001, "modApply", modStrings.T("err.noprofile")
    ResolveProfile = profileName
End Function

Public Sub ApplyLevelToRange(ByVal rng As Range, ByVal level As Long, Optional ByVal profileName As String = "")
    Dim doc As Document, prof As Object, st As String, p As Paragraph, en As Long, ed As String
    Set doc = rng.Document
    profileName = ResolveProfile(doc, profileName)
    modUndo.BeginUndo doc, "Danh so"
    On Error GoTo EH
    Set prof = modProfile.LoadProfile(profileName)
    st = modProfile.LevelStyle(prof, level)
    If Len(st) = 0 Then Err.Raise vbObjectError + 1003, "modApply", modStrings.T("err.nolevel")
    modLists.EnsureProfile doc, prof
    modLists.SetActiveProfile doc, profileName
    For Each p In rng.Paragraphs
        p.Style = doc.Styles(st)
    Next p
    modUndo.EndUndo
    Exit Sub
EH:
    en = Err.Number: ed = Err.Description
    modUndo.EndUndo
    Err.Raise en, "modApply", ed
End Sub

' delta = +1 tang cap (vao sau hon), -1 giam cap. Bo qua doan khong thuoc ho so (tru khi tang cap -> cap 1).
Public Sub ChangeLevelRange(ByVal rng As Range, ByVal delta As Long, Optional ByVal profileName As String = "")
    Dim doc As Document, prof As Object, p As Paragraph, cur As Long, nw As Long, n As Long, en As Long, ed As String
    Set doc = rng.Document
    profileName = ResolveProfile(doc, profileName)
    modUndo.BeginUndo doc, "Doi cap"
    On Error GoTo EH
    Set prof = modProfile.LoadProfile(profileName)
    modLists.EnsureProfile doc, prof
    n = modProfile.LevelCount(prof)
    For Each p In rng.Paragraphs
        cur = modProfile.LevelOfStyle(prof, p.Style.NameLocal)
        If cur = 0 Then
            If delta > 0 Then nw = FirstLevel(prof) Else nw = 0
        Else
            nw = cur
            Do
                nw = nw + Sgn(delta)
            Loop While nw >= 1 And nw <= n And Len(modProfile.LevelStyle(prof, nw)) = 0
            If nw < 1 Or nw > n Then nw = 0
        End If
        If nw > 0 Then p.Style = doc.Styles(modProfile.LevelStyle(prof, nw))
    Next p
    modUndo.EndUndo
    Exit Sub
EH:
    en = Err.Number: ed = Err.Description
    modUndo.EndUndo
    Err.Raise en, "modApply", ed
End Sub

Private Function FirstLevel(ByVal prof As Object) As Long
    Dim i As Long
    For i = 1 To modProfile.LevelCount(prof)
        If Len(modProfile.LevelStyle(prof, i)) > 0 Then
            FirstLevel = i
            Exit Function
        End If
    Next i
End Function
