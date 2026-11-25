Attribute VB_Name = "modConvert"
Option Explicit
' Chuyen so go tay ("Dieu 1.", "1.1", "a)") thanh danh so tu dong theo ho so. Co xem truoc.
' Nhan dien bang bieu thuc chinh quy trong ho so (levelN.detect); chon cap co khop DAI NHAT.

Private Function MatchLevel(ByVal prof As Object, ByVal txt As String, ByRef matchLen As Long) As Long
    Dim i As Long, re As Object, m As Object, best As Long, pat As String
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = False
    For i = 1 To modProfile.LevelCount(prof)
        pat = modProfile.PGet(prof, modProfile.LevelKey(i, "detect"))
        If Len(pat) > 0 And Len(modProfile.LevelStyle(prof, i)) > 0 Then
            re.Pattern = pat
            If re.Test(txt) Then
                Set m = re.Execute(txt)(0)
                If m.Length > best Then
                    best = m.Length
                    MatchLevel = i
                End If
            End If
        End If
    Next i
    matchLen = best
End Function

Private Function ParaText(ByVal p As Paragraph) As String
    ParaText = Replace(Replace(p.Range.Text, vbCr, ""), Chr(7), "")
End Function

' Tra ve nhieu dong "stt|cap|tien to se bo", moi dong ket thuc bang vbLf
Public Function PreviewTyped(ByVal rng As Range, ByVal profileName As String) As String
    Dim prof As Object, p As Paragraph, i As Long, lv As Long, ml As Long, t As String, s As String
    Set prof = modProfile.LoadProfile(profileName)
    For Each p In rng.Paragraphs
        i = i + 1
        t = ParaText(p)
        lv = MatchLevel(prof, t, ml)
        If lv > 0 And Len(t) > ml Then s = s & CStr(i) & "|" & CStr(lv) & "|" & Left$(t, ml) & vbLf
    Next p
    PreviewTyped = s
End Function

' Tra ve so doan da chuyen
Public Function ConvertTyped(ByVal rng As Range, ByVal profileName As String) As Long
    Dim doc As Document, prof As Object, p As Paragraph, lv As Long, ml As Long, t As String
    Dim n As Long, r As Range, st As String, en As Long, ed As String
    Set doc = rng.Document
    modUndo.BeginUndo doc, "Chuyen so go tay"
    On Error GoTo EH
    Set prof = modProfile.LoadProfile(profileName)
    modLists.EnsureProfile doc, prof
    modLists.SetActiveProfile doc, profileName
    For Each p In rng.Paragraphs
        t = ParaText(p)
        lv = MatchLevel(prof, t, ml)
        If lv > 0 And Len(t) > ml Then
            Set r = doc.Range(p.Range.Start, p.Range.Start + ml)
            r.Delete
            p.Style = doc.Styles(modProfile.LevelStyle(prof, lv))
            n = n + 1
        End If
    Next p
    modUndo.EndUndo
    ConvertTyped = n
    Exit Function
EH:
    en = Err.Number: ed = Err.Description
    modUndo.EndUndo
    Err.Raise en, "modConvert", ed
End Function
