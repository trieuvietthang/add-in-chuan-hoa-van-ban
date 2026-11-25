Attribute VB_Name = "modProfile"
Option Explicit
' Ho so (Profile): doc INI -> Dictionary khoa "section.key" (chu thuong). Khong hard-code thong so the thuc.

Public Sub ParseIni(ByVal txt As String, ByVal d As Object)
    Dim lines() As String, i As Long, ln As String, sec As String, p As Long
    If Len(txt) > 0 Then If AscW(Left$(txt, 1)) = &HFEFF Then txt = Mid$(txt, 2)
    txt = Replace(txt, vbCrLf, vbLf)
    txt = Replace(txt, vbCr, vbLf)
    lines = Split(txt, vbLf)
    For i = 0 To UBound(lines)
        ln = Trim$(lines(i))
        If Len(ln) > 0 Then
            If Left$(ln, 1) <> ";" And Left$(ln, 1) <> "#" Then
                If Left$(ln, 1) = "[" Then
                    sec = LCase$(Mid$(ln, 2, Len(ln) - 2))
                Else
                    p = InStr(ln, "=")
                    If p > 1 Then d(sec & "." & LCase$(Trim$(Left$(ln, p - 1)))) = Trim$(Mid$(ln, p + 1))
                End If
            End If
        End If
    Next i
End Sub

Public Function LoadProfile(ByVal profileName As String) As Object
    Dim d As Object, f As String
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    f = modConfig.ConfigDir & "\profiles\" & profileName & ".ini"
    If Len(Dir$(f)) = 0 Then Err.Raise vbObjectError + 1002, "modProfile", modStrings.T("err.noconfig")
    ParseIni modConfig.ReadUtf8(f), d
    Set LoadProfile = d
End Function

Public Function PGet(ByVal prof As Object, ByVal key As String, Optional ByVal def As String = "") As String
    If prof.Exists(LCase$(key)) Then PGet = prof(LCase$(key)) Else PGet = def
End Function

Public Function LevelCount(ByVal prof As Object) As Long
    LevelCount = CLng(Val(PGet(prof, "profile.levels", "0")))
End Function

Public Function LevelKey(ByVal level As Long, ByVal field As String) As String
    LevelKey = "level" & CStr(level) & "." & field
End Function

Public Function LevelStyle(ByVal prof As Object, ByVal level As Long) As String
    LevelStyle = PGet(prof, LevelKey(level, "style"))
End Function

' Tra ve cap (1..n) cua mot ten style trong ho so; 0 neu khong thuoc ho so
Public Function LevelOfStyle(ByVal prof As Object, ByVal styleName As String) As Long
    Dim i As Long
    For i = 1 To LevelCount(prof)
        If StrComp(LevelStyle(prof, i), styleName, vbTextCompare) = 0 Then
            LevelOfStyle = i
            Exit Function
        End If
    Next i
End Function
