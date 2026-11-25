Attribute VB_Name = "modIniEdit"
Option Explicit
' Doc/sua/ghi tep INI (UTF-8) giu nguyen chu thich va thu tu; kiem tra gia tri theo config\ui.vi.ini. Dung cho form "Sua ho so".

Private mUi As Object

Public Sub ResetUi()
    Set mUi = Nothing
End Sub

' Chuoi giao dien (co dau) tu config\ui.vi.ini; khoa dang "text.title", "hints.page.top"...
Public Function Ui(ByVal key As String) As String
    Dim f As String
    If mUi Is Nothing Then
        Set mUi = CreateObject("Scripting.Dictionary")
        mUi.CompareMode = 1
        f = modConfig.ConfigDir & "\ui.vi.ini"
        If Len(Dir$(f)) = 0 Then Err.Raise vbObjectError + 1010, "modIniEdit", modStrings.T("err.noconfig")
        modProfile.ParseIni modConfig.ReadUtf8(f), mUi
    End If
    If mUi.Exists(LCase$(key)) Then Ui = mUi(LCase$(key))
End Function

Public Function ReadText(ByVal path As String) As String
    Dim t As String
    t = modConfig.ReadUtf8(path)
    If Len(t) > 0 Then If AscW(Left$(t, 1)) = &HFEFF Then t = Mid$(t, 2)
    ReadText = t
End Function

' Ten cac [muc], moi ten mot dong (vbLf)
Public Function SectionNames(ByVal txt As String) As String
    Dim a() As String, i As Long, ln As String, r As String
    a = Split(Replace(txt, vbCrLf, vbLf), vbLf)
    For i = 0 To UBound(a)
        ln = Trim$(a(i))
        If Len(ln) > 2 Then
            If Left$(ln, 1) = "[" And Right$(ln, 1) = "]" Then
                If Len(r) > 0 Then r = r & vbLf
                r = r & Mid$(ln, 2, Len(ln) - 2)
            End If
        End If
    Next i
    SectionNames = r
End Function

' Cac dong "khoa<TAB>gia_tri" cua mot muc (theo thu tu trong tep), cach nhau vbLf
Public Function SectionItems(ByVal txt As String, ByVal sec As String) As String
    Dim a() As String, i As Long, ln As String, cur As String, p As Long, r As String
    a = Split(Replace(txt, vbCrLf, vbLf), vbLf)
    For i = 0 To UBound(a)
        ln = Trim$(a(i))
        If Len(ln) > 0 Then
            If Left$(ln, 1) = "[" Then
                cur = LCase$(Mid$(ln, 2, Len(ln) - 2))
            ElseIf Left$(ln, 1) <> ";" And Left$(ln, 1) <> "#" And cur = LCase$(sec) Then
                p = InStr(ln, "=")
                If p > 1 Then
                    If Len(r) > 0 Then r = r & vbLf
                    r = r & Trim$(Left$(ln, p - 1)) & vbTab & Trim$(Mid$(ln, p + 1))
                End If
            End If
        End If
    Next i
    SectionItems = r
End Function

' "levelN" -> "level" de dung chung kieu/goi y
Public Function Pat(ByVal sec As String) As String
    Dim s As String
    s = LCase$(sec)
    If Left$(s, 5) = "level" And Len(s) > 5 Then
        If IsNumeric(Mid$(s, 6)) Then s = "level"
    End If
    Pat = s
End Function

Public Function Hint(ByVal sec As String, ByVal key As String) As String
    Hint = Ui("hints." & Pat(sec) & "." & LCase$(key))
    If Len(Hint) = 0 Then Hint = key
End Function

Public Function IsReadOnly(ByVal sec As String, ByVal key As String) As Boolean
    IsReadOnly = (Ui("readonly." & Pat(sec) & "." & LCase$(key)) = "1")
End Function

Private Function IsNum(ByVal s As String) As Boolean
    Dim i As Long, ch As String, dots As Long
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch = "." Then
            dots = dots + 1
            If dots > 1 Or Len(s) = 1 Then Exit Function
        ElseIf ch < "0" Or ch > "9" Then
            Exit Function
        End If
    Next i
    IsNum = True
End Function

' Tra ve "" neu hop le, nguoc lai la ly do (co dau, dung trong form)
Public Function Validate(ByVal sec As String, ByVal key As String, ByVal value As String) As String
    Dim spec As String, a() As String, v As String, d As Double, lo As Double, hi As Double, i As Long, ok As Boolean, items() As String
    If InStr(value, vbCr) > 0 Or InStr(value, vbLf) > 0 Then Validate = Ui("text.t.line"): Exit Function
    spec = Ui("types." & Pat(sec) & "." & LCase$(key))
    If Len(spec) = 0 Then Exit Function
    a = Split(spec, ":")
    v = Replace(Trim$(value), ",", ".")
    Select Case LCase$(a(0))
        Case "num", "int"
            lo = Val(a(1)): hi = Val(a(2))
            ok = IsNum(v)
            If ok And LCase$(a(0)) = "int" Then ok = (InStr(v, ".") = 0)
            If ok Then d = Val(v): ok = (d >= lo And d <= hi)
            If Not ok Then Validate = Replace(Replace(Ui("text.t." & LCase$(a(0))), "{0}", a(1)), "{1}", a(2))
        Case "enum"
            items = Split(a(1), ",")
            For i = 0 To UBound(items)
                If StrComp(Trim$(value), items(i), vbTextCompare) = 0 Then ok = True
            Next i
            If Not ok Then Validate = Replace(Ui("text.t.enum"), "{0}", a(1))
        Case "bool"
            If Trim$(value) <> "0" And Trim$(value) <> "1" Then Validate = Ui("text.t.bool")
    End Select
End Function

Public Sub WriteUtf8(ByVal path As String, ByVal text As String)
    Dim st As Object, b As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.WriteText text
    st.Position = 0
    st.Type = 1
    st.Position = 3     ' bo BOM
    Set b = CreateObject("ADODB.Stream")
    b.Type = 1
    b.Open
    st.CopyTo b
    b.SaveToFile path, 2
    b.Close
    st.Close
End Sub

' edits: Dictionary "muc|khoa" (chu thuong) -> gia tri moi. Giu chu thich/thu tu/kieu xuong dong; sao luu .bak. Tra ve so dong da doi.
Public Function SaveEdits(ByVal path As String, ByVal edits As Object) As Long
    Dim txt As String, eol As String, a() As String, i As Long, ln As String, sec As String, p As Long, k As String, nv As String, n As Long
    txt = ReadText(path)
    If InStr(txt, vbCrLf) > 0 Then eol = vbCrLf Else eol = vbLf
    a = Split(Replace(txt, vbCrLf, vbLf), vbLf)
    For i = 0 To UBound(a)
        ln = Trim$(a(i))
        If Len(ln) > 0 Then
            If Left$(ln, 1) = "[" Then
                sec = LCase$(Mid$(ln, 2, Len(ln) - 2))
            ElseIf Left$(ln, 1) <> ";" And Left$(ln, 1) <> "#" Then
                p = InStr(a(i), "=")
                If p > 1 Then
                    k = sec & "|" & LCase$(Trim$(Left$(a(i), p - 1)))
                    If edits.Exists(k) Then
                        nv = Left$(a(i), p) & Trim$(edits(k))
                        If nv <> a(i) Then a(i) = nv: n = n + 1
                    End If
                End If
            End If
        End If
    Next i
    If n > 0 Then
        FileCopy path, path & ".bak"
        WriteUtf8 path, Join(a, eol)
    End If
    SaveEdits = n
End Function
