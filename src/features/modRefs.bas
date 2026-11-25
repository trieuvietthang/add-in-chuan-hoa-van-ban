Attribute VB_Name = "modRefs"
Option Explicit
' Dan chieu chong: kiem tra "Dieu N" / "khoan K Dieu N" / "diem x khoan K Dieu N", chuyen thanh REF field
' (tu cap nhat khi doi so), cap nhat, chen moi. Mau regex nam trong config\refs.vi.ini (chu Viet co dau).
' Chi tim dan chieu den muc "dieu" (levelN.refkind=dieu) trong ho so. Dan chieu ra van ban ngoai (cua Luat ...) bi bo qua.
' Luat da lam: R1 dich khong ton tai, R2 ten trong ngoac khop dieu khac, R4 dan chieu go tay, R5 khoan/diem khong thuoc dieu.

Private mConv As Long
Private mSkip As Long

Private Function LoadCfg() As Object
    Dim d As Object, f As String
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    f = modConfig.ConfigDir & "\refs.vi.ini"
    If Len(Dir$(f)) = 0 Then Err.Raise vbObjectError + 1005, "modRefs", modStrings.T("err.noconfig")
    modProfile.ParseIni modConfig.ReadUtf8(f), d
    Set LoadCfg = d
End Function

Private Function ReFirst(ByVal pat As String, ByVal s As String, ByVal ic As Boolean) As Object
    Dim re As Object, ms As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = pat
    re.IgnoreCase = ic
    re.Global = False
    Set ms = re.Execute(s)
    If ms.Count > 0 Then Set ReFirst = ms(0)
End Function

Private Function NumOf(ByVal s As String, ByVal first As Boolean) As String
    Dim m As Object
    If first Then
        Set m = ReFirst("\d+", s, False)
        If Not m Is Nothing Then NumOf = m.Value
    Else
        Set m = ReFirst("(\d+)\D*$", s, False)
        If Not m Is Nothing Then NumOf = m.SubMatches(0)
    End If
End Function

Private Function Norm(ByVal s As String) As String
    s = LCase$(Trim$(Replace(s, vbCr, " ")))
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    Norm = s
End Function

Private Function ParaTxt(ByVal p As Paragraph) As String
    ParaTxt = Replace(Replace(p.Range.Text, vbCr, ""), Chr(7), "")
End Function

' Chi muc: "d|N" -> vi tri dau doan; "t|N" -> tieu de; "k|N|K"; "p|N|K|x"
Private Function BuildIndex(ByVal doc As Document, ByVal prof As Object) As Object
    Dim ix As Object, p As Paragraph, lv As Long, kind As String, ls As String, curD As String, curK As String
    Set ix = CreateObject("Scripting.Dictionary")
    ix.CompareMode = 1
    For Each p In doc.Paragraphs
        lv = modProfile.LevelOfStyle(prof, p.Style.NameLocal)
        If lv > 0 Then
            kind = LCase$(modProfile.PGet(prof, modProfile.LevelKey(lv, "refkind")))
            If Len(kind) > 0 Then
                ls = p.Range.ListFormat.ListString
                Select Case kind
                    Case "dieu"
                        curD = NumOf(ls, True)
                        curK = ""
                        ix("d|" & curD) = p.Range.Start
                        ix("t|" & curD) = ParaTxt(p)
                    Case "khoan"
                        curK = NumOf(ls, False)
                        ix("k|" & curD & "|" & curK) = p.Range.Start
                    Case "diem"
                        ix("p|" & curD & "|" & curK & "|" & Replace(Replace(LCase$(ls), ")", ""), ".", "")) = p.Range.Start
                End Select
            End If
        End If
    Next p
    Set BuildIndex = ix
End Function

Private Function BmFor(ByVal doc As Document, ByVal pstart As Long) As String
    Dim b As Bookmark, pr As Range, nm As String, hid As Boolean
    Set pr = doc.Range(pstart, pstart).Paragraphs(1).Range
    hid = doc.Bookmarks.ShowHidden
    doc.Bookmarks.ShowHidden = True
    For Each b In doc.Bookmarks
        If Left$(b.Name, 5) = "_STC_" Then
            If b.Start >= pr.Start And b.Start < pr.End Then
                BmFor = b.Name
                doc.Bookmarks.ShowHidden = hid
                Exit Function
            End If
        End If
    Next b
    nm = "_STC_" & CStr(pr.Start)
    Do While doc.Bookmarks.Exists(nm)
        nm = nm & "x"
    Loop
    ' Bat dau tu ky tu thu 2: go Enter o dau doan (chen Dieu moi phia tren) khong keo bookmark sang doan moi
    If pr.End - pr.Start >= 3 Then
        doc.Bookmarks.Add nm, doc.Range(pr.Start + 1, pr.End - 1)
    Else
        doc.Bookmarks.Add nm, doc.Range(pr.Start, pr.End - 1)
    End If
    BmFor = nm
    doc.Bookmarks.ShowHidden = hid
End Function

Private Function InsideRef(ByVal rng As Range) As Boolean
    Dim f As Field
    For Each f In rng.Paragraphs(1).Range.Fields
        If f.Type = 3 Then
            If rng.Start >= f.Result.Start And rng.End <= f.Result.End Then
                InsideRef = True
                Exit Function
            End If
        End If
    Next f
End Function

Private Function Scan(ByVal doc As Document, ByVal profileName As String, ByVal convert As Boolean) As String
    Dim prof As Object, cfg As Object, ix As Object, rng As Range, pat As String, out As String
    Dim mt As String, n As String, ps As Long, pe As Long, pre As String, fwd As String, m As Object, m2 As Object
    Dim x As String, k As String, bad As Boolean, key As Variant, tt As String, nextStart As Long, fs As Long
    Dim f As Field, bm As String, lastEnd As Long
    lastEnd = -1
    Set prof = modProfile.LoadProfile(profileName)
    Set cfg = LoadCfg()
    Set ix = BuildIndex(doc, prof)
    If ix.Count = 0 Then Exit Function
    mConv = 0: mSkip = 0
    If convert Then
        For Each key In ix.Keys
            If Left$(CStr(key), 2) = "d|" Then bm = BmFor(doc, CLng(ix(key))): ix("b|" & Mid$(CStr(key), 3)) = bm
        Next key
    End If
    pat = modProfile.PGet(cfg, "refs.find_dieu")
    Set rng = doc.Content
    rng.Find.ClearFormatting
    Do
        If Not rng.Find.Execute(FindText:=pat, MatchWildcards:=True, Forward:=True, Wrap:=0) Then Exit Do
        mt = rng.Text
        n = NumOf(mt, True)
        nextStart = rng.End
        ' Word co the mo rong vung tim toi ca field -> gap lai cung ket qua; bao dam luon tien ve phia truoc
        If rng.End <= lastEnd Then
            nextStart = lastEnd + 1
        ElseIf Not InsideRef(rng) Then
            ps = rng.Paragraphs(1).Range.Start
            pe = rng.Paragraphs(1).Range.End
            pre = doc.Range(IIf(rng.Start - 40 > ps, rng.Start - 40, ps), rng.Start).Text
            fwd = doc.Range(rng.End, IIf(rng.End + 80 < pe, rng.End + 80, pe)).Text
            If ReFirst(modProfile.PGet(cfg, "refs.external"), fwd, True) Is Nothing Then
                x = "": k = ""
                Set m = ReFirst(modProfile.PGet(cfg, "refs.pre"), pre, True)
                If Not m Is Nothing Then
                    x = m.SubMatches(0) & ""
                    k = m.SubMatches(1) & ""
                End If
                bad = False
                If Not ix.Exists("d|" & n) Then
                    out = out & "R1|" & rng.Start & "|" & rng.End & "|" & n & vbLf
                    bad = True
                Else
                    If Len(k) > 0 Then
                        If Not ix.Exists("k|" & n & "|" & k) Then
                            out = out & "R5|" & rng.Start & "|" & rng.End & "|" & n & vbLf
                            bad = True
                        ElseIf Len(x) > 0 Then
                            If Not ix.Exists("p|" & n & "|" & k & "|" & LCase$(x)) Then
                                out = out & "R5|" & rng.Start & "|" & rng.End & "|" & n & vbLf
                                bad = True
                            End If
                        End If
                    End If
                    Set m2 = ReFirst(modProfile.PGet(cfg, "refs.title"), fwd, False)
                    If Not m2 Is Nothing Then
                        tt = Norm(m2.SubMatches(0))
                        If tt <> Norm(CStr(ix("t|" & n))) Then
                            For Each key In ix.Keys
                                If Left$(CStr(key), 2) = "t|" Then
                                    If Norm(CStr(ix(key))) = tt Then
                                        out = out & "R2|" & rng.Start & "|" & rng.End & "|" & n & vbLf
                                        bad = True
                                        Exit For
                                    End If
                                End If
                            Next key
                        End If
                    End If
                End If
                If Not bad Then
                    If convert Then
                        fs = rng.Start
                        Set f = doc.Fields.Add(rng, -1, "REF " & CStr(ix("b|" & n)) & " \n \h", False)
                        f.Update
                        If f.Result.Text <> mt Then
                            f.Delete
                            doc.Range(fs, fs).InsertAfter mt
                            mSkip = mSkip + 1
                            nextStart = fs + Len(mt)
                        Else
                            mConv = mConv + 1
                            nextStart = f.Result.End + 1
                        End If
                    Else
                        out = out & "R4|" & rng.Start & "|" & rng.End & "|" & n & vbLf
                    End If
                End If
            End If
        End If
        lastEnd = nextStart
        If nextStart >= doc.Content.End Then Exit Do
        rng.SetRange nextStart, doc.Content.End
    Loop
    Scan = out
End Function

' Moi dong "luat|bat dau|ket thuc|so dieu"; chuoi rong = khong co dan chieu noi bo nao
Public Function CheckRefs(ByVal doc As Document, ByVal profileName As String) As String
    CheckRefs = Scan(doc, profileName, False)
End Function

' Chuyen dan chieu "Dieu N" go tay hop le thanh REF field; tra ve so da chuyen. 1 UndoRecord.
Public Function ConvertRefs(ByVal doc As Document, ByVal profileName As String) As Long
    Dim en As Long, ed As String, dummy As String
    modUndo.BeginUndo doc, "Chuyen dan chieu"
    On Error GoTo EH
    dummy = Scan(doc, profileName, True)
    ConvertRefs = mConv
    modUndo.EndUndo
    Exit Function
EH:
    en = Err.Number: ed = Err.Description
    modUndo.EndUndo
    Err.Raise en, "modRefs", ed
End Function

Public Function UpdateRefs(ByVal doc As Document) As Long
    Dim f As Field, n As Long, en As Long, ed As String
    modUndo.BeginUndo doc, "Cap nhat dan chieu"
    On Error GoTo EH
    For Each f In doc.Fields
        If f.Type = 3 Then
            If InStr(f.Code.Text, "_STC_") > 0 Then
                f.Update
                n = n + 1
            End If
        End If
    Next f
    UpdateRefs = n
    modUndo.EndUndo
    Exit Function
EH:
    en = Err.Number: ed = Err.Description
    modUndo.EndUndo
    Err.Raise en, "modRefs", ed
End Function

' Chen dan chieu toi Dieu so n tai vi tri pos (REF field). Tra ve chu hien ra, rong neu khong co Dieu do.
Public Function InsertRef(ByVal doc As Document, ByVal pos As Long, ByVal dieuNo As Long, ByVal profileName As String) As String
    Dim prof As Object, ix As Object, f As Field, en As Long, ed As String
    modUndo.BeginUndo doc, "Chen dan chieu"
    On Error GoTo EH
    Set prof = modProfile.LoadProfile(profileName)
    Set ix = BuildIndex(doc, prof)
    If ix.Exists("d|" & CStr(dieuNo)) Then
        Set f = doc.Fields.Add(doc.Range(pos, pos), -1, "REF " & BmFor(doc, CLng(ix("d|" & CStr(dieuNo)))) & " \n \h", False)
        f.Update
        InsertRef = f.Result.Text
    End If
    modUndo.EndUndo
    Exit Function
EH:
    en = Err.Number: ed = Err.Description
    modUndo.EndUndo
    Err.Raise en, "modRefs", ed
End Function
