Attribute VB_Name = "modSpell"
Option Explicit
' Kiem tra chinh ta offline (M5, KB-D3). Chi BAO, khong sua chu. Du lieu (am tiet, cap sai-dung, nhom bien the...) nam trong config\spell.vi.ini.
' Ma loi: S1 hai dau cach, S2 dau cach truoc dau cau, S3 thieu dau cach sau dau cau, S4 lap tu, S5 ngoac/nhay khong can,
' S6 "dieu N" viet thuong dau cau, W1 cap sai-dung, W2 nhat quan bien the, W3 am tiet khong hop le, W4 thuat ngu hoa/thuong.
' Ket qua: moi dong "ma|bat_dau|ket_thuc|chi_tiet" (vi tri tuyet doi trong tai lieu), sap theo vi tri.

Private mCfg As Object
Private mTone As Object
Private mOnsets() As String
Private mNuclei As Object
Private mOpen As Object
Private mNeed As Object
Private mDiph As Object
Private mFrontOnly As Object
Private mFrontNever As Object
Private mCoda As Object
Private mIgnore As Object
Private mDupOk As Object
Private mCache As Object
Private mMarks As String
Private mS6 As String
Private mPairsK() As String
Private mPairsV() As String
Private mPairsN As Long
Private mTerms() As String
Private mTermsN As Long
Private mGroups() As String
Private mGroupsN As Long
Private mOut() As String
Private mKey() As Long
Private mN As Long
Private mW2() As String
Private mW2N As Long
Private mLoaded As Boolean
' Bang cap am tiet: bang bam dia chi mo bang mang (Scripting.Dictionary cham han khi >100k muc: nap ~5x, tra ~3x)
Private mBK() As String
Private mBV() As Long
Private mBMask As Long
' Ghi nho ket qua theo chu goc trong mot lan kiem tra (tu: W3, cap: W5) - van ban lap lai tu rat nhieu
Private mTK() As String
Private mTV() As String
Private mTMask As Long
Private mTN As Long
Private mPK() As String
Private mPV() As String
Private mPMask As Long
Private mPN As Long
Private mVar As Object
Private mBiLoaded As Boolean
Private mW5On As Boolean
Private mW5Min As Long
Private mLastW5 As Long
Private mUni As Object
Private mW5Ratio As Long

' ---------- cau hinh ----------
Private Function ToSet(ByVal csv As String) As Object
    Dim d As Object, a() As String, i As Long
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 0
    If Len(csv) > 0 Then
        a = Split(csv, ",")
        For i = 0 To UBound(a)
            If Len(Trim$(a(i))) > 0 Then d(Trim$(a(i))) = 1
        Next i
    End If
    Set ToSet = d
End Function

Private Sub LoadCfg()
    Dim f As String, k As Variant, s As String, j As Long, t As String, o() As String, i As Long, tmp As String, a() As String
    Set mCfg = CreateObject("Scripting.Dictionary")
    mCfg.CompareMode = 1
    f = modConfig.ConfigDir & "\spell.vi.ini"
    If Len(Dir$(f)) = 0 Then Err.Raise vbObjectError + 1006, "modSpell", modStrings.T("err.noconfig")
    modProfile.ParseIni modConfig.ReadUtf8(f), mCfg
    Set mTone = CreateObject("Scripting.Dictionary")
    mTone.CompareMode = 0
    For Each k In mCfg.Keys
        If Left$(CStr(k), 6) = "tones." Then
            s = mCfg(k)
            t = Mid$(CStr(k), 7)
            For j = 1 To Len(s)
                mTone(Mid$(s, j, 1)) = t & CStr(j - 1)
            Next j
        End If
    Next k
    mTone(ChrW(273)) = ChrW(273) & "0"
    mMarks = ChrW(259) & ChrW(226) & ChrW(234) & ChrW(244) & ChrW(417) & ChrW(432)
    o = Split(modProfile.PGet(mCfg, "syll.onsets"), ",")
    ReDim mOnsets(0 To UBound(o))
    For i = 0 To UBound(o)
        mOnsets(i) = Trim$(o(i))
    Next i
    For i = 0 To UBound(mOnsets) - 1
        For j = i + 1 To UBound(mOnsets)
            If Len(mOnsets(j)) > Len(mOnsets(i)) Then
                tmp = mOnsets(i): mOnsets(i) = mOnsets(j): mOnsets(j) = tmp
            End If
        Next j
    Next i
    Set mNuclei = ToSet(modProfile.PGet(mCfg, "syll.nuclei"))
    Set mOpen = ToSet(modProfile.PGet(mCfg, "syll.open_only"))
    Set mNeed = ToSet(modProfile.PGet(mCfg, "syll.need_coda"))
    Set mDiph = ToSet(modProfile.PGet(mCfg, "syll.diph"))
    Set mFrontOnly = ToSet(modProfile.PGet(mCfg, "syll.front_only"))
    Set mFrontNever = ToSet(modProfile.PGet(mCfg, "syll.front_never"))
    Set mCoda = CreateObject("Scripting.Dictionary")
    mCoda.CompareMode = 0
    For Each k In mCfg.Keys
        If Left$(CStr(k), 10) = "syll.coda." Then mCoda(Mid$(CStr(k), 11)) = "," & mCfg(k) & ","
    Next k
    Set mIgnore = ToSet(LCase$(modProfile.PGet(mCfg, "ignore.words")))
    Set mDupOk = ToSet(LCase$(modProfile.PGet(mCfg, "dup_ok.words")))
    mS6 = modProfile.PGet(mCfg, "rules.s6_term")
    ' cap sai-dung
    mPairsN = 0: mTermsN = 0: mGroupsN = 0
    ReDim mPairsK(0 To 63): ReDim mPairsV(0 To 63)
    ReDim mGroups(0 To 31)
    For Each k In mCfg.Keys
        If Left$(CStr(k), 6) = "pairs." Then
            If mPairsN > UBound(mPairsK) Then ReDim Preserve mPairsK(0 To mPairsN * 2): ReDim Preserve mPairsV(0 To mPairsN * 2)
            mPairsK(mPairsN) = LCase$(Mid$(CStr(k), 7)): mPairsV(mPairsN) = mCfg(k): mPairsN = mPairsN + 1
        ElseIf Left$(CStr(k), 7) = "groups." Then
            If mGroupsN > UBound(mGroups) Then ReDim Preserve mGroups(0 To mGroupsN * 2)
            mGroups(mGroupsN) = LCase$(mCfg(k)): mGroupsN = mGroupsN + 1
        End If
    Next k
    s = modProfile.PGet(mCfg, "terms.list")
    If Len(s) > 0 Then
        a = Split(s, ",")
        ReDim mTerms(0 To UBound(a))
        For i = 0 To UBound(a)
            mTerms(i) = Trim$(a(i))
        Next i
        mTermsN = UBound(a) + 1
    End If
    Set mCache = CreateObject("Scripting.Dictionary")
    mCache.CompareMode = 0
    mLoaded = True
End Sub

' Goi sau khi sua spell.vi.ini (form Sua ho so) de nap lai cau hinh o lan kiem tra ke tiep
Public Sub ResetConfig()
    mLoaded = False
    mBiLoaded = False
End Sub

' ---------- am tiet (thuan) ----------
Public Function SyllableOk(ByVal w As String) As Boolean
    If Not mLoaded Then LoadCfg
    SyllableOk = CheckSyll(LCase$(w))
End Function

Private Function CheckSyll(ByVal w As String) As Boolean
    Dim b As String, i As Long, ch As String, v As String, tp As Long, ti As Long, n As Long, oi As Long, o As String, lo As Long
    n = Len(w)
    If n = 0 Then Exit Function
    For i = 1 To n
        ch = Mid$(w, i, 1)
        If mTone.Exists(ch) Then
            v = mTone(ch)
            b = b & Left$(v, Len(v) - 1)
            If Right$(v, 1) <> "0" Then
                If tp > 0 Then Exit Function
                tp = i: ti = CLng(Right$(v, 1))
            End If
        ElseIf ch >= "a" And ch <= "z" Then
            b = b & ch
        Else
            Exit Function
        End If
    Next i
    For oi = 0 To UBound(mOnsets) + 1
        If oi <= UBound(mOnsets) Then o = mOnsets(oi) Else o = ""
        lo = Len(o)
        If Left$(b, lo) = o Then
            If RestOk(b, lo, tp, ti, o) Then
                CheckSyll = True
                Exit Function
            End If
        End If
    Next oi
End Function

Private Function RestOk(ByVal b As String, ByVal lo As Long, ByVal tp As Long, ByVal ti As Long, ByVal o As String) As Boolean
    Dim rest As String, cod As Variant, c As String, nuc As String
    If tp > 0 And tp <= lo Then Exit Function
    rest = Mid$(b, lo + 1)
    If Len(rest) = 0 Then Exit Function
    For Each cod In Array("", "c", "ch", "m", "n", "ng", "nh", "p", "t")
        c = CStr(cod)
        If Len(c) <= Len(rest) - 1 Then
            If Right$(rest, Len(c)) = c Then
                nuc = Left$(rest, Len(rest) - Len(c))
                If NucOk(nuc, c, o, lo, tp, ti) Then
                    RestOk = True
                    Exit Function
                End If
            End If
        End If
    Next cod
End Function

Private Function NucOk(ByVal nuc As String, ByVal c As String, ByVal o As String, ByVal lo As Long, ByVal tp As Long, ByVal ti As Long) As Boolean
    Dim first As String, isFront As Boolean, key As String
    If Not mNuclei.Exists(nuc) Then Exit Function
    first = Left$(nuc, 1)
    isFront = (first = "e" Or first = ChrW(234) Or first = "i" Or first = "y")
    If mFrontOnly.Exists(o) Then
        If Not isFront Then Exit Function
    End If
    If mFrontNever.Exists(o) Then
        If isFront Then
            If Not (o = "g" And first = "i") Then Exit Function
        End If
    End If
    If o = "qu" Then
        If first = "u" Or first = "o" Or first = ChrW(432) Then Exit Function
    End If
    If o = "gi" Then
        If first = "i" Or first = "y" Then Exit Function
    End If
    If lo = 0 Then
        If Left$(nuc, 2) = "i" & ChrW(234) Then Exit Function
    End If
    If Len(c) = 0 Then
        If mNeed.Exists(nuc) Then Exit Function
    Else
        If mOpen.Exists(nuc) Then Exit Function
        If mDiph.Exists(nuc) Then key = "diph" Else key = Right$(nuc, 1)
        If Not mCoda.Exists(key) Then Exit Function
        If InStr(mCoda(key), "," & c & ",") = 0 Then Exit Function
        If c = "c" Or c = "ch" Or c = "p" Or c = "t" Then
            If ti <> 2 And ti <> 5 Then Exit Function
        End If
    End If
    If ti > 0 Then
        If Not ToneOk(nuc, c, tp - lo) Then Exit Function
    End If
    NucOk = True
End Function

Private Function ToneOk(ByVal nuc As String, ByVal c As String, ByVal rel As Long) As Boolean
    Dim n As Long, i As Long, last As Long
    n = Len(nuc)
    If rel < 1 Or rel > n Then Exit Function
    If n = 1 Then
        ToneOk = True
        Exit Function
    End If
    For i = 1 To n
        If InStr(mMarks, Mid$(nuc, i, 1)) > 0 Then last = i
    Next i
    If last > 0 Then
        ToneOk = (rel = last)
    ElseIf n = 2 Then
        If nuc = "oa" Or nuc = "oe" Or nuc = "uy" Then
            If Len(c) = 0 Then ToneOk = (rel = 1 Or rel = 2) Else ToneOk = (rel = 2)
        Else
            ToneOk = (rel = 1)
        End If
    Else
        ToneOk = (rel = 2)
    End If
End Function

' Bo dau cho thong bao (MsgBox ANSI)
Public Function Plain(ByVal s As String) As String
    Dim i As Long, ch As String, v As String, r As String
    If Not mLoaded Then LoadCfg
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If mTone.Exists(ch) Then
            v = mTone(ch)
            r = r & Left$(v, Len(v) - 1)
        ElseIf AscW(ch) = 272 Then
            r = r & "D"
        Else
            r = r & ch
        End If
    Next i
    ' nguyen am co dau mu/ranh (a, e, o, u, d) ve chu goc
    r = Replace(r, ChrW(259), "a"): r = Replace(r, ChrW(226), "a"): r = Replace(r, ChrW(234), "e")
    r = Replace(r, ChrW(244), "o"): r = Replace(r, ChrW(417), "o"): r = Replace(r, ChrW(432), "u"): r = Replace(r, ChrW(273), "d")
    Plain = r
End Function

' ---------- tien ich ky tu ----------
Private Function IsLetterW(ByVal code As Long) As Boolean
    If (code >= 65 And code <= 90) Or (code >= 97 And code <= 122) Then
        IsLetterW = True
    ElseIf code >= 192 And code <= 591 And code <> 215 And code <> 247 Then
        IsLetterW = True
    ElseIf code >= &H1E00 And code <= &H1EFF Then
        IsLetterW = True
    End If
End Function

Private Function IsBreak(ByVal ch As String) As Boolean
    IsBreak = (ch = vbCr Or ch = Chr(7) Or ch = vbTab Or ch = Chr(11) Or ch = Chr(12) Or ch = Chr(1) Or ch = "")
End Function

Private Function IsDigitCh(ByVal ch As String) As Boolean
    IsDigitCh = (ch >= "0" And ch <= "9" And Len(ch) = 1)
End Function

Private Function IsLowerCh(ByVal ch As String) As Boolean
    IsLowerCh = (LCase$(ch) = ch And UCase$(ch) <> ch)
End Function

Private Function IsUpperCh(ByVal ch As String) As Boolean
    IsUpperCh = (UCase$(ch) = ch And LCase$(ch) <> ch)
End Function

Private Function IsLetS(ByRef c As String, ByVal i As Long) As Boolean
    If i < 1 Or i > Len(c) Then Exit Function
    IsLetS = IsLetterW(AscW(Mid$(c, i, 1)))
End Function

' Cum giua hai khoang trang quanh vi tri i co phai dia chi web/email/duong dan
Private Function IsUrlish(ByRef c As String, ByVal i As Long) As Boolean
    Dim a As Long, b As Long, chunk As String
    a = i: b = i
    Do While a > 1
        If Mid$(c, a - 1, 1) = " " Then Exit Do
        a = a - 1
    Loop
    Do While b < Len(c)
        If Mid$(c, b + 1, 1) = " " Then Exit Do
        b = b + 1
    Loop
    chunk = LCase$(Mid$(c, a, b - a + 1))
    IsUrlish = (InStr(chunk, "@") > 0 Or InStr(chunk, "http") > 0 Or InStr(chunk, "www") > 0 Or InStr(chunk, "\") > 0 Or InStr(chunk, "/") > 0)
End Function

Private Function WordLenBack(ByRef c As String, ByVal i As Long) As Long
    Do While i >= 1
        If Not IsLetS(c, i) Then Exit Do
        WordLenBack = WordLenBack + 1
        i = i - 1
    Loop
End Function

Private Function BoundaryOk(ByRef c As String, ByVal s As Long, ByVal e As Long) As Boolean
    ' s..e-1 la doan khop; hai ben khong duoc la chu cai
    BoundaryOk = (Not IsLetS(c, s - 1)) And (Not IsLetS(c, e))
End Function

' ---------- ghi ket qua ----------
Private Sub EmitAbs(ByVal rule As String, ByVal s As Long, ByVal e As Long, ByVal det As String)
    If mN > UBound(mOut) Then
        ReDim Preserve mOut(0 To mN * 2)
        ReDim Preserve mKey(0 To mN * 2)
    End If
    mOut(mN) = rule & "|" & CStr(s) & "|" & CStr(e) & "|" & det
    mKey(mN) = s
    mN = mN + 1
End Sub

Private Sub Emit(ByVal rule As String, ByVal base As Long, ByVal i1 As Long, ByVal i2 As Long, ByVal det As String)
    EmitAbs rule, base + i1 - 1, base + i2 - 1, det
End Sub

Private Sub SortOut()
    Dim i As Long, j As Long, k As Long, tk As Long, ts As String
    ' chen (on dinh): phan lon da co thu tu
    For i = 1 To mN - 1
        k = mKey(i): ts = mOut(i)
        j = i - 1
        Do While j >= 0
            If mKey(j) <= k Then Exit Do
            mKey(j + 1) = mKey(j): mOut(j + 1) = mOut(j)
            j = j - 1
        Loop
        mKey(j + 1) = k: mOut(j + 1) = ts
    Next i
End Sub

' ---------- tang 1: kieu chu / dau cau ----------
Private Sub ScanTyp(ByVal c As String, ByVal base As Long)
    Dim n As Long, i As Long, j As Long, ch As String, pv As String, nx As String, n3 As String
    n = Len(c)
    i = 1
    Do While i <= n
        ch = Mid$(c, i, 1)
        If ch = " " Then
            j = i
            Do While j < n
                If Mid$(c, j + 1, 1) <> " " Then Exit Do
                j = j + 1
            Loop
            If i > 1 And j < n Then
                pv = Mid$(c, i - 1, 1)
                nx = Mid$(c, j + 1, 1)
                If j > i Then
                    If Not IsBreak(pv) And Not IsBreak(nx) Then Emit "S1", base, i, j + 1, ""
                Else
                    If Not IsBreak(pv) Then
                        If InStr(",;:!?)", nx) > 0 Then
                            Emit "S2", base, i, i + 2, ""
                        ElseIf nx = "." And pv <> "." Then
                            If i + 2 > n Then
                                n3 = ""
                            Else
                                n3 = Mid$(c, i + 2, 1)
                            End If
                            If n3 = " " Or IsBreak(n3) Then Emit "S2", base, i, i + 2, ""
                        End If
                    End If
                End If
            End If
            i = j + 1
        Else
            If ch = "," Or ch = ";" Then
                If i > 1 And IsLetS(c, i + 1) Then
                    If Not IsUrlish(c, i) Then Emit "S3", base, i, i + 1, ""
                End If
            ElseIf ch = "." Then
                If i > 1 And i < n Then
                    If IsLowerCh(Mid$(c, i - 1, 1)) And IsUpperCh(Mid$(c, i + 1, 1)) Then
                        If WordLenBack(c, i - 1) >= 3 Then
                            If Not IsUrlish(c, i) Then Emit "S3", base, i, i + 1, ""
                        End If
                    End If
                End If
            ElseIf ch = ":" Then
                If i > 1 And i < n Then
                    If IsLowerCh(Mid$(c, i - 1, 1)) And IsLetS(c, i + 1) Then
                        If Not IsUrlish(c, i) Then Emit "S3", base, i, i + 1, ""
                    End If
                End If
            End If
            i = i + 1
        End If
    Loop
End Sub

Private Function IsListMark(ByRef c As String, ByVal i As Long) As Boolean
    Dim j As Long, m As Long, ln As Long, ch As String
    j = i - 1
    Do While j >= 1
        ch = Mid$(c, j, 1)
        If Not (IsLetterW(AscW(ch)) Or IsDigitCh(ch)) Then Exit Do
        j = j - 1
    Loop
    ln = i - 1 - j
    If ln < 1 Or ln > 4 Then Exit Function
    m = j
    Do While m >= 1
        If Mid$(c, m, 1) <> " " Then Exit Do
        m = m - 1
    Loop
    IsListMark = (m < 1)
End Function

Private Sub Balance(ByVal c As String, ByVal base As Long, ByVal o As String, ByVal cl As String)
    Dim i As Long, n As Long, depth As Long, firstOpen As Long, ch As String
    n = Len(c)
    For i = 1 To n
        ch = Mid$(c, i, 1)
        If ch = o Then
            If depth = 0 Then firstOpen = i
            depth = depth + 1
        ElseIf ch = cl Then
            If depth > 0 Then
                depth = depth - 1
            ElseIf cl = ")" Then
                If Not IsListMark(c, i) Then
                    Emit "S5", base, i, i + 1, ""
                    Exit Sub
                End If
            Else
                Emit "S5", base, i, i + 1, ""
                Exit Sub
            End If
        End If
    Next i
    If depth > 0 Then Emit "S5", base, firstOpen, firstOpen + 1, ""
End Sub

Private Sub ScanBrackets(ByVal c As String, ByVal base As Long)
    Dim k As Long, cnt As Long, p As Long
    If InStr(c, "(") + InStr(c, ")") + InStr(c, "[") + InStr(c, "]") + InStr(c, ChrW(8220)) + InStr(c, ChrW(8221)) + InStr(c, """") = 0 Then Exit Sub
    Balance c, base, "(", ")"
    Balance c, base, "[", "]"
    Balance c, base, ChrW(8220), ChrW(8221)
    p = InStr(c, """")
    If p > 0 Then
        k = p
        cnt = 0
        Do While k > 0
            cnt = cnt + 1
            k = InStr(k + 1, c, """")
        Loop
        If cnt Mod 2 = 1 Then Emit "S5", base, p, p + 1, ""
    End If
End Sub

Private Sub ScanS6(ByVal c As String, ByVal base As Long)
    Dim p As Long, q As Long, ch As String, ok As Boolean
    If Len(mS6) = 0 Then Exit Sub
    p = InStr(1, c, mS6 & " ", vbBinaryCompare)
    Do While p > 0
        If p > 1 And p + Len(mS6) + 1 <= Len(c) Then
            If IsDigitCh(Mid$(c, p + Len(mS6) + 1, 1)) And Mid$(c, p - 1, 1) = " " Then
                q = p - 1
                Do While q >= 1
                    If Mid$(c, q, 1) <> " " Then Exit Do
                    q = q - 1
                Loop
                If q >= 1 Then
                    ch = Mid$(c, q, 1)
                    If ch = "." Or ch = "?" Or ch = "!" Then Emit "S6", base, p, p + Len(mS6), ""
                End If
            End If
        End If
        p = InStr(p + 1, c, mS6 & " ", vbBinaryCompare)
    Loop
End Sub

' ---------- tang 2/4: cap sai-dung, thuat ngu, nhom bien the ----------
Private Sub ScanDict(ByVal c As String, ByVal base As Long)
    Dim lc As String, k As Long, p As Long, s As String, a() As String, vi As Long, gi As Long
    lc = LCase$(c)
    For k = 0 To mPairsN - 1
        s = mPairsK(k)
        p = InStr(1, lc, s, vbBinaryCompare)
        Do While p > 0
            If BoundaryOk(lc, p, p + Len(s)) Then Emit "W1", base, p, p + Len(s), mPairsV(k)
            p = InStr(p + 1, lc, s, vbBinaryCompare)
        Loop
    Next k
    For k = 0 To mTermsN - 1
        s = LCase$(mTerms(k))
        p = InStr(1, lc, s, vbBinaryCompare)
        Do While p > 0
            If BoundaryOk(lc, p, p + Len(s)) Then
                ' Bo qua chu IN HOA toan bo (tieu de "BEN B") - chi bao khi lan lon hoa/thuong sai
                If StrComp(Mid$(c, p, Len(s)), mTerms(k), vbBinaryCompare) <> 0 And StrComp(Mid$(c, p, Len(s)), UCase$(mTerms(k)), vbBinaryCompare) <> 0 Then Emit "W4", base, p, p + Len(s), mTerms(k)
            End If
            p = InStr(p + 1, lc, s, vbBinaryCompare)
        Loop
    Next k
    For gi = 0 To mGroupsN - 1
        a = Split(mGroups(gi), "|")
        For vi = 0 To UBound(a)
            s = a(vi)
            If Len(s) > 0 Then
                p = InStr(1, lc, s, vbBinaryCompare)
                Do While p > 0
                    If BoundaryOk(lc, p, p + Len(s)) Then
                        If mW2N > UBound(mW2) Then ReDim Preserve mW2(0 To mW2N * 2)
                        mW2(mW2N) = CStr(gi) & "|" & CStr(vi) & "|" & CStr(base + p - 1) & "|" & CStr(base + p - 1 + Len(s))
                        mW2N = mW2N + 1
                    End If
                    p = InStr(p + 1, lc, s, vbBinaryCompare)
                Loop
            End If
        Next vi
    Next gi
End Sub

Private Sub FinishW2()
    Dim gi As Long, vi As Long, i As Long, a() As String, f() As String, cnt() As Long, nv As Long, used As Long, best As Long
    For gi = 0 To mGroupsN - 1
        a = Split(mGroups(gi), "|")
        nv = UBound(a) + 1
        ReDim cnt(0 To nv - 1)
        For i = 0 To mW2N - 1
            f = Split(mW2(i), "|")
            If CLng(f(0)) = gi Then cnt(CLng(f(1))) = cnt(CLng(f(1))) + 1
        Next i
        used = 0: best = 0
        For vi = 0 To nv - 1
            If cnt(vi) > 0 Then used = used + 1
            If cnt(vi) > cnt(best) Then best = vi
        Next vi
        If used >= 2 Then
            For i = 0 To mW2N - 1
                f = Split(mW2(i), "|")
                If CLng(f(0)) = gi And CLng(f(1)) <> best Then EmitAbs "W2", CLng(f(2)), CLng(f(3)), a(best)
            Next i
        End If
    Next gi
End Sub

' ---------- tang 5: cap am tiet la (W5) ----------
' ---------- bang bam dia chi mo (khoa String khac rong, so sanh nhi phan) ----------
Private Function HashS(ByRef k As String) As Long
    Dim b() As Byte, i As Long, h As Long
    b = k
    For i = 0 To UBound(b) Step 2
        h = ((h * 31) + b(i) + b(i + 1) * 7) And &HFFFFF
    Next i
    HashS = h
End Function

' Vi tri cua khoa k trong bang (o chua k, hoac o trong dau tien neu chua co). Bang phai con o trong.
Private Function HSlot(ByRef tk() As String, ByVal mask As Long, ByRef k As String) As Long
    Dim p As Long
    p = HashS(k) And mask
    Do While Len(tk(p)) > 0
        If tk(p) = k Then Exit Do
        p = (p + 1) And mask
    Loop
    HSlot = p
End Function

' So lan cua cap "a b" trong ngu lieu; -1 neu khong co
Private Function BiCount(ByRef k As String) As Long
    Dim p As Long
    p = HSlot(mBK, mBMask, k)
    If Len(mBK(p)) > 0 Then BiCount = mBV(p) Else BiCount = -1
End Function

Private Sub ResetTokMemo()
    mTMask = &H7FFF&: mTN = 0     ' hau to & bat buoc: &HFFFF khong co & la Integer = -1
    ReDim mTK(0 To mTMask): ReDim mTV(0 To mTMask)
End Sub

Private Sub ResetPairMemo()
    mPMask = &HFFFF&: mPN = 0
    ReDim mPK(0 To mPMask): ReDim mPV(0 To mPMask)
End Sub

' Du lieu: config\bigram.vi.txt ("u<TAB>am_tiet<TAB>n" va "b<TAB>a b<TAB>n"). Tuy chon: thieu tep thi bo qua W5.
Private Sub LoadBigrams()
    Dim f As String, a() As String, i As Long, ln As String, t As Long, nm As String, n As Long, k As String, sz As Long, nb As Long
    mBiLoaded = True
    mBMask = 0
    ReDim mBK(0 To 0): ReDim mBV(0 To 0)
    Set mVar = CreateObject("Scripting.Dictionary")
    mVar.CompareMode = 0
    Set mUni = CreateObject("Scripting.Dictionary")
    mUni.CompareMode = 0
    mW5On = (modProfile.PGet(mCfg, "w5.enabled") <> "0")
    mW5Min = 40
    If IsNumeric(modProfile.PGet(mCfg, "w5.min_alt")) Then mW5Min = CLng(modProfile.PGet(mCfg, "w5.min_alt"))
    mW5Ratio = 25
    If IsNumeric(modProfile.PGet(mCfg, "w5.min_ratio")) Then mW5Ratio = CLng(modProfile.PGet(mCfg, "w5.min_ratio"))
    f = modConfig.ConfigDir & "\bigram.vi.txt"
    If Len(Dir$(f)) = 0 Then mW5On = False
    If Not mW5On Then Exit Sub
    a = Split(modConfig.ReadUtf8(f), vbLf)
    ' kich thuoc luy thua 2, >= 3 lan so dong (ty le lap day <= 1/3), toi da 2^20 (gioi han cua HashS)
    sz = 1024
    Do While sz < (UBound(a) + 1) * 3 And sz < &H100000
        sz = sz * 2
    Loop
    mBMask = sz - 1
    ReDim mBK(0 To mBMask): ReDim mBV(0 To mBMask)
    For i = 0 To UBound(a)
        ln = a(i)
        If Len(ln) > 4 Then
            If Right$(ln, 1) = vbCr Then ln = Left$(ln, Len(ln) - 1)
            t = InStr(3, ln, vbTab)
            If t > 3 And Left$(ln, 1) <> ";" Then
                nm = Mid$(ln, 3, t - 3)
                n = CLng(Mid$(ln, t + 1))
                If Left$(ln, 1) = "b" Then
                    t = HSlot(mBK, mBMask, nm)
                    If Len(mBK(t)) = 0 Then
                        ' bang day toi 3/4 (chi khi tep > ~1,4 trieu dong): bo qua cap moi thay vi lap vo han
                        If nb < (mBMask \ 4) * 3 Then mBK(t) = nm: mBV(t) = n: nb = nb + 1
                    Else
                        mBV(t) = n
                    End If
                Else
                    mUni(nm) = n
                    k = Plain(nm)
                    If mVar.Exists(k) Then mVar(k) = mVar(k) & "|" & nm Else mVar(k) = nm
                End If
            End If
        End If
    Next i
End Sub

Private Function SyllCached(ByVal lw As String) As Boolean
    If Not mCache.Exists(lw) Then
        If CheckSyll(lw) Then mCache(lw) = "1" Else mCache(lw) = "0"
    End If
    SyllCached = (mCache(lw) = "1")
End Function

Private Sub BestAlt(ByVal w As String, ByVal other As String, ByVal isLeft As Boolean, ByRef bestN As Long, ByRef bestW As String)
    Dim v() As String, i As Long, key As String, n As Long, u0 As Long
    key = Plain(w)
    If Not mVar.Exists(key) Then Exit Sub
    If mUni.Exists(w) Then u0 = mUni(w)
    If u0 < 1 Then u0 = 1
    v = Split(mVar(key), "|")
    For i = 0 To UBound(v)
        If v(i) <> w And mUni(v(i)) >= u0 * mW5Ratio Then
            If isLeft Then key = v(i) & " " & other Else key = other & " " & v(i)
            n = BiCount(key)
            If n > bestN Then bestN = n: bestW = v(i)
        End If
    Next i
End Sub

' Cap am tiet hop le nhung khong co trong ngu lieu, trong khi doi dau/ma o mot am tiet cho ra cap pho bien -> nghi go sai dau
Private Sub CheckPair(ByRef c As String, ByVal base As Long, ByVal s1 As Long, ByVal e1 As Long, ByVal s2 As Long, ByVal e2 As Long)
    Dim x As String, y As String, k As String, p As Long, r As String
    x = Mid$(c, s1, e1 - s1): y = Mid$(c, s2, e2 - s2)
    k = x & " " & y
    p = HSlot(mPK, mPMask, k)
    If Len(mPK(p)) > 0 Then
        r = mPV(p)
    Else
        r = PairVerdict(x, y)
        If mPN >= mPMask \ 2 Then
            ResetPairMemo
            p = HSlot(mPK, mPMask, k)
        End If
        mPK(p) = k: mPV(p) = r: mPN = mPN + 1
    End If
    If Len(r) = 0 Then Exit Sub
    ' phan phu thuoc vi tri (dinh so/ky hieu, chong bao trung) - khong ghi nho
    If Glued(c, s1, e1) Or Glued(c, s2, e2) Then Exit Sub
    If Left$(r, 1) = "1" Then
        If s1 <= mLastW5 Then Exit Sub
        Emit "W5", base, s1, e1, Mid$(r, 2)
        mLastW5 = e1
    Else
        If s2 <= mLastW5 Then Exit Sub
        Emit "W5", base, s2, e2, Mid$(r, 2)
        mLastW5 = e2
    End If
End Sub

' Phan chi phu thuoc chu cua cap: "" = khong nghi; "1<goi y>" = nghi am tiet trai; "2<goi y>" = nghi am tiet phai
Private Function PairVerdict(ByVal x As String, ByVal y As String) As String
    Dim a As String, b As String, n1 As Long, n2 As Long, w1 As String, w2 As String
    If Len(x) < 2 Or Len(y) < 2 Then Exit Function
    If x = UCase$(x) Or y = UCase$(y) Then Exit Function
    a = LCase$(x): b = LCase$(y)
    If a = b Then Exit Function
    If BiCount(a & " " & b) >= 0 Then Exit Function
    If mIgnore.Exists(a) Or mIgnore.Exists(b) Or mIgnore.Exists(a & " " & b) Then Exit Function
    If Not SyllCached(a) Then Exit Function
    If Not SyllCached(b) Then Exit Function
    BestAlt a, b, True, n1, w1
    BestAlt b, a, False, n2, w2
    If n1 < mW5Min And n2 < mW5Min Then Exit Function
    If n1 >= n2 Then PairVerdict = "1" & w1 & " " & b Else PairVerdict = "2" & a & " " & w2
End Function

' ---------- tang 3 + lap tu ----------
Private Function Glued(ByRef c As String, ByVal s As Long, ByVal e As Long) As Boolean
    Dim p As String, q As String
    If s > 1 Then p = Mid$(c, s - 1, 1)
    If e <= Len(c) Then q = Mid$(c, e, 1)
    If IsDigitCh(p) Or IsDigitCh(q) Then Glued = True: Exit Function
    If p = "@" Or q = "@" Or p = "_" Or q = "_" Or p = "\" Or q = "\" Or p = "/" Or q = "/" Then Glued = True: Exit Function
    If p = "." And s > 2 Then If IsLetS(c, s - 2) Then Glued = True: Exit Function
    If q = "." And IsLetS(c, e + 1) Then Glued = True
End Function

Private Function TokenBad(ByRef c As String, ByVal s As Long, ByVal e As Long) As Boolean
    Dim txt As String, p As Long, r As String
    If e - s < 2 Then Exit Function
    txt = Mid$(c, s, e - s)
    p = HSlot(mTK, mTMask, txt)
    If Len(mTK(p)) > 0 Then
        r = mTV(p)
    Else
        If TokenTextBad(txt) Then r = "1" Else r = "0"
        If mTN >= mTMask \ 2 Then
            ResetTokMemo
            p = HSlot(mTK, mTMask, txt)
        End If
        mTK(p) = txt: mTV(p) = r: mTN = mTN + 1
    End If
    ' phan phu thuoc vi tri - khong ghi nho
    If r = "1" Then TokenBad = Not Glued(c, s, e)
End Function

' Phan chi phu thuoc chu cua am tiet (W3)
Private Function TokenTextBad(ByVal txt As String) As Boolean
    Dim lw As String, i As Long, hasHi As Boolean, r As String
    If txt = UCase$(txt) Then Exit Function
    If txt <> LCase$(txt) Then
        If Mid$(txt, 2) <> LCase$(Mid$(txt, 2)) Then Exit Function
    End If
    lw = LCase$(txt)
    If InStr(lw, "f") + InStr(lw, "j") + InStr(lw, "w") + InStr(lw, "z") > 0 Then Exit Function
    For i = 1 To Len(lw)
        If AscW(Mid$(lw, i, 1)) > 127 Then hasHi = True: Exit For
    Next i
    If Not hasHi Then Exit Function
    If mIgnore.Exists(lw) Then Exit Function
    TokenTextBad = Not SyllCached(lw)
End Function

Private Sub ScanTokens(ByVal c As String, ByVal base As Long)
    Dim n As Long, i As Long, s As Long, tn As Long, k As Long, ts() As Long, te() As Long, w1 As String, w2 As String
    n = Len(c)
    If n = 0 Then Exit Sub
    ReDim ts(1 To n \ 2 + 2)
    ReDim te(1 To n \ 2 + 2)
    i = 1
    Do While i <= n
        If IsLetterW(AscW(Mid$(c, i, 1))) Then
            s = i
            Do While i <= n
                If Not IsLetterW(AscW(Mid$(c, i, 1))) Then Exit Do
                i = i + 1
            Loop
            tn = tn + 1: ts(tn) = s: te(tn) = i
        Else
            i = i + 1
        End If
    Loop
    mLastW5 = 0
    For k = 1 To tn
        If TokenBad(c, ts(k), te(k)) Then Emit "W3", base, ts(k), te(k), ""
        If k < tn Then
            If te(k) + 1 = ts(k + 1) And Mid$(c, te(k), 1) = " " Then
                If mW5On Then CheckPair c, base, ts(k), te(k), ts(k + 1), te(k + 1)
                If te(k) - ts(k) = te(k + 1) - ts(k + 1) Then
                    w1 = LCase$(Mid$(c, ts(k), te(k) - ts(k)))
                    w2 = LCase$(Mid$(c, ts(k + 1), te(k + 1) - ts(k + 1)))
                    If w1 = w2 And Len(w1) >= 2 Then
                        If Not mDupOk.Exists(w1 & " " & w2) Then
                            If Not Glued(c, ts(k), te(k)) And Not Glued(c, ts(k + 1), te(k + 1)) Then Emit "S4", base, ts(k), te(k + 1), ""
                        End If
                    End If
                End If
            End If
        End If
    Next k
End Sub

' ---------- che ma truong (field code) ----------
Private Function MaskCodes(ByVal t As String) As String
    Dim i As Long, n As Long, code As Long, depth As Long, top As Long, st() As Boolean, c As String
    c = t
    n = Len(t)
    ReDim st(0 To 15)
    For i = 1 To n
        code = AscW(Mid$(t, i, 1))
        If code = 19 Then
            top = top + 1
            If top > UBound(st) Then ReDim Preserve st(0 To top * 2)
            st(top) = True
            depth = depth + 1
            Mid$(c, i, 1) = Chr(1)
        ElseIf code = 20 Then
            If top > 0 Then
                If st(top) Then st(top) = False: depth = depth - 1
            End If
            Mid$(c, i, 1) = Chr(1)
        ElseIf code = 21 Then
            If top > 0 Then
                If st(top) Then depth = depth - 1
                top = top - 1
            End If
            Mid$(c, i, 1) = Chr(1)
        ElseIf depth > 0 Then
            Mid$(c, i, 1) = Chr(1)
        End If
    Next i
    MaskCodes = c
End Function

' ---------- vong chinh ----------
Public Function CheckSpell(ByVal doc As Document, Optional ByVal profile As String = "") As String
    Dim p As Paragraph, r As Range, t As String, c As String, base As Long, hasFields As Boolean
    Dim tocS() As Long, tocE() As Long, tocN As Long, i As Long, skip As Boolean, tc As TableOfContents, ln() As String
    LoadCfg
    If Not mBiLoaded Then LoadBigrams
    ResetTokMemo
    ResetPairMemo
    mN = 0: mW2N = 0
    ReDim mOut(0 To 255): ReDim mKey(0 To 255): ReDim mW2(0 To 63)
    tocN = doc.TablesOfContents.Count
    If tocN > 0 Then
        ReDim tocS(1 To tocN): ReDim tocE(1 To tocN)
        i = 0
        For Each tc In doc.TablesOfContents
            i = i + 1
            tocS(i) = tc.Range.Start: tocE(i) = tc.Range.End
        Next tc
    End If
    hasFields = (doc.Fields.Count > 0)
    For Each p In doc.Paragraphs
        Set r = p.Range
        skip = False
        base = r.Start
        For i = 1 To tocN
            If base >= tocS(i) And base < tocE(i) Then skip = True: Exit For
        Next i
        If Not skip Then
            r.TextRetrievalMode.IncludeHiddenText = True
            If hasFields Then r.TextRetrievalMode.IncludeFieldCodes = True
            t = r.Text
            If Len(t) > 1 Then
                If InStr(t, Chr(19)) > 0 Then c = MaskCodes(t) Else c = t
                ScanTyp c, base
                ScanBrackets c, base
                ScanS6 c, base
                ScanDict c, base
                ScanTokens c, base
            End If
        End If
    Next p
    FinishW2
    If mN = 0 Then
        CheckSpell = ""
        Exit Function
    End If
    SortOut
    ReDim ln(0 To mN - 1)
    For i = 0 To mN - 1
        ln(i) = mOut(i)
    Next i
    CheckSpell = Join(ln, vbLf)
End Function
