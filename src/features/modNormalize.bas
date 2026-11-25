Attribute VB_Name = "modNormalize"
Option Explicit
' Chuan hoa trang / phong / gian dong / so trang theo chuan (config/standards/*.ini) + ho so.
' Cung mot loi chay cho xem truoc (apply=False, khong ghi) va ap dung (apply=True): dem so thay doi.
' Khong dung Font.Reset (mat in dam, S4). Chi sua thuoc tinh sai, nen chay lai = 0 thay doi.

Private mTotal As Long
Private mOut As String

Private Sub Note(ByVal key As String, ByVal n As Long)
    If n > 0 Then
        mOut = mOut & key & "|" & CStr(n) & vbLf
        mTotal = mTotal + n
    End If
End Sub

Private Function LoadStd(ByVal stdName As String) As Object
    Dim d As Object, f As String
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    f = modConfig.ConfigDir & "\standards\" & stdName & ".ini"
    If Len(Dir$(f)) = 0 Then Err.Raise vbObjectError + 1003, "modNormalize", modStrings.T("err.noconfig")
    modProfile.ParseIni modConfig.ReadUtf8(f), d
    Set LoadStd = d
End Function

Private Function Num(ByVal std As Object, ByVal key As String, ByVal def As String) As Single
    Num = CSng(Val(modProfile.PGet(std, key, def)))
End Function

' ---- Trang ----
Private Function DoPage(ByVal doc As Document, ByVal std As Object, ByVal apply As Boolean) As Long
    Dim sec As Section, n As Long, t As Single, b As Single, l As Single, r As Single, w As Single, h As Single
    Dim ew As Single, eh As Single
    t = CentimetersToPoints(Num(std, "page.top", "2"))
    b = CentimetersToPoints(Num(std, "page.bottom", "2"))
    l = CentimetersToPoints(Num(std, "page.left", "3"))
    r = CentimetersToPoints(Num(std, "page.right", "2"))
    w = CentimetersToPoints(21)
    h = CentimetersToPoints(29.7)
    For Each sec In doc.Sections
        With sec.PageSetup
            If .Orientation = 0 Then
                ew = w: eh = h
            Else
                ew = h: eh = w
            End If
            If Abs(.PageWidth - ew) > 1 Or Abs(.PageHeight - eh) > 1 Then
                n = n + 1
                If apply Then .PaperSize = 7
            End If
            If Abs(.TopMargin - t) > 0.5 Then
                n = n + 1
                If apply Then .TopMargin = t
            End If
            If Abs(.BottomMargin - b) > 0.5 Then
                n = n + 1
                If apply Then .BottomMargin = b
            End If
            If Abs(.LeftMargin - l) > 0.5 Then
                n = n + 1
                If apply Then .LeftMargin = l
            End If
            If Abs(.RightMargin - r) > 0.5 Then
                n = n + 1
                If apply Then .RightMargin = r
            End If
        End With
    Next sec
    DoPage = n
End Function

' ---- Style: fiPts = -1 -> khong dong den thut dau dong (do danh so quan ly) ----
Private Function FixStyle(ByVal st As Style, ByVal std As Object, ByVal apply As Boolean, ByVal fiPts As Single) As Long
    Dim n As Long, fn As String, fs As Single, ls As Single, sa As Single, sb As Single
    fn = modProfile.PGet(std, "base.font", "Times New Roman")
    fs = Num(std, "base.size", "13")
    ls = LinesToPoints(Num(std, "base.linespacing", "1.15"))
    sa = Num(std, "base.spaceafter", "6")
    sb = Num(std, "base.spacebefore", "0")
    With st
        If .Font.Name <> fn Then
            n = n + 1
            If apply Then .Font.Name = fn
        End If
        If Abs(.Font.Size - fs) > 0.01 Then
            n = n + 1
            If apply Then .Font.Size = fs
        End If
        With .ParagraphFormat
            If .LineSpacingRule <> 5 Or Abs(.LineSpacing - ls) > 0.1 Then
                n = n + 1
                If apply Then
                    .LineSpacingRule = 5
                    .LineSpacing = ls
                End If
            End If
            If Abs(.SpaceAfter - sa) > 0.1 Then
                n = n + 1
                If apply Then .SpaceAfter = sa
            End If
            If Abs(.SpaceBefore - sb) > 0.1 Then
                n = n + 1
                If apply Then .SpaceBefore = sb
            End If
            If fiPts >= 0 Then
                If Abs(.FirstLineIndent - fiPts) > 0.5 Then
                    n = n + 1
                    If apply Then .FirstLineIndent = fiPts
                End If
            End If
        End With
    End With
    FixStyle = n
End Function

Private Function DoStyles(ByVal doc As Document, ByVal prof As Object, ByVal std As Object, ByVal apply As Boolean) As Long
    Dim n As Long, i As Long, st As Style, fi As Single
    fi = CentimetersToPoints(Num(std, "base.firstindent", "0"))
    n = n + FixStyle(doc.Styles(-1), std, apply, fi)        ' Normal
    n = n + FixStyle(doc.Styles(-32), std, apply, 0)        ' Header
    n = n + FixStyle(doc.Styles(-33), std, apply, 0)        ' Footer
    For i = 1 To modProfile.LevelCount(prof)
        Set st = Nothing
        On Error Resume Next
        Set st = doc.Styles(modProfile.LevelStyle(prof, i))
        On Error GoTo 0
        If Not st Is Nothing Then
            n = n + FixStyle(st, std, apply, -1)
            If st.ParagraphFormat.PageBreakBefore Then
                n = n + 1
                If apply Then st.ParagraphFormat.PageBreakBefore = False
            End If
        End If
    Next i
    Set st = Nothing
    On Error Resume Next
    Set st = doc.Styles("STC_TocTitle")
    On Error GoTo 0
    If Not st Is Nothing Then
        If st.ParagraphFormat.PageBreakBefore Then
            n = n + 1
            If apply Then st.ParagraphFormat.PageBreakBefore = False
        End If
    End If
    DoStyles = n
End Function

' ---- Tung doan: font/co sai (ke ca run) va gian dong truc tiep; bang chi sua font, bo thut dau dong ----
Private Sub DoParas(ByVal doc As Document, ByVal std As Object, ByVal apply As Boolean, ByRef nFont As Long, ByRef nSize As Long, ByRef nSpace As Long)
    Dim p As Paragraph, r As Range, fn As String, fs As Single, ls As Single, sa As Single, sb As Single, inTbl As Boolean
    fn = modProfile.PGet(std, "base.font", "Times New Roman")
    fs = Num(std, "base.size", "13")
    ls = LinesToPoints(Num(std, "base.linespacing", "1.15"))
    sa = Num(std, "base.spaceafter", "6")
    sb = Num(std, "base.spacebefore", "0")
    For Each p In doc.Paragraphs
        Set r = p.Range
        inTbl = r.Information(12)
        If r.Font.Name <> fn Then
            nFont = nFont + 1
            If apply Then r.Font.Name = fn
        End If
        If Abs(r.Font.Size - fs) > 0.01 Then
            nSize = nSize + 1
            If apply Then r.Font.Size = fs
        End If
        If inTbl Then
            If Abs(p.FirstLineIndent) > 0.5 Then
                nSpace = nSpace + 1
                If apply Then p.FirstLineIndent = 0
            End If
        Else
            If p.LineSpacingRule <> 5 Or Abs(p.LineSpacing - ls) > 0.1 Or Abs(p.SpaceAfter - sa) > 0.1 Or Abs(p.SpaceBefore - sb) > 0.1 Then
                nSpace = nSpace + 1
                If apply Then
                    p.LineSpacingRule = 5
                    p.LineSpacing = ls
                    p.SpaceAfter = sa
                    p.SpaceBefore = sb
                End If
            End If
        End If
    Next p
End Sub

' ---- So trang (field PAGE) o chan/dau trang; khong ghi de noi dung co san ----
Private Function HasPageField(ByVal rng As Range) As Boolean
    Dim f As Field
    For Each f In rng.Fields
        If f.Type = 33 Then HasPageField = True
    Next f
End Function

Private Function DoPageNumber(ByVal doc As Document, ByVal std As Object, ByVal apply As Boolean) As Long
    Dim sec As Section, hf As HeaderFooter, where As String, first As Boolean, n As Long, r As Range, want As Boolean
    where = LCase$(modProfile.PGet(std, "pagenumber.where", "footer"))
    If where = "none" Then Exit Function
    first = (Val(modProfile.PGet(std, "pagenumber.firstpage", "0")) <> 0)
    want = Not first   ' DifferentFirstPage = True -> trang 1 khong co so
    For Each sec In doc.Sections
        If sec.PageSetup.DifferentFirstPageHeaderFooter <> IIf(want, -1, 0) Then
            n = n + 1
            If apply Then sec.PageSetup.DifferentFirstPageHeaderFooter = IIf(want, -1, 0)
        End If
        If where = "header" Then Set hf = sec.Headers(1) Else Set hf = sec.Footers(1)
        If Not (sec.Index > 1 And hf.LinkToPrevious) Then
            If Not HasPageField(hf.Range) Then
                n = n + 1
                If apply Then
                    If Len(hf.Range.Text) > 1 Then hf.Range.InsertParagraphAfter
                    Set r = hf.Range.Paragraphs(hf.Range.Paragraphs.Count).Range
                    r.Collapse 1
                    doc.Fields.Add r, 33
                    hf.Range.Paragraphs(hf.Range.Paragraphs.Count).Alignment = 1
                End If
            End If
        End If
    Next sec
    DoPageNumber = n
End Function

Private Function Core(ByVal doc As Document, ByVal profileName As String, ByVal stdName As String, ByVal apply As Boolean) As String
    Dim prof As Object, std As Object, nF As Long, nS As Long, nP As Long
    mOut = "": mTotal = 0
    Set prof = modProfile.LoadProfile(profileName)
    If Len(stdName) = 0 Then stdName = modProfile.PGet(prof, "profile.standard", "ND30")
    Set std = LoadStd(stdName)
    Note "norm.page", DoPage(doc, std, apply)
    Note "norm.style", DoStyles(doc, prof, std, apply)
    DoParas doc, std, apply, nF, nS, nP
    Note "norm.font", nF
    Note "norm.size", nS
    Note "norm.space", nP
    Note "norm.footer", DoPageNumber(doc, std, apply)
    Core = mOut
End Function

' Tra ve cac dong "khoa|so luong" (khoa tra cuu qua modStrings.T); chuoi rong = khong co gi can sua
Public Function PreviewStandard(ByVal doc As Document, ByVal profileName As String, ByVal stdName As String) As String
    PreviewStandard = Core(doc, profileName, stdName, False)
End Function

' Tra ve tong so thay doi; mot UndoRecord
Public Function ApplyStandard(ByVal doc As Document, ByVal profileName As String, ByVal stdName As String) As Long
    Dim en As Long, ed As String
    modUndo.BeginUndo doc, "Chuan hoa trang va phong"
    On Error GoTo EH
    Core doc, profileName, stdName, True
    ApplyStandard = mTotal
    modUndo.EndUndo
    Exit Function
EH:
    en = Err.Number: ed = Err.Description
    modUndo.EndUndo
    Err.Raise en, "modNormalize", ed
End Function
