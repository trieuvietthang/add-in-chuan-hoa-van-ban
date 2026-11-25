Attribute VB_Name = "modToc"
Option Explicit
' Muc luc = TOC field dua tren danh sach style cua ho so ([toc] entries=Style:Cap;Style:Cap) - khong go tay.
' Tieu de "MUC LUC" dung style STC_TocTitle (khong co outline -> khong roi vao muc luc).
' Chen lan 2 = cap nhat field hien co (idempotent).

Private Function TocCode(ByVal prof As Object) As String
    Dim ent() As String, kv() As String, i As Long, sep As String, lst As String
    sep = Application.International(17)
    If Len(sep) = 0 Then sep = ","
    ent = Split(modProfile.PGet(prof, "toc.entries", ""), ";")
    For i = 0 To UBound(ent)
        kv = Split(Trim$(ent(i)), ":")
        If UBound(kv) = 1 Then
            If Len(lst) > 0 Then lst = lst & sep
            lst = lst & Trim$(kv(0)) & sep & Trim$(kv(1))
        End If
    Next i
    If Len(lst) = 0 Then Err.Raise vbObjectError + 1004, "modToc", modStrings.T("err.noconfig")
    TocCode = "TOC \h \z \t """ & lst & """"
End Function

Private Function TitleStyle(ByVal doc As Document) As Style
    Dim st As Style
    On Error Resume Next
    Set st = doc.Styles("STC_TocTitle")
    On Error GoTo 0
    If st Is Nothing Then
        Set st = doc.Styles.Add("STC_TocTitle", 1)
        st.BaseStyle = doc.Styles(-1)
        st.Font = doc.Styles(-1).Font
        st.ParagraphFormat = doc.Styles(-1).ParagraphFormat
    End If
    st.Font.Bold = True
    st.ParagraphFormat.PageBreakBefore = False
    With st.ParagraphFormat
        .Alignment = 1
        .FirstLineIndent = 0
        .LeftIndent = 0
        .OutlineLevel = 10
        .SpaceAfter = 12
        .KeepWithNext = True
    End With
    Set TitleStyle = st
End Function

Private Sub FixTocStyles(ByVal doc As Document)
    Dim k As Long, st As Style, w As Single
    With doc.Sections(1).PageSetup
        w = .PageWidth - .LeftMargin - .RightMargin
    End With
    For k = 1 To 3
        Set st = doc.Styles(-19 - k)   ' wdStyleTOC1..3 = -20..-22
        With st.ParagraphFormat
            .FirstLineIndent = 0
            .LeftIndent = (k - 1) * CentimetersToPoints(0.5)
            .Alignment = 0
            .TabStops.ClearAll
            .TabStops.Add w, 2, 1
        End With
    Next k
End Sub

Private Function UpdateAll(ByVal doc As Document, ByVal code As String) As Long
    Dim f As Field, n As Long
    For Each f In doc.Fields
        If f.Type = 13 Then
            f.Code.Text = " " & code & " "
            f.Update
            n = n + 1
        End If
    Next f
    UpdateAll = n
End Function

' paraIndex: so thu tu doan se chen muc luc TRUOC no (0 = dau van ban). Tra ve so field muc luc sau khi chen.
Public Function InsertToc(ByVal doc As Document, ByVal paraIndex As Long, ByVal profileName As String) As Long
    Dim prof As Object, code As String, pos As Long, r As Range, ttl As String, en As Long, ed As String, n As Long
    Dim ps As Long
    modUndo.BeginUndo doc, "Chen muc luc"
    On Error GoTo EH
    Set prof = modProfile.LoadProfile(profileName)
    code = TocCode(prof)
    n = UpdateAll(doc, code)
    If n = 0 Then
        FixTocStyles doc
        ttl = modProfile.PGet(prof, "toc.title", "MUC LUC")
        ' Con tro trong bang (vd khoi quoc hieu): chen ngay SAU bang, khong chen vao o bang
        If paraIndex > 0 Then
            If doc.Paragraphs(paraIndex).Range.Information(12) Then
                pos = doc.Paragraphs(paraIndex).Range.Tables(1).Range.End
                paraIndex = doc.Range(0, pos).Paragraphs.Count + 1
            End If
        End If
        If paraIndex <= 0 Then pos = doc.Content.Start Else pos = doc.Paragraphs(paraIndex).Range.Start
        Set r = doc.Range(pos, pos)
        r.InsertBefore ttl & vbCr & vbCr
        doc.Paragraphs(IIf(paraIndex <= 0, 1, paraIndex)).Style = TitleStyle(doc)
        ps = IIf(paraIndex <= 0, 2, paraIndex + 1)
        doc.Paragraphs(IIf(paraIndex <= 0, 1, paraIndex)).Range.ListFormat.RemoveNumbers
        doc.Paragraphs(ps).Style = doc.Styles(-1)
        doc.Paragraphs(ps).Range.ListFormat.RemoveNumbers
        Set r = doc.Paragraphs(ps).Range
        r.Collapse 1
        doc.Fields.Add r, 13, Mid$(code, 5), False
        n = UpdateAll(doc, code)
    Else
        FixTocStyles doc
    End If
    InsertToc = n
    modUndo.EndUndo
    Exit Function
EH:
    en = Err.Number: ed = Err.Description
    modUndo.EndUndo
    Err.Raise en, "modToc", ed
End Function

' So doan thuoc cac style dua vao muc luc cua ho so (0 -> muc luc se trong)
Public Function CountSources(ByVal doc As Document, ByVal profileName As String) As Long
    Dim prof As Object, ent() As String, kv() As String, i As Long, p As Paragraph, nm As String, names As String, n As Long
    Set prof = modProfile.LoadProfile(profileName)
    ent = Split(modProfile.PGet(prof, "toc.entries", ""), ";")
    For i = 0 To UBound(ent)
        kv = Split(Trim$(ent(i)), ":")
        names = names & "|" & LCase$(Trim$(kv(0))) & "|"
    Next i
    For Each p In doc.Paragraphs
        nm = LCase$(p.Style.NameLocal)
        If InStr(names, "|" & nm & "|") > 0 Then n = n + 1
    Next p
    CountSources = n
End Function

' Cap nhat muc luc hien co (so trang + muc moi). Tra ve so muc luc da cap nhat.
Public Function UpdateToc(ByVal doc As Document) As Long
    Dim f As Field, n As Long, en As Long, ed As String
    modUndo.BeginUndo doc, "Cap nhat muc luc"
    On Error GoTo EH
    For Each f In doc.Fields
        If f.Type = 13 Then n = n + 1
    Next f
    Dim i As Long
    For i = 1 To doc.TablesOfContents.Count
        doc.TablesOfContents(i).Update
    Next i
    UpdateToc = n
    modUndo.EndUndo
    Exit Function
EH:
    en = Err.Number: ed = Err.Description
    modUndo.EndUndo
    Err.Raise en, "modToc", ed
End Function
