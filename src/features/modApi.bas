Attribute VB_Name = "modApi"
Option Explicit
' API mong: tham so chi gom chuoi/so (Application.Run khong truyen duoc doi tuong tu ben ngoai).
' docName = ten tai lieu (vd "Document1"); first/last = chi so doan (1-based); first = 0 -> ca van ban.

Private Function RangeOf(ByVal docName As String, ByVal first As Long, ByVal last As Long) As Range
    Dim doc As Document
    Set doc = Documents(docName)
    If first <= 0 Then
        Set RangeOf = doc.Content
    Else
        If last < first Then last = first
        Set RangeOf = doc.Range(doc.Paragraphs(first).Range.Start, doc.Paragraphs(last).Range.End)
    End If
End Function

Public Sub ApplyLevelParas(ByVal docName As String, ByVal first As Long, ByVal last As Long, ByVal level As Long, ByVal profileName As String)
    modApply.ApplyLevelToRange RangeOf(docName, first, last), level, profileName
End Sub

Public Sub ChangeLevelParas(ByVal docName As String, ByVal first As Long, ByVal last As Long, ByVal delta As Long, ByVal profileName As String)
    modApply.ChangeLevelRange RangeOf(docName, first, last), delta, profileName
End Sub

Public Function PreviewTypedParas(ByVal docName As String, ByVal first As Long, ByVal last As Long, ByVal profileName As String) As String
    PreviewTypedParas = modConvert.PreviewTyped(RangeOf(docName, first, last), profileName)
End Function

Public Function ConvertTypedParas(ByVal docName As String, ByVal first As Long, ByVal last As Long, ByVal profileName As String) As Long
    ConvertTypedParas = modConvert.ConvertTyped(RangeOf(docName, first, last), profileName)
End Function

Public Function PreviewNormalize(ByVal docName As String, ByVal profileName As String, ByVal stdName As String) As String
    PreviewNormalize = modNormalize.PreviewStandard(Documents(docName), profileName, stdName)
End Function

Public Function NormalizeDoc(ByVal docName As String, ByVal profileName As String, ByVal stdName As String) As Long
    NormalizeDoc = modNormalize.ApplyStandard(Documents(docName), profileName, stdName)
End Function

Public Function InsertTocAt(ByVal docName As String, ByVal paraIndex As Long, ByVal profileName As String) As Long
    InsertTocAt = modToc.InsertToc(Documents(docName), paraIndex, profileName)
End Function

Public Function CountTocSources(ByVal docName As String, ByVal profileName As String) As Long
    CountTocSources = modToc.CountSources(Documents(docName), profileName)
End Function

Public Function UpdateTocDoc(ByVal docName As String) As Long
    UpdateTocDoc = modToc.UpdateToc(Documents(docName))
End Function

Public Function CheckSpellDoc(ByVal docName As String, ByVal profileName As String) As String
    CheckSpellDoc = modSpell.CheckSpell(Documents(docName), profileName)
End Function

Public Function SyllableOk(ByVal w As String) As String
    If modSpell.SyllableOk(w) Then SyllableOk = "1" Else SyllableOk = "0"
End Function

' Kiem thu form "Sua ho so": nap form (khong hien), chon tep/muc, sua mot o, bam Luu, tra ve "probe|trang_thai"
Public Function FormTest(ByVal fileIdx As Long, ByVal sec As String, ByVal key As String, ByVal value As String) As String
    Dim pr As String, st As String
    Load frmProfile
    frmProfile.Pick fileIdx, sec
    pr = frmProfile.Probe()
    If Len(key) > 0 Then st = frmProfile.TestEdit(key, value)
    Unload frmProfile
    FormTest = pr & "|" & st
End Function

Public Function IniValidate(ByVal sec As String, ByVal key As String, ByVal value As String) As String
    IniValidate = modIniEdit.Validate(sec, key, value)
End Function

Public Function IniSave(ByVal path As String, ByVal sec As String, ByVal key As String, ByVal value As String) As String
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d(LCase$(sec) & "|" & LCase$(key)) = value
    IniSave = CStr(modIniEdit.SaveEdits(path, d))
End Function

Public Function CheckRefsDoc(ByVal docName As String, ByVal profileName As String) As String
    CheckRefsDoc = modRefs.CheckRefs(Documents(docName), profileName)
End Function

Public Function ConvertRefsDoc(ByVal docName As String, ByVal profileName As String) As Long
    ConvertRefsDoc = modRefs.ConvertRefs(Documents(docName), profileName)
End Function

Public Function UpdateRefsDoc(ByVal docName As String) As Long
    UpdateRefsDoc = modRefs.UpdateRefs(Documents(docName))
End Function

Public Function InsertRefDoc(ByVal docName As String, ByVal pos As Long, ByVal dieuNo As Long, ByVal profileName As String) As String
    InsertRefDoc = modRefs.InsertRef(Documents(docName), pos, dieuNo, profileName)
End Function

' Chan doan: chay tung buoc, tra ve "ok ..." hoac "ERR so: mo ta" (dung khi lenh bi bao "Can't run the specified macro")
Public Function Diag(ByVal docName As String, ByVal profileName As String) As String
    Dim prof As Object, doc As Document, stg As String
    On Error GoTo EH
    stg = "doc": Set doc = Documents(docName)
    stg = "load": Set prof = modProfile.LoadProfile(profileName)
    stg = "ensure": modLists.EnsureProfile doc, prof
    stg = "apply": modApply.ApplyLevelToRange doc.Paragraphs(1).Range, 1, profileName
    Diag = "ok levels=" & CStr(modProfile.LevelCount(prof))
    Exit Function
EH:
    Diag = "ERR [" & stg & "] " & CStr(Err.Number) & ": " & Err.Description
End Function
