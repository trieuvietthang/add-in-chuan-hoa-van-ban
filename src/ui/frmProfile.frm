VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProfile
   Caption         =   "frmProfile"
   ClientHeight    =   8800
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10400
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Form "Sua ho so": bo soan INI tong quat (chuan the thuc, ho so danh so, nguong chinh ta). Cac o nhap duoc dung luc chay (Controls.Add).

Private WithEvents cboFile As MSForms.ComboBox
Private WithEvents cboSec As MSForms.ComboBox
Private WithEvents btnSave As MSForms.CommandButton
Private WithEvents btnClose As MSForms.CommandButton
Private fra As MSForms.Frame
Private lblStatus As MSForms.Label
Private mFiles() As String
Private mWhite() As String
Private mFileN As Long
Private mCurFile As Long
Private mTxt As String
Private mSec As String
Private mRowN As Long
Private mRowVal() As String
Private mPend As Object
Private mWarned As Boolean

Private Function PathOf(ByVal idx As Long) As String
    PathOf = modConfig.ConfigDir & "\" & mFiles(idx) & ".ini"
End Function

Private Sub SetStatus(ByVal s As String)
    lblStatus.Caption = s
End Sub

Private Sub UserForm_Initialize()
    Dim l As MSForms.Label, recs() As String, f() As String, i As Long, w As Single, h As Single
    modIniEdit.ResetUi
    Me.Caption = modIniEdit.Ui("text.title")
    Set mPend = CreateObject("Scripting.Dictionary")
    mPend.CompareMode = 1
    w = Me.InsideWidth: h = Me.InsideHeight
    Set l = Me.Controls.Add("Forms.Label.1", "lblIntro", True)
    l.Left = 12: l.Top = 6: l.Width = w - 24: l.Height = 30: l.WordWrap = True
    l.Caption = modIniEdit.Ui("text.intro")
    Set l = Me.Controls.Add("Forms.Label.1", "lblFile", True)
    l.Left = 12: l.Top = 46: l.Width = 40: l.Height = 16
    l.Caption = modIniEdit.Ui("text.file")
    Set cboFile = Me.Controls.Add("Forms.ComboBox.1", "cboFile", True)
    cboFile.Left = 54: cboFile.Top = 44: cboFile.Width = 200: cboFile.Height = 20: cboFile.Style = 2
    Set l = Me.Controls.Add("Forms.Label.1", "lblSec", True)
    l.Left = 270: l.Top = 46: l.Width = 40: l.Height = 16
    l.Caption = modIniEdit.Ui("text.section")
    Set cboSec = Me.Controls.Add("Forms.ComboBox.1", "cboSec", True)
    cboSec.Left = 312: cboSec.Top = 44: cboSec.Width = 160: cboSec.Height = 20: cboSec.Style = 2
    Set fra = Me.Controls.Add("Forms.Frame.1", "fra", True)
    fra.Left = 12: fra.Top = 74: fra.Width = w - 24: fra.Height = h - 74 - 44
    fra.Caption = "": fra.ScrollBars = 2
    Set btnSave = Me.Controls.Add("Forms.CommandButton.1", "btnSave", True)
    btnSave.Left = w - 190: btnSave.Top = h - 32: btnSave.Width = 84: btnSave.Height = 24
    btnSave.Caption = modIniEdit.Ui("text.save")
    Set btnClose = Me.Controls.Add("Forms.CommandButton.1", "btnClose", True)
    btnClose.Left = w - 96: btnClose.Top = h - 32: btnClose.Width = 84: btnClose.Height = 24
    btnClose.Caption = modIniEdit.Ui("text.close")
    Set lblStatus = Me.Controls.Add("Forms.Label.1", "lblStatus", True)
    lblStatus.Left = 12: lblStatus.Top = h - 34: lblStatus.Width = w - 210: lblStatus.Height = 28: lblStatus.WordWrap = True
    recs = Split(modIniEdit.Ui("files.list"), ";")
    mFileN = UBound(recs) + 1
    ReDim mFiles(0 To mFileN - 1): ReDim mWhite(0 To mFileN - 1)
    For i = 0 To mFileN - 1
        f = Split(recs(i), "|")
        mFiles(i) = f(0)
        If UBound(f) >= 2 Then
            If Len(Trim$(f(2))) > 0 Then mWhite(i) = "," & LCase$(Replace(f(2), " ", "")) & ","
        End If
        cboFile.AddItem f(1)
    Next i
    mCurFile = -1
    cboFile.ListIndex = 0
    If mCurFile <> 0 Then LoadFile 0      ' su kien Change co the chua chay trong Initialize
    If mRowN = 0 And cboSec.ListCount > 0 Then ShowSection
End Sub

Private Sub ClearRows()
    Dim i As Long
    For i = 0 To mRowN - 1
        fra.Controls.Remove "lbl" & CStr(i)
        fra.Controls.Remove "txt" & CStr(i)
    Next i
    mRowN = 0
End Sub

Private Sub CapturePending()
    Dim i As Long, t As MSForms.TextBox, pk As String
    For i = 0 To mRowN - 1
        Set t = fra.Controls("txt" & CStr(i))
        If t.Enabled Then
            pk = CStr(mCurFile) & "|" & mSec & "|" & CStr(t.Tag)
            If t.Text <> mRowVal(i) Then
                mPend(pk) = t.Text
            ElseIf mPend.Exists(pk) Then
                mPend.Remove pk
            End If
        End If
    Next i
End Sub

Private Sub LoadFile(ByVal idx As Long)
    Dim names() As String, i As Long, n As String
    mCurFile = idx
    mTxt = modIniEdit.ReadText(PathOf(idx))
    ClearRows
    cboSec.Clear
    names = Split(modIniEdit.SectionNames(mTxt), vbLf)
    For i = 0 To UBound(names)
        n = names(i)
        If Len(n) > 0 Then
            If Len(mWhite(idx)) = 0 Or InStr(mWhite(idx), "," & LCase$(n) & ",") > 0 Then cboSec.AddItem n
        End If
    Next i
    If cboSec.ListCount > 0 Then cboSec.ListIndex = 0
End Sub

Private Sub ShowSection()
    Dim items() As String, kv() As String, i As Long, l As MSForms.Label, t As MSForms.TextBox, pk As String, v As String
    ClearRows
    If cboSec.ListIndex < 0 Then Exit Sub
    mSec = cboSec.List(cboSec.ListIndex)
    If Len(modIniEdit.SectionItems(mTxt, mSec)) = 0 Then Exit Sub
    items = Split(modIniEdit.SectionItems(mTxt, mSec), vbLf)
    ReDim mRowVal(0 To UBound(items))
    For i = 0 To UBound(items)
        kv = Split(items(i), vbTab)
        v = ""
        If UBound(kv) >= 1 Then v = kv(1)
        mRowVal(i) = v
        pk = CStr(mCurFile) & "|" & mSec & "|" & kv(0)
        Set l = fra.Controls.Add("Forms.Label.1", "lbl" & CStr(i), True)
        l.Left = 6: l.Top = 8 + i * 24: l.Width = 230: l.Height = 18
        l.Caption = modIniEdit.Hint(mSec, kv(0))
        Set t = fra.Controls.Add("Forms.TextBox.1", "txt" & CStr(i), True)
        t.Left = 240: t.Top = 6 + i * 24: t.Width = fra.InsideWidth - 258: t.Height = 20
        t.Tag = kv(0)
        t.ControlTipText = kv(0)
        If mPend.Exists(pk) Then t.Text = mPend(pk) Else t.Text = v
        If modIniEdit.IsReadOnly(mSec, kv(0)) Then t.Enabled = False
    Next i
    mRowN = UBound(items) + 1
    fra.ScrollHeight = mRowN * 24 + 16
    fra.ScrollTop = 0
End Sub

Private Sub cboFile_Change()
    If cboFile.ListIndex < 0 Then Exit Sub
    If cboFile.ListIndex = mCurFile Then Exit Sub
    CapturePending
    LoadFile cboFile.ListIndex
End Sub

Private Sub cboSec_Change()
    If mRowN > 0 Then CapturePending
    ShowSection
End Sub

Private Sub btnSave_Click()
    Dim k As Variant, p() As String, msg As String, fi As Long, ed As Object, total As Long
    On Error GoTo EH
    CapturePending
    mWarned = False
    If mPend.Count = 0 Then SetStatus modIniEdit.Ui("text.nochange"): Exit Sub
    For Each k In mPend.Keys
        p = Split(CStr(k), "|")
        msg = modIniEdit.Validate(p(1), p(2), CStr(mPend(k)))
        If Len(msg) > 0 Then
            SetStatus Replace(Replace(modIniEdit.Ui("text.badvalue"), "{0}", modIniEdit.Hint(p(1), p(2))), "{1}", msg)
            Exit Sub
        End If
    Next k
    For fi = 0 To mFileN - 1
        Set ed = CreateObject("Scripting.Dictionary")
        ed.CompareMode = 1
        For Each k In mPend.Keys
            p = Split(CStr(k), "|")
            If CLng(p(0)) = fi Then ed(LCase$(p(1)) & "|" & LCase$(p(2))) = CStr(mPend(k))
        Next k
        If ed.Count > 0 Then total = total + modIniEdit.SaveEdits(PathOf(fi), ed)
    Next fi
    mPend.RemoveAll
    modSpell.ResetConfig
    mTxt = modIniEdit.ReadText(PathOf(mCurFile))
    ShowSection
    SetStatus Replace(modIniEdit.Ui("text.saved"), "{0}", CStr(total))
    Exit Sub
EH:
    SetStatus Err.Description
End Sub

Private Sub btnClose_Click()
    CapturePending
    If mPend.Count > 0 And Not mWarned Then
        mWarned = True
        SetStatus modIniEdit.Ui("text.pending")
        Exit Sub
    End If
    Me.Hide
End Sub

' ---- ham ho tro kiem thu (khong dung trong giao dien) ----
Public Function Probe() As String
    Probe = CStr(mFileN) & "|" & CStr(cboSec.ListCount) & "|" & CStr(mRowN)
End Function

Public Sub Pick(ByVal fileIdx As Long, ByVal sec As String)
    Dim i As Long
    cboFile.ListIndex = fileIdx
    For i = 0 To cboSec.ListCount - 1
        If StrComp(cboSec.List(i), sec, vbTextCompare) = 0 Then cboSec.ListIndex = i
    Next i
End Sub

Public Function TestEdit(ByVal key As String, ByVal value As String) As String
    Dim i As Long
    For i = 0 To mRowN - 1
        If StrComp(CStr(fra.Controls("txt" & CStr(i)).Tag), key, vbTextCompare) = 0 Then fra.Controls("txt" & CStr(i)).Text = value
    Next i
    btnSave_Click
    TestEdit = lblStatus.Caption
End Function
