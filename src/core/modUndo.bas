Attribute VB_Name = "modUndo"
Option Explicit
' Moi lenh ghi vao van ban nam trong MOT UndoRecord (Ctrl+Z hoan tac 1 lan).
' UndoRecord gan voi tai lieu dang active, nen kich hoat tai lieu dich truoc khi bat dau
' (ngoai thuc te tai lieu dang soan da la active nen khong doi gi).

Public Sub BeginUndo(ByVal doc As Document, ByVal title As String)
    On Error Resume Next
    If Not (ActiveDocument Is doc) Then doc.Activate
    Application.UndoRecord.StartCustomRecord title
    On Error GoTo 0
End Sub

Public Sub EndUndo()
    On Error Resume Next
    If Application.UndoRecord.IsRecordingCustomRecord Then Application.UndoRecord.EndCustomRecord
    On Error GoTo 0
End Sub
