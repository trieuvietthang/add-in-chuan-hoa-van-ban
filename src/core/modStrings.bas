Attribute VB_Name = "modStrings"
Option Explicit
' Chuoi tieng Viet doc tu config\messages.vi.ini (UTF-8). Tep VBA chi chua ASCII.

Private mMsg As Object

Public Function Hello() As String
    Hello = "Xin ch" & ChrW(224) & "o"
End Function

Public Function T(ByVal key As String) As String
    If mMsg Is Nothing Then
        Set mMsg = CreateObject("Scripting.Dictionary")
        mMsg.CompareMode = 1
        On Error Resume Next
        modProfile.ParseIni modConfig.ReadUtf8(modConfig.ConfigDir & "\messages.vi.ini"), mMsg
        On Error GoTo 0
    End If
    If mMsg.Exists("msg." & key) Then T = mMsg("msg." & key) Else T = key
End Function
