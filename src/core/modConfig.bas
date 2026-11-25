Attribute VB_Name = "modConfig"
Option Explicit
' Vi tri thu muc config (canh tep .dotm) va doc tep UTF-8.

Private mDir As String

Public Function ConfigDir() As String
    If Len(mDir) = 0 Then mDir = ThisDocument.Path & "\config"
    ConfigDir = mDir
End Function

Public Sub SetConfigDir(ByVal p As String)
    mDir = p
End Sub

Public Function ReadUtf8(ByVal path As String) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile path
    ReadUtf8 = st.ReadText
    st.Close
End Function
