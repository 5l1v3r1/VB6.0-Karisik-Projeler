Attribute VB_Name = "Module1"
Public Function UserName() As String
On Error Resume Next
Dim oSysInfo As Object
Set oSysInfo = CreateObject("WinNTSystemInfo")
UserName = oSysInfo.UserName
Set oSysInfo = Nothing
End Function
