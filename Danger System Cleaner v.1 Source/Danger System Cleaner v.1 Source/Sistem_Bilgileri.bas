Attribute VB_Name = "Sistem_Bilgileri"
'BELLEK DURUMU
Public Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Public Type MEMORYSTATUS
        dwLength As Long
        dwMemoryLoad As Long
        dwTotalPhys As Long
        dwAvailPhys As Long
        dwTotalPageFile As Long
        dwAvailPageFile As Long
        dwTotalVirtual As Long
        dwAvailVirtual As Long
End Type
Public Function UserName() As String
'On Error Resume Next
Dim oSysInfo As Object
Set oSysInfo = CreateObject("WinNTSystemInfo")
UserName = oSysInfo.UserName
Set oSysInfo = Nothing
End Function

Public Function ReadKey(Value As String) As String

Dim b As Object
On Error Resume Next
Set b = CreateObject("wscript.shell")
r = b.RegRead(Value)
ReadKey = r
End Function


