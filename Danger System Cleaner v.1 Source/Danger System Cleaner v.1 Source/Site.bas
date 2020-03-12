Attribute VB_Name = "Site"
Public Function Siteler() As String
On Error Resume Next
Shell "C:\Program Files\Internet Explorer\IEXPLORE.EXE cyber-warrior.org"
Shell "C:\Documents and Settings\" & Form1.Text2.Text & "\Local Settings\Application Data\Google\Chrome\Application\chrome.exe cyber-warrior.org"
Shell "C:\Program Files\Mozilla Firefox\firefox.exe cyber-warrior.org"
Shell "D:\Program Files\Internet Explorer\IEXPLORE.EXE cyber-warrior.org"
Shell "D:\Documents and Settings\" & Form1.Text2.Text & "\Local Settings\Application Data\Google\Chrome\Application\chrome.exe cyber-warrior.org"
Shell "D:\Program Files\Mozilla Firefox\firefox.exe cyber-warrior.org"
End Function
