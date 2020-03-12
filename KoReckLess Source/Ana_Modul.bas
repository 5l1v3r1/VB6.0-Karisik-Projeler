Attribute VB_Name = "Ana_Modul"

Public Type CHAR_INFORMATION
NEAR As Long
MEID As Long
TID As Long
NT As Long
MAXMP As Long
MP As Long
MaxHP As Long
HP As Long
CLASS As Long
LVL As Long
GOLD As Long
EXP As Long
MAXEXP As Long
ZONE As Long
X As Long
Y As Long
Z As Long
End Type


Public Type PARTY_VAULE
ID As Long
LVL As Long
RACE As Long
HP As Long
MaxHP As Long
MP As Long
MAXMP As Long
End Type


Public Type PARTY_INFORMATION
COUNT As Long
M(8) As PARTY_VAULE
End Type

Public Const MAXINV_ARRAY = 41

Public Type INV_VAULE
ID As Long
EXT As Long
End Type

Public Type INV_INFORMATION
SLOT(MAXINV_ARRAY) As INV_VAULE
End Type
Public Declare Function ReadProcessMem Lib "kernel32" Alias "ReadProcessMemory" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByRef lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public KOHandle As Long
'Priheal datediff
Public PriHealTimer As Date
Public PriHealDiff As Long
Public Declare Function Loot Lib "DLL.dll" (ByVal SLOT As String, ByVal Enable As Boolean) As Boolean
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Private Declare Function SetCordinate Lib "DLL.dll" (ByVal SLOT As String, ByVal X As Long, ByVal Y As Long) As Boolean
Public Declare Function SendPacket Lib "DLL.dll" (ByVal SLOT As String, ByRef lpBuffer As Any, ByVal nSize As Long) As Boolean
Public Declare Function CharInfo Lib "DLL.dll" () As CHAR_INFORMATION
Public Cinfo As CHAR_INFORMATION
Public Pinfo As PARTY_INFORMATION
Public Iinfo As INV_INFORMATION

Public Type MMOVE
X As Long
Y As Long
End Type


Public Function DDA(XA As Long, YA As Long, XB As Long, YB As Long, step As Long) As MMOVE
Dim ret As MMOVE
On Error Resume Next
Dim dx, dy, steps, DisX, DisY As Long
dx = XB - XA
dy = YB - YA
Dim xIncrement, yIncrement  As Long
ret.X = XA
ret.Y = YA

If (Abs(dx) > Abs(dy)) Then
steps = Abs(dx)
Else
steps = Abs(dy)
End If

xIncrement = dx / steps
yIncrement = dy / steps

For i = 0 To step

  ret.X = ret.X + xIncrement
  ret.Y = ret.Y + yIncrement

Next

DDA = ret
'powered by tarja
End Function


Function GetDistance(X As Long, Y As Long, X2 As Long, Y2 As Long) As Long
GetDistance = ((X - X2) ^ 2 + (Y - Y2) ^ 2) ^ 0.5
End Function
Public Function CMove(X As Long, Y As Long, steps As Long) As Long

Cinfo = CharInfo()

Dim ret As MMOVE
ret = DDA(Cinfo.X, Cinfo.Y, X, Y, steps)

Run ret.X, ret.Y, 1
Sleep 250
Do While GetDistance(Cinfo.X, Cinfo.Y, X, Y) > 1
Cinfo = CharInfo()
ret = DDA(Cinfo.X, Cinfo.Y, X, Y, steps)
Run ret.X, ret.Y, 3
Sleep 250
Loop
Run X, Y, 0



End Function


Public Function Run(X As Long, Y As Long, RTYPE As Long)
PacketSend "06" & Mid(AlignDWORD(X * 10), 1, 4) & Mid(AlignDWORD(Y * 10), 1, 4) & "2F002D00" & Mid(AlignDWORD(RTYPE), 1, 2) & Mid(AlignDWORD(X * 10), 1, 4) & Mid(AlignDWORD(Y * 10), 1, 4) & "2F00"
Call SetCordinate("sample_mailslot", X, Y)
End Function


Public Function ReadLong(Addr As Long) As Long 'read a 4 byte value
    Dim Value As Long
    If KOHandle <> 0 Then
    ReadProcessMem KOHandle, Addr, Value, 4, 0&
    End If
    ReadLong = Value
End Function



Public Function ConvHEX2ByteArray(pStr As String, pByte() As Byte)
On Error Resume Next
Dim i As Long
Dim j As Long
ReDim pByte(1 To Len(pStr) / 2)

j = LBound(pByte) - 1
For i = 1 To Len(pStr) Step 2
    j = j + 1
    pByte(j) = CByte("&H" & Mid(pStr, i, 2))
Next
End Function
Public Function Send(Pack() As Byte, pSize As Long)
    SendPacket "sample_mailslot", Pack(LBound(Pack)), pSize
End Function
Public Function PacketSend(Packet As String)
Dim pBytes() As Byte
ConvHEX2ByteArray Packet, pBytes
Send pBytes, UBound(pBytes) - LBound(pBytes) + 1
End Function


Function Percent(Vaule As Long, ByVal VaulePercent As Long) As Long
Percent = Round(Vaule * VaulePercent / 100, 0)
End Function

Public Function SetHealSkills(point As Long) As Long
Dim SkillPoint As Long
SkillPoint = point
If SkillPoint > 0 And SkillPoint < 9 Then
SetHealSkills = 1
End If
If SkillPoint > 9 And SkillPoint < 18 Then
SetHealSkills = 2
End If
If SkillPoint > 18 And SkillPoint < 27 Then
SetHealSkills = 3
End If
If SkillPoint > 27 And SkillPoint < 36 Then
SetHealSkills = 4
End If
If SkillPoint > 36 And SkillPoint < 45 Then
SetHealSkills = 5
End If
If SkillPoint > 45 Then
SetHealSkills = 6
End If
End Function

'HEAL
Public Function PriHeal(HealID As Long, target As Long)



PriHealDiff = DateDiff("s", PriHealTimer, Now)
If PriHealDiff >= 1 Then

Cinfo = CharInfo()

PacketSend "3101" & AlignDWORD(HealID) & Mid(AlignDWORD(Cinfo.MEID), 1, 4) & Mid(AlignDWORD(target), 1, 4) & "00000000000000000000000000000F00"
PacketSend "3103" & AlignDWORD(HealID) & Mid(AlignDWORD(Cinfo.MEID), 1, 4) & Mid(AlignDWORD(target), 1, 4) & "0000000000000000000000000000"
PriHealTimer = Now
End If

End Function

Function PriHeal2(Totalhp As Long, Curhp As Long, Targetid As Long)
On Error Resume Next
Dim healhp, CharClass As Long
healhp = Totalhp - Curhp
Cinfo = CharInfo()
CharClass = Cinfo.CLASS
If healhp <> "0" Then

Select Case SetHealSkills(Form2.Text5.Text)

Case 1
'60 heal
If healhp < Totalhp Then
If healhp > 20 Then
PriHeal CharClass & "500", Targetid
End If
End If

'>>>>>>>>>>>
Case 2

'60 heal
If healhp < 60 Then
If healhp > 10 Then
PriHeal CharClass & "500", Targetid
End If
End If

'240 heal
If healhp < Totalhp Then
If healhp > 60 Then
PriHeal CharClass & "509", Targetid
End If
End If


'>>>>>>>>>>>>>>>
Case 3

'60 heal
If healhp < 60 Then
If healhp > 20 Then
PriHeal CharClass & "500", Targetid
End If
End If

'240 heal
If healhp < 240 Then
If healhp > 60 Then
PriHeal CharClass & "509", Targetid
End If
End If


'360 heal
If healhp < Totalhp Then
If healhp > 240 Then
PriHeal CharClass & "518", Targetid
End If
End If


'>>>>>>>>>>>>>>
Case 4

'60 heal
If healhp < 60 Then
If healhp > 20 Then
PriHeal CharClass & "500", Targetid
End If
End If

'240 heal
If healhp < 240 Then
If healhp > 60 Then
PriHeal CharClass & "509", Targetid
End If
End If


'360 heal
If healhp < 360 Then
If healhp > 240 Then
PriHeal CharClass & "518", Targetid
End If
End If

'720 heal
If healhp < Totalhp Then
If healhp > 360 Then
PriHeal CharClass & "527", Targetid
End If
End If




'>>>>>>>>>>>>>>>>>>>>
Case 5
'60 heal
If healhp < 60 Then
If healhp > 20 Then
PriHeal CharClass & "500", Targetid
End If
End If

'240 heal
If healhp < 240 Then
If healhp > 60 Then
PriHeal CharClass & "509", Targetid
End If
End If


'360 heal
If healhp < 360 Then
If healhp > 240 Then
PriHeal CharClass & "518", Targetid
End If
End If

'720 heal
If healhp < 720 Then
If healhp > 360 Then
PriHeal CharClass & "527", Targetid
End If
End If

'960 heal
If healhp < Totalhp Then
If healhp > 720 Then
PriHeal CharClass & "536", Targetid
End If
End If


'>>>>>>>>>>
Case 6
'60 heal
If healhp < 60 Then
If healhp > 20 Then
PriHeal CharClass & "500", Targetid
End If
End If

'240 heal
If healhp < 240 Then
If healhp > 60 Then
PriHeal CharClass & "509", Targetid
End If
End If


'360 heal
If healhp < 360 Then
If healhp > 240 Then
PriHeal CharClass & "518", Targetid
End If
End If

'720 heal
If healhp < 720 Then
If healhp > 360 Then
PriHeal CharClass & "527", Targetid
End If
End If

'960 heal
If healhp < 960 Then
If healhp > 720 Then
PriHeal CharClass & "536", Targetid
End If
End If


'1920 heal
If healhp < Totalhp Then
If healhp > 960 Then
PriHeal CharClass & "545", Targetid
End If
End If


End Select
End If

End Function

'Function created by cecil

Public Function KarakterID()
Cinfo = CharInfo()
KarakterID = Mid(AlignDWORD(Cinfo.MEID), 1, 4)
End Function

Public Function PotBas(PotID)
PacketSend "3103" & PotID & "7A0700" & KarakterID & KarakterID
End Function
