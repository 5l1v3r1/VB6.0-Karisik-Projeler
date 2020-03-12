VERSION 5.00
Begin VB.UserControl Program_Yakalayýcý_Kapatýcý 
   ClientHeight    =   4425
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7110
   MaskPicture     =   "Program_Yakalayýcý_Kapatýcý.ctx":0000
   ScaleHeight     =   4425
   ScaleWidth      =   7110
   ToolboxBitmap   =   "Program_Yakalayýcý_Kapatýcý.ctx":1255
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   4080
      Width           =   1935
   End
   Begin VB.ListBox List2 
      Height          =   3375
      Left            =   3600
      TabIndex        =   5
      Top             =   360
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   1095
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   4680
      Width           =   6855
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   480
      Top             =   6000
   End
   Begin VB.ListBox List1 
      Height          =   3375
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Listeyi Yenile ve Tekrardan Tara"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   4080
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "Bulunanlar"
      Height          =   255
      Left            =   3600
      TabIndex        =   9
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Label4"
      Height          =   195
      Left            =   3600
      TabIndex        =   8
      Top             =   3720
      Width           =   480
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "60"
      Height          =   195
      Left            =   3720
      TabIndex        =   7
      Top             =   4080
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   3720
      Width           =   480
   End
   Begin VB.Label Label3 
      Caption         =   "Açýk Programlar"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "Program_Yakalayýcý_Kapatýcý"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Type PROCESSENTRY32
dwSize As Long
cntUsage As Long
th32ProcessID As Long
th32DefaultHeapID As Long
th32ModuleID As Long
cntThreads As Long
th32ParentProcessID As Long
pcPriClassBase As Long
dwFlags As Long
szExeFile As String * 260
End Type
Dim lSnapShot As Long, lNextProcess As Long, Program As PROCESSENTRY32

Private Sub Command2_Click()
'List1deki Aktif Olan Ýþlemleri Yeniler
Timer1.Enabled = False
Label5.Caption = "60"
TerminateProcess OpenProcess(0, False, Val(List1.Text)), lExitCode
DoEvents
List1.Clear
Bul
Label1.Caption = "Toplam Aktif Ýþlem : " & List1.ListCount - 1
Timer1.Enabled = True
End Sub
Private Sub Command1_Click()
'manuel arama yapma
Bulucu
End Sub

Private Sub Bul()
lSnapShot = CreateToolhelp32Snapshot(&H2&, 0&)
Program.dwSize = Len(Program)
lNextProcess = Process32First(lSnapShot, Program)
Do While lNextProcess
'List1.AddItem Program.th32ProcessID & " " & Program.szExeFile id deðeri
List1.AddItem Program.szExeFile
lNextProcess = Process32Next(lSnapShot, Program)
Loop
End Sub



Private Sub Timer1_Timer()
Label5.Caption = Label5.Caption - 1

If Label5.Caption = "50" Then

Text2.Text = "winamp.exe"


Bulucu
Label4.Caption = "Toplam Aktif Bulunanlar : " & List2.ListCount
End If

If Label5.Caption = "45" Then
Text2.Text = "URL BULUCU.exe"
Bulucu
Label4.Caption = "Toplam Aktif Bulunanlar : " & List2.ListCount
End If

If Label5.Caption = "35" Then
Text2.Text = "taskeng.exe"
Bulucu
Label4.Caption = "Toplam Aktif Bulunanlar : " & List2.ListCount
End If

If Label5.Caption = "25" Then
Text2.Text = "taskhost.exe"
Bulucu
Label4.Caption = "Toplam Aktif Bulunanlar : " & List2.ListCount
End If
If Label5.Caption = "15" Then
Text2.Text = "ollydbg.exe"
Bulucu

Label4.Caption = "Toplam Aktif Bulunanlar : " & List2.ListCount
End If
If Label5.Caption = "5" Then
Text2.Text = "VB6.exe"
Bulucu

Label4.Caption = "Toplam Aktif Bulunanlar : " & List2.ListCount
End If
If Label5.Caption = "0" Then
Text2.Text = "URL BULUCU.exe"
Bulucu
Timer1.Enabled = False

Label5.Caption = "60"
Label4.Caption = "Toplam Aktif Bulunanlar : " & List2.ListCount
End If

If Label5.Caption = "60" Then
'Timer1.Enabled = True
End If

End Sub

Private Function Bulucu()
On Error Resume Next
Dim ArananKelime As String
Dim KelimeninYeri, AramayaBasla As Integer
On Error GoTo hata
ArananKelime = Text2 'text2 içindeki kelimeyi arayacaðýz
AramayaBasla = Text1.SelStart + Text1.SelLength 'arama yapýlacak metin uzunluðunda arama yapacaðýz
If AramayaBasla = 0 Or AramayaBasla = Len(Text1.Text) Then AramayaBasla = 1 'aranan kelime bulunmazsa baþa döneceðiz
KelimeninYeri = InStr(AramayaBasla, Text1.Text, ArananKelime, vbTextCompare)
Text1.SetFocus 'kelime bulunduðunda iþaretliyoruz
Text1.SelStart = KelimeninYeri - 1
Text1.SelLength = Len(ArananKelime)

List2.AddItem Text1.SelText

Exit Function
hata: 'arama metni sonuna geldiðimizde baþtan bir daha baþlýyoruz
Text1.SelStart = 1
End Function

Private Sub UserControl_Initialize()
'Açýk olan iþlemler listesini çekiyor
Bul
'List1 içindeki toplam deðerleri listeliyor
Label1.Caption = "Toplam Aktif Ýþlem : " & List1.ListCount - 1
'List1'deki deðelerleri Text1 içine aktarýyor
Dim i As Integer
For i = 0 To List1.ListCount
If Text1.Text = "" Then
Text1.Text = List1.list(i)
Else
Text1.Text = Text1.Text & vbNewLine & List1.list(i)
End If
Next i
Timer1.Enabled = True
End Sub
