VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Knight Online Baglanma"
   ClientHeight    =   1200
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   3720
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   1440
      Locked          =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   2400
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog Cmd1 
      Left            =   480
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Kapat"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Oyunu Baslat"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "C:\NTTGame\KnightOnlineEn"
      Top             =   360
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   """KnightOnline.exe"" Dizini Seç :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function CreateP Lib "DLL.dll" (ByVal Direct As String, ByVal DLL As String, ByVal pHANDLE As Long) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long


Dim ArananKelime As String
Dim KelimeninYeri, AramayaBasla As Integer
Private Function Bul()
On Error GoTo hata

ArananKelime = "\KnightOnLine.exe" 'text2 içindeki kelimeyi arayacaðýz
AramayaBasla = Text2.SelStart + Text2.SelLength 'arama yapýlacak metin uzunluðunda arama yapacaðýz
If AramayaBasla = 0 Or AramayaBasla = Len(Text2.Text) Then AramayaBasla = 1 'aranan kelime bulunmazsa baþa döneceðiz
KelimeninYeri = InStr(AramayaBasla, Text2.Text, ArananKelime, vbTextCompare)
Text2.SetFocus 'kelime bulunduðunda iþaretliyoruz
Text2.SelStart = KelimeninYeri - 1
Text2.SelLength = Len(ArananKelime)
Text2.SelText = ""
Text1.Text = Text2.Text

'Knight Baglama
Call WritePrivateProfileString("Screen", "Path", App.Path, Text1.Text & "\Option.ini")
CreateP Text1.Text, "plugin.dll", KOHandle
Form1.Hide
Form2.Show

Exit Function
hata: ' arama metni sonuna geldiðimizde baþtan bir daha baþlýyoruz
Text2.SelStart = 1
End Function



Private Sub Command1_Click()
Cmd1.ShowOpen
Dim koYolu As String
koYolu = Cmd1.FileName
Text2.Text = koYolu
Bul
End Sub

Private Sub Command2_Click()
Dim msg As VbMsgBoxResult
msg = MsgBox("Programi kapatmak mi istiyorsunuz ?", vbInformation + vbYesNo, "Kapatiliyor...")
If msg = vbYes Then
End
End If
End Sub


Private Sub Form_Load()
Cmd1.DialogTitle = "KnightOnline.exe Seçiniz"
Cmd1.Filter = "KnightOnline.exe |*.exe"
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Label1_Click()
Form1.Hide
Form2.Show
End Sub
