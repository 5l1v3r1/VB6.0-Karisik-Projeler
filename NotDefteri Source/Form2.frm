VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bul"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5730
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Büyük Küçük Harf Eþleþtir"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1220
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Caption         =   "Yön"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      TabIndex        =   4
      Top             =   600
      Width           =   1935
      Begin VB.OptionButton Option2 
         Caption         =   "Aþaðý"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   960
         TabIndex        =   6
         Top             =   270
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Yukarý"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ýptal"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sonrakini Bul"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Aranan :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ArananKelime As String
Dim KelimeninYeri, AramayaBasla As Integer
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40

Private Sub Command1_Click()
If Option1.Value = True Then
On Error GoTo hata
On Error Resume Next
ArananKelime = Text1 'text2 içindeki kelimeyi arayacaðýz
AramayaBasla = Form1.Text1.SelStart + Form1.Text1.SelLength 'arama yapýlacak metin uzunluðunda arama yapacaðýz
If AramayaBasla = 0 Or AramayaBasla = Len(Form1.Text1.Text) Then AramayaBasla = 1 'aranan kelime bulunmazsa baþa döneceðiz
KelimeninYeri = InStr(AramayaBasla, Form1.Text1.Text, ArananKelime, vbTextCompare)
Form1.Text1.SetFocus 'kelime bulunduðunda iþaretliyoruz
Form1.Text1.SelStart = KelimeninYeri - 1
Form1.Text1.SelLength = Len(ArananKelime)
Exit Sub
hata: 'arama metni sonuna geldiðimizde baþtan bir daha baþlýyoruz
Text1.SelStart = 1
End If

If Option2.Value > True Then
On Error GoTo hata1
ArananKelime = Text1 'text2 içindeki kelimeyi arayacaðýz

AramayaBasla = Form1.Text1.SelStart + Form1.Text1.SelLength 'arama yapýlacak metin uzunluðunda arama yapacaðýz
If AramayaBasla = 0 Or AramayaBasla = Len(Form1.Text1.Text) Then AramayaBasla = 1 'aranan kelime bulunmazsa baþa döneceðiz
KelimeninYeri = InStr(AramayaBasla, Form1.Text1.Text, ArananKelime, vbTextCompare)
Form1.Text1.SetFocus 'kelime bulunduðunda iþaretliyoruz
Form1.Text1.SelStart = KelimeninYeri - 1
Form1.Text1.SelLength = Len(ArananKelime)
Exit Sub
hata1: 'arama metni sonuna geldiðimizde baþtan bir daha baþlýyoruz
Text1.SelStart = 1
End If
End Sub


Private Sub Command2_Click()
Form2.Hide
End Sub



Private Sub Form_Load()
Dim lFlag As Long
    'Form her zaman üstte
    lFlag = HWND_TOPMOST
    SetWindowPos Form2.hwnd, lFlag, Form2.Left / Screen.TwipsPerPixelX, Form2.Top / Screen.TwipsPerPixelY, Form2.Width / Screen.TwipsPerPixelX, Form2.Height / Screen.TwipsPerPixelY, SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form2.Hide
End Sub

