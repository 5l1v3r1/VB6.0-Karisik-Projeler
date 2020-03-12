VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "ActiveX.ocx"
Begin VB.Form Form1 
   Caption         =   "Beyaz Sayfalar | Copyright™ Osman Yavuz"
   ClientHeight    =   7485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   ScaleHeight     =   7485
   ScaleWidth      =   7425
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Ileri"
      Height          =   375
      Left            =   3720
      TabIndex        =   27
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Yeniden Ara"
      Height          =   375
      Left            =   2520
      TabIndex        =   26
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Text8 
      Height          =   3015
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   24
      Top             =   4320
      Width           =   7215
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   9960
      TabIndex        =   23
      Text            =   $"Form1.frx":0000
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   9960
      TabIndex        =   22
      Text            =   "Yardim"
      Top             =   240
      Width           =   1095
   End
   Begin XtremeSuiteControls.WebBrowser Web1 
      Height          =   2775
      Left            =   9960
      TabIndex        =   21
      Top             =   3240
      Width           =   3255
      _Version        =   851968
      _ExtentX        =   5741
      _ExtentY        =   4895
      _StockProps     =   173
      BackColor       =   -2147483643
      ScriptErrorsSuppressed=   -1  'True
   End
   Begin VB.TextBox Text5 
      Height          =   2175
      Left            =   9960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   19
      Text            =   "Form1.frx":0007
      Top             =   960
      Width           =   3255
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   11760
      Top             =   120
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Durdur"
      Height          =   375
      Left            =   1320
      TabIndex        =   18
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Baslat"
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Height          =   3615
      Left            =   3360
      TabIndex        =   4
      Top             =   600
      Width           =   3975
      Begin VB.Frame Frame3 
         Caption         =   "-Dogrulama"
         Height          =   2415
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   3735
         Begin VB.CommandButton Command4 
            Caption         =   "Yenile"
            Height          =   255
            Left            =   2880
            TabIndex        =   25
            Top             =   2040
            Width           =   735
         End
         Begin VB.TextBox Text4 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   600
            TabIndex        =   16
            Top             =   2040
            Width           =   2175
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            FillColor       =   &H00FFFFFF&
            ForeColor       =   &H80000008&
            Height          =   1455
            Left            =   120
            ScaleHeight     =   1425
            ScaleWidth      =   2865
            TabIndex        =   14
            Top             =   240
            Width           =   2895
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Kod"
            Height          =   195
            Left            =   240
            TabIndex        =   15
            Top             =   2040
            Width           =   285
         End
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   2520
         TabIndex        =   12
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Form1.frx":000D
         Left            =   2520
         List            =   "Form1.frx":0107
         TabIndex        =   10
         Text            =   "Isparta-246"
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   720
         TabIndex        =   9
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   720
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Semt"
         Height          =   195
         Left            =   2040
         TabIndex        =   11
         Top             =   720
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Sehir"
         Height          =   195
         Left            =   2040
         TabIndex        =   7
         Top             =   360
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Isim"
         Height          =   195
         Left            =   360
         TabIndex        =   6
         Top             =   720
         Width           =   270
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Soyad"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   450
      End
   End
   Begin MSComDlg.CommonDialog Cmd1 
      Left            =   11280
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "-Soyad Listesi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   3135
      Begin VB.CommandButton Command1 
         Caption         =   "Aktar"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   2880
         Width           =   855
      End
      Begin VB.ListBox List1 
         Height          =   2595
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "Toplam Soyad :"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   3240
         Width           =   2655
      End
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Durum"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   4920
      TabIndex        =   20
      Top             =   240
      Width           =   555
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Cmd1.ShowOpen
List1.Clear
If Dir(Cmd1.FileName) <> "" Then
If Cmd1.FileName = "" Then
MsgBox "Dosya Seçilmedi", vbInformation, "Uyarý ;"
Else
Open Cmd1.FileName For Input As #1
While Not EOF(1)
Input #1, a
List1.AddItem a
Wend
Close #1
Label1.Caption = "Toplam Soyad: " & List1.ListCount
End If
End If
With List1
.ListIndex = .ListIndex + 1
Label1.Caption = "Toplam Soyad : " & List1.ListCount
Text1.Text = .List(.ListIndex)
If .ListIndex = .ListCount - 1 Then
.ListIndex = -1
End If
End With
End Sub

Private Sub Command2_Click()
'Timer1.Enabled = True
Web1.Document.All.Item("NAM").Value = Text1.Text
Web1.Document.All.Item("GIV").Value = Text2.Text
Web1.Document.All.Item("STN").Value = "246-Isparta"
Web1.Document.All.Item("captchaStr").Value = Text4.Text
Web1.Document.All.Item("QNA").Value = Text3.Text
Web1.Navigate "javascript:DoSearch()"
End Sub

Private Sub Command3_Click()
Timer1.Enabled = False
End Sub

Private Sub Command4_Click()
Web1.Refresh
End Sub

Private Sub Command5_Click()
Web1.Navigate "http://www.ttrehber.turktelekom.com.tr/trk-wp/IDA2"
End Sub

Private Sub Form_Load()
'*Cdm1 Pencere Baþlýðý
Cmd1.DialogTitle = "Soyad Listenizi Seçiniz..."
'*Seçilecek Dosyanýn Uzantýsý
Cmd1.Filter = ".txt Dosyasý|*.txt"
Web1.Navigate "http://www.ttrehber.turktelekom.com.tr/trk-wp/IDA2"
End Sub



Private Sub List1_Click()
Text1.Text = List1.Text
End Sub

Private Sub Text5_Change()
On Error Resume Next
Dim AranacakYer, aranan As String
Dim ilk, son
AranacakYer = Text5.Text
deg1 = deg1 & AranacakYer

aranan = Text6
ilk = InStr(1, deg1, aranan) + Len(aranan)
If ilk <> Len(aranan) Then
son = InStr(ilk, deg1, Text7) - ilk
Text8.Text = Mid(deg1, ilk, son)
Exit Sub
MsgBox "Bitti"
End If
End Sub

Private Sub Text8_Change()
Open "C:\Referanslar.txt" For Append As #1
Print #1, Text8.Text;
Close #1
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
With List1
.ListIndex = .ListIndex + 1
Label1.Caption = "Toplam Soyad : " & List1.ListCount
Text1.Text = .List(.ListIndex)
If .ListIndex = .ListCount - 1 Then
.ListIndex = -1
End If
End With
End Sub


Private Sub Web1_DownloadComplete()
On Error Resume Next
Dim O As Object
 Text4.Text = ""
Set O = Web1.Document.bOdy.createControlRange()
Call O.Add(Web1.Document.All("captchaImage"))
Call O.execCommand("Copy")
 
Set Picture1.Picture = Clipboard.GetData
Text5.Text = Web1.Document.bOdy.innertext

If InStr(1, Text5, "Geçersiz arama kriteri.") Then
Label7.Caption = "Geçersiz arama kriteri.!"
End If
If InStr(1, Text5, "Yanlis sözcük dogrulama.") Then
Label7.Caption = "Yanlis sözcük dogrulama..!"
End If
'

If InStr(1, Text5, "Abone bulunamadi.") Then
Label7.Caption = "Abone bulunamadi..!"
End If
End Sub
