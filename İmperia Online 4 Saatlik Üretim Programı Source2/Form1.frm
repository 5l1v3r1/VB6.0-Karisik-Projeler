VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   BackColor       =   &H80000016&
   Caption         =   "�mperia Online | 4 Saatlik �retim Program�"
   ClientHeight    =   7845
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11055
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   162
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7845
   ScaleWidth      =   11055
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   5040
      Top             =   240
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   240
      Left            =   8760
      TabIndex        =   29
      Top             =   360
      Width           =   735
   End
   Begin VB.ListBox List3 
      Height          =   1740
      Left            =   4200
      TabIndex        =   28
      Top             =   7920
      Width           =   3495
   End
   Begin VB.ListBox List2 
      Height          =   1740
      Left            =   120
      TabIndex        =   27
      Top             =   7920
      Width           =   3975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Durdur"
      Height          =   495
      Left            =   1920
      TabIndex        =   24
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Ba�lat"
      Height          =   495
      Left            =   240
      TabIndex        =   23
      Top             =   120
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   4440
      Top             =   240
   End
   Begin MSComDlg.CommonDialog Cmd1 
      Left            =   3720
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000D&
      Caption         =   "-Facebook Otomatik Kay�t"
      Height          =   6855
      Left            =   5520
      TabIndex        =   8
      Top             =   840
      Width           =   5415
      Begin VB.TextBox Text11 
         Height          =   1215
         Left            =   2640
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         Text            =   "Form1.frx":0000
         Top             =   5400
         Width           =   2535
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   3135
         Left            =   240
         TabIndex        =   25
         Top             =   3480
         Width           =   2295
         ExtentX         =   4048
         ExtentY         =   5530
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
      Begin VB.TextBox Text10 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   360
         Left            =   1560
         TabIndex        =   22
         Text            =   "1995"
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   375
         Left            =   960
         TabIndex        =   21
         Text            =   "05"
         Top             =   2880
         Width           =   495
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   360
         Left            =   360
         TabIndex        =   20
         Text            =   "08"
         Top             =   2880
         Width           =   495
      End
      Begin VB.TextBox Text7 
         Enabled         =   0   'False
         Height          =   360
         Left            =   2880
         TabIndex        =   19
         Top             =   2160
         Width           =   2295
      End
      Begin VB.TextBox Text6 
         Enabled         =   0   'False
         Height          =   360
         Left            =   360
         TabIndex        =   18
         Text            =   "Erkek"
         Top             =   2160
         Width           =   2295
      End
      Begin VB.TextBox Text5 
         Height          =   360
         Left            =   360
         TabIndex        =   17
         Top             =   1440
         Width           =   4815
      End
      Begin VB.TextBox Text4 
         Enabled         =   0   'False
         Height          =   360
         Left            =   2880
         TabIndex        =   16
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         Height          =   360
         Left            =   360
         TabIndex        =   15
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         Caption         =   "Do�um Tarihi"
         Height          =   240
         Left            =   360
         TabIndex        =   14
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         Caption         =   "Yeni �ifre"
         Height          =   240
         Left            =   2880
         TabIndex        =   13
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         Caption         =   "Cinsiyet"
         Height          =   240
         Left            =   360
         TabIndex        =   12
         Top             =   1920
         Width           =   705
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         Caption         =   "E-posta"
         Height          =   240
         Left            =   360
         TabIndex        =   11
         Top             =   1200
         Width           =   705
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         Caption         =   "Soyad�n"
         Height          =   240
         Left            =   2880
         TabIndex        =   10
         Top             =   480
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         Caption         =   "Ad�n"
         Height          =   240
         Left            =   360
         TabIndex        =   9
         Top             =   480
         Width           =   405
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      Caption         =   "-Ge�erli Eposta Adresleri"
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   5295
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000D&
         Caption         =   "Manuel Eposta Giri�i"
         Height          =   1335
         Left            =   240
         TabIndex        =   5
         Top             =   5400
         Width           =   4815
         Begin VB.CommandButton Command2 
            Caption         =   "Listeye Aktar"
            Height          =   375
            Left            =   240
            TabIndex        =   7
            Top             =   840
            Width           =   1935
         End
         Begin VB.TextBox Text2 
            Height          =   360
            Left            =   240
            TabIndex        =   6
            Top             =   360
            Width           =   4335
         End
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   345
         Left            =   1800
         TabIndex        =   4
         Top             =   4920
         Width           =   855
      End
      Begin VB.ListBox List1 
         Height          =   3900
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   4815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Aktar"
         Height          =   345
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         Caption         =   "Toplam Email"
         Height          =   240
         Left            =   360
         TabIndex        =   3
         Top             =   4920
         Width           =   1260
      End
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   255
      Left            =   5640
      TabIndex        =   31
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Label8"
      Height          =   240
      Left            =   5640
      TabIndex        =   30
      Top             =   480
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
'* Dosya Se�me Penceresi A�ar
Cmd1.ShowOpen
'* List1'i Tamamen Temizler
List1.Clear
'* �f D�ng�s� Ba�lang��
If Dir(Cmd1.FileName) <> "" Then
If Cmd1.FileName = "" Then
'* Dosya Se�ilmezse Bilgi Mesaj� verir
MsgBox "Dosya Se�ilmedi", vbInformation, "Uyar� ;"
'* Taramay� Ba�lat Butonunu Aktifle�tirir
'* De�ilse
Else
'* Se�ilen Dosyay� List1'e aktar�r
Open Cmd1.FileName For Input As #1
While Not EOF(1)
Input #1, a
List1.AddItem a
Wend
Close #1
'* Label'e Listbox'daki Toplam De�eri Yans�t�r
Text1.Text = List1.ListCount
'* Taramay� Ba�lat Butonunu Aktifle�tirir
End If
'* �f D�ng�s� Biti�
End If


End Sub

Private Sub Command2_Click()
List1.AddItem (Text2.Text)
Text1.Text = List1.ListCount

End Sub

Private Sub Command3_Click()
Timer1.Enabled = True

End Sub

Private Sub Command4_Click()
Timer1.Enabled = False

End Sub


Private Sub �ye_ol()
On Error Resume Next
WebBrowser1.Document.All.Item("firstname").Value = Text3.Text
WebBrowser1.Document.All.Item("lastname").Value = Text4.Text
WebBrowser1.Document.All.Item("email").Value = Text5.Text
WebBrowser1.Document.All.Item("gender").Value = "2"
WebBrowser1.Document.All.Item("pass").Value = Text7.Text

WebBrowser1.Document.All.Item("day").Value = Text8.Text
WebBrowser1.Document.All.Item("month").Value = Text9.Text
WebBrowser1.Document.All.Item("year").Value = Text10.Text

WebBrowser1.Document.All.Item("submit").Click
Text5.Text = ""
End Sub
Private Sub Rastgele_�retim()
Randomize
Dim harfler(4) As String
Dim say�lar(4) As Integer
harfler(0) = Chr(Math.Round(Rnd() * (122 - 97) + 97))
harfler(1) = Chr(Math.Round(Rnd() * (122 - 97) + 97))
harfler(2) = Chr(Math.Round(Rnd() * (122 - 97) + 97))
harfler(3) = Chr(Math.Round(Rnd() * (122 - 97) + 97))
harfler(4) = Chr(Math.Round(Rnd() * (122 - 97) + 97))
say�lar(0) = Math.Round(Rnd() * (9 - 0) + 0)
say�lar(1) = Math.Round(Rnd() * (9 - 0) + 0)
say�lar(2) = Math.Round(Rnd() * (9 - 0) + 0)
say�lar(3) = Math.Round(Rnd() * (9 - 0) + 0)
say�lar(4) = Math.Round(Rnd() * (9 - 0) + 0)
Text3.Text = "danger"
Text4.Text = "ousman"
Text7.Text = say�lar(1) & say�lar(2) & harfler(2) & harfler(0) & say�lar(4) & say�lar(0) & say�lar(3) & say�lar(1)
End Sub



Private Sub Command5_Click()
�ye_ol

End Sub

Private Sub Form_Load()
Cmd1.Filter = ".txt Dosyas�|*.txt"
WebBrowser1.Navigate ("https://m.facebook.com/r.php?refid=8")

End Sub


Private Sub Timer1_Timer()
On Error Resume Next
With List1
WebBrowser1.Navigate ("https://m.facebook.com/r.php?refid=8")
Text5.Text = ""
'*�ndex say�s�n� 1 artt�r�r
.ListIndex = .ListIndex + 1
Timer1.Interval = 10000
'*Se�ili eleman�n de�eri
Text5.Text = .List(.ListIndex)
Rastgele_�retim
'* Listenin son eleman�na gelince bilgi mesaj� verir
If .ListIndex = .ListCount - 1 Then
.ListIndex = -1
'* Timer1'i Durdurur
Timer1.Enabled = False
Timer1.Interval = 2000
MsgBox "bitti"
End If
End With

End Sub

Private Sub Timer2_Timer()
On Error Resume Next
Label9.Caption = Len(Text5.Text)

If Label9.Caption = 0 Then
Else

�ye_ol

End If
End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
On Error Resume Next
Text11.Text = WebBrowser1.Document.bOdy.innerHTML

If InStr(1, Text11, "Arkada�lar�n�n seni bulabilmesi i�in okul ve i� bilgilerini ekle.") Then
Open App.Path & "\deneme.txt" For Append As #1
Print #1, Text5.Text & " - " & Text7.Text & vbCrLf
Close #1

'MsgBox "Kullan�c� ad� veya parola hatal�.!", vbCritical, "Hata;"
Label8.Caption = "Kay�t Tamamland�"
WebBrowser1.Navigate ("https://m.facebook.com/logout.php?h=AfdOkh3Ll6IyJIjo&t=1395942607&refid=8")
End If
If InStr(1, Text11, "E-posta Adresini Onayla") Then
'MsgBox "Kullan�c� ad� veya parola hatal�.!", vbCritical, "Hata;"
Label8.Caption = "Kullan�lan bi eposta"
'WebBrowser1.Navigate ("https://m.facebook.com/logout.php?h=AfdOkh3Ll6IyJIjo&t=1395942607&refid=8")
End If
'L�tfen ge�erli bir e-posta adresi veya cep telefonu numaras� gir.
'E-posta Adresini Onayla
End Sub

