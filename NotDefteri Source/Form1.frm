VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Not Defteri | _DangerOusMaN_"
   ClientHeight    =   7065
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   11415
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog Cmd2 
      Left            =   12000
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog Cmd1 
      Left            =   12240
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Height          =   7095
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   0
      Width           =   11415
   End
   Begin VB.Menu dosya 
      Caption         =   "&Dosya"
      Begin VB.Menu yeni 
         Caption         =   "&Yeni"
         Shortcut        =   ^N
      End
      Begin VB.Menu aç 
         Caption         =   "&Aç..."
         Shortcut        =   ^O
      End
      Begin VB.Menu kaydet 
         Caption         =   "&Kaydet"
         Shortcut        =   ^S
      End
      Begin VB.Menu farklýkaydet 
         Caption         =   "&Farklý Kaydet..."
      End
      Begin VB.Menu bosluk 
         Caption         =   "-"
      End
      Begin VB.Menu cýkýs 
         Caption         =   "&Çýkýþ"
      End
   End
   Begin VB.Menu duzen 
      Caption         =   "&Düzen"
      Begin VB.Menu kes 
         Caption         =   "&Kes"
         Shortcut        =   ^X
      End
      Begin VB.Menu kopyala 
         Caption         =   "&Kopyala"
         Shortcut        =   ^C
      End
      Begin VB.Menu yapýþtýr 
         Caption         =   "&Yapýþtýr"
         Shortcut        =   ^V
      End
      Begin VB.Menu sil 
         Caption         =   "&Sil"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu bos3 
         Caption         =   "-"
      End
      Begin VB.Menu bul 
         Caption         =   "&Bul"
         Shortcut        =   ^F
      End
      Begin VB.Menu bos4 
         Caption         =   "-"
      End
      Begin VB.Menu tümünüseç 
         Caption         =   "&Tümünü Seç"
         Shortcut        =   ^A
      End
      Begin VB.Menu saattarih 
         Caption         =   "&Saat/Tarih"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu hakkýnda 
      Caption         =   "&Hakkýnda"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub aç_Click()
'* Dosya Seçme Penceresi Açar
Cmd1.ShowOpen
Text1.Text = ""
'* Ýf Döngüsü Baþlangýç
If Dir(Cmd1.FileName) <> "" Then

If Cmd1.FileName = "" Then
'* Dosya Seçilmezse Bilgi Mesajý verir
MsgBox "Dosya Seçilmedi", vbInformation, "Uyarý ;"
'* Deðilse
Else
'* Seçilen Dosyayý Text1'e aktarýr
Open Cmd1.FileName For Input As #1
Do While Not EOF(1)
Line Input #1, SATIR
metin = metin + Chr(10) + Chr(13) + SATIR
Loop
Close #1
Text1.Text = metin
Form1.Caption = Cmd1.FileTitle & " - Not Defteri | _DangerOusMaN_"
End If
'* Ýf Döngüsü Bitiþ
End If
End Sub

Private Sub bul_Click()
Form2.Show
End Sub

Private Sub farklýkaydet_Click()
kaydet_Click
End Sub

Private Sub Form_Load()
Cmd1.DialogTitle = "Metin Dosyasý Seçiniz"
Cmd1.Filter = "Metin Belgeleri(*.txt) |*.txt"

Cmd2.DialogTitle = "Kaydetmek Ýstediðiniz Dizin"
Cmd2.Filter = "Metin Belgeleri(*.txt) |*.txt"
End Sub



Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Form1.Hide
Form2.Hide
End
End Sub

Private Sub kaydet_Click()
Cmd2.ShowSave
If Cmd2.FileName = "" Then
'* Dosya Seçilmezse Bilgi Mesajý verir
MsgBox "Dosya Seçilmedi", vbInformation, "Uyarý ;"
'* Deðilse
Else
Open Cmd2.FileName For Output As #2
Print #2, Text1.Text
Close #2
End If
End Sub

Private Sub kes_Click()
Clipboard.SetText Text1.SelText
Text1.SelText = ""
End Sub

Private Sub kopyala_Click()
Clipboard.SetText Text1.SelText
End Sub

Private Sub saattarih_Click()
Text1.Text = Now
End Sub

Private Sub sil_Click()
Text1.Text = ""
End Sub

Private Sub tümünüseç_Click()
Dim Tümseç
Tümseç = Len(Text1.Text)
Text1.SelStart = 0
Text1.SelLength = Tümseç
Text1.SetFocus
End Sub

Private Sub yapýþtýr_Click()
Text1.SelText = Clipboard.GetText(1)
End Sub

Private Sub yeni_Click()
Dim Mesaj As VbMsgBoxResult
If Not Trim(Text1.Text) = "" Then
Mesaj = MsgBox("Deðiþiklikleri Kaydetmek Ýstiyor musunuz?", vbYesNoCancel, "Hata")
If Mesaj = vbYes Then
'Call save_now
Text1.Text = ""
End If
If Mesaj = vbNo Then
Text1.Text = ""
End If
End If
End Sub
