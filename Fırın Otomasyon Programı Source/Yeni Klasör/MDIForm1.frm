VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   9750
   ClientLeft      =   2490
   ClientTop       =   2580
   ClientWidth     =   12090
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.Menu anapanel 
      Caption         =   "&Gelir-Gider Tablosu"
   End
   Begin VB.Menu yetki_verme 
      Caption         =   "&Yetkilendirme"
      Begin VB.Menu yetki_düzenleme 
         Caption         =   "Yetki Düzenleme"
      End
      Begin VB.Menu kullanici_degistirme 
         Caption         =   "Kullanici Degistirme"
      End
   End
   Begin VB.Menu ayalar 
      Caption         =   "&Ayarlar"
      Begin VB.Menu firin_ekle 
         Caption         =   "Firin Adi Düzenleme"
      End
      Begin VB.Menu kullanici_log 
         Caption         =   "Kullanici Giris Kayitlari"
      End
      Begin VB.Menu bos 
         Caption         =   "-"
      End
      Begin VB.Menu veritabani_yedekle 
         Caption         =   "Veritabani Yedekleme"
      End
      Begin VB.Menu veritabani_bilgileri 
         Caption         =   "Veritabani Bilgileri"
      End
   End
   Begin VB.Menu programi_kapat 
      Caption         =   "&Programi Kapat"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub anapanel_Click()
On Error Resume Next
Form2.Show
Form3.Hide
Form4.Hide
Form5.Hide
Form6.Hide
Form7.Hide
Form8.Hide
End Sub
Private Sub firin_ekle_Click()
On Error Resume Next
Form3.Hide
Form2.Hide
Form5.Hide
Form6.Hide
Form7.Hide
Form8.Hide
Form4.Show
End Sub
Private Sub kullanici_degistirme_Click()
On Error Resume Next
Form1.Text1.Text = ""
Form1.Text2.Text = ""
Form3.Hide
Form2.Hide
Form4.Hide
Form5.Hide
Form6.Hide
Form7.Hide
Form8.Hide
MDIForm1.Hide
Form1.Show
End Sub
Private Sub kullanici_log_Click()
On Error Resume Next
Form3.Hide
Form2.Hide
Form4.Hide
Form6.Hide
Form7.Hide
Form8.Hide
Form5.Show
End Sub
Private Sub MDIForm_Load()
Form4.Show
Form4.Visible = False
MDIForm1.Caption = Form4.Text1.Text
End Sub
Private Sub MDIForm_Unload(Cancel As Integer)
End
End Sub
Private Sub programi_kapat_Click()
Dim cýkýs As String
cýkýs = MsgBox("Programdan Çikis Yapilsin mi?", vbYesNo + vbInformation, "Kapatiliyor...")
If cýkýs = vbYes Then
Unload Me
End If
End Sub
Private Sub veritabani_bilgileri_Click()
On Error Resume Next
Form3.Hide
Form2.Hide
Form4.Hide
Form5.Hide
Form6.Hide
Form7.Hide
Form8.Show
End Sub
Private Sub veritabani_yedekle_Click()
On Error Resume Next
Form3.Hide
Form2.Hide
Form4.Hide
Form5.Hide
Form6.Hide
Form7.Hide
Form8.Hide
Form9.Show
End Sub
Private Sub yetki_düzenleme_Click()
On Error Resume Next
Form3.Show
Form2.Hide
Form4.Hide
Form5.Hide
Form6.Hide
Form7.Hide
Form8.Hide
End Sub
