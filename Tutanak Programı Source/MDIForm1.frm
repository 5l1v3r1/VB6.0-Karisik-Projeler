VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Hastem Tutanak Programý | Coded By / Osman Yavuz"
   ClientHeight    =   8340
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10635
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu tutanak 
      Caption         =   "Tutanak Formu"
   End
   Begin VB.Menu eklemeeee 
      Caption         =   "Kayýt Ekleme"
      Begin VB.Menu musteri_ekle 
         Caption         =   "&Müþteri Ekle"
      End
      Begin VB.Menu urun_ekle 
         Caption         =   "&Ürün Ekle"
      End
      Begin VB.Menu calýsan_ekle 
         Caption         =   "&Çalýþan Ekle"
      End
   End
   Begin VB.Menu bos 
      Caption         =   ""
      Enabled         =   0   'False
   End
   Begin VB.Menu yazdýr 
      Caption         =   "Yazdýr"
   End
   Begin VB.Menu sacxzc 
      Caption         =   ""
   End
   Begin VB.Menu tutanak_geçmiþi 
      Caption         =   "Tutanak Geçmiþi"
   End
   Begin VB.Menu bos2 
      Caption         =   ""
      Enabled         =   0   'False
   End
   Begin VB.Menu kapat 
      Caption         =   "Programý Kapat"
   End
   Begin VB.Menu asd 
      Caption         =   "asd"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub asd_Click()
Dim ExcelNesne As Object
Set ExcelNesne = CreateObject("Excel.SHEET")
ExcelNesne.Application.Visible = True
ExcelNesne.Application.Cells(1, 1).Value = "asd"
ExcelNesne.Application.Save ("c:/asd.xls" = "denem")
MsgBox ""
End Sub

Private Sub calýsan_ekle_Click()
Form13.Show
Form1.Hide
End Sub

Private Sub kapat_Click()
cýkýs = MsgBox("Programdan Çýkýþ Yapýlsýn mý ?", vbYesNo + vbInformation, "Çýkýþ yapýlýyor...")
If cýkýs = vbYes Then
End
End If
End Sub

Private Sub musteri_ekle_Click()
Form1.Hide
Form11.Show
End Sub

Private Sub tutanak_Click()
Form1.Show
End Sub

Private Sub tutanak_geçmiþi_Click()
Form1.Hide
Form14.Show
End Sub

Private Sub urun_ekle_Click()
Form12.Show
Form1.Hide
End Sub

Private Sub yazdýr_Click()
On Error Resume Next
Dim ExcelNesne As Object
Set ExcelNesne = CreateObject("Excel.SHEET")
ExcelNesne.Application.Visible = True

ExcelNesne.Application.Cells(2, 4).Font.Size = 12
ExcelNesne.Application.Cells(2, 4).Font.Bold = True
'ExcelNesne.Application.Cells(4, 8).Font.Underline = True
ExcelNesne.Application.Cells(2, 4).Font.Color = vbBlack
ExcelNesne.Application.Cells(2, 4).ColumnWidth = 12
ExcelNesne.Application.Cells(2, 4).Value = "TUTANAK"
 
ExcelNesne.Application.Cells(3, 1).Font.Size = 12
ExcelNesne.Application.Cells(3, 1).Font.Bold = True
ExcelNesne.Application.Cells(3, 1).Font.Underline = True
ExcelNesne.Application.Cells(3, 1).Font.Color = vbBlack
ExcelNesne.Application.Cells(3, 1).ColumnWidth = 12
ExcelNesne.Application.Cells(3, 1).Value = "HASTEM"

ExcelNesne.Application.Cells(5, 1).Font.Size = 10
ExcelNesne.Application.Cells(5, 1).Font.Bold = True
ExcelNesne.Application.Cells(5, 1).Font.Color = vbBlack
ExcelNesne.Application.Cells(5, 1).ColumnWidth = 10
ExcelNesne.Application.Cells(5, 1).Value = "Sayýn              "

ExcelNesne.Application.Cells(5, 2).Font.Size = 10
'ExcelNesne.Application.Cells(7, 2).Font.Bold = True
ExcelNesne.Application.Cells(5, 2).Font.Underline = True
ExcelNesne.Application.Cells(5, 2).Font.Color = vbBlack
ExcelNesne.Application.Cells(5, 2).ColumnWidth = 10
ExcelNesne.Application.Cells(5, 2).Value = Form1.DataGrid1.Text


ExcelNesne.Application.Cells(7, 1).Font.Size = 10
'ExcelNesne.Application.Cells(9, 1).Font.Bold = True
'ExcelNesne.Application.Cells(9, 1).Font.Underline = True
ExcelNesne.Application.Cells(7, 1).Font.Color = vbBlack
ExcelNesne.Application.Cells(7, 1).ColumnWidth = 10
ExcelNesne.Application.Cells(7, 1).Value = "Aþaðýda miktarlarý belirtilen ürünleri   " & Time & " / " & Date & "   tarihinde tarafýnýza eksiksiz teslim edilmiþtir."

ExcelNesne.Application.Cells(8, 7).Font.Size = 10
'ExcelNesne.Application.Cells(10, 11).Font.Bold = True
'ExcelNesne.Application.Cells(10, 11).Font.Underline = True
ExcelNesne.Application.Cells(8, 7).Font.Color = vbBlack
ExcelNesne.Application.Cells(8, 7).ColumnWidth = 10
ExcelNesne.Application.Cells(8, 7).Value = "Saygýlarýmýzla."

'ÜRÜN ADI
ExcelNesne.Application.Cells(10, 1).Font.Size = 10
ExcelNesne.Application.Cells(10, 1).Font.Bold = True
ExcelNesne.Application.Cells(10, 1).Font.Color = vbBlack
ExcelNesne.Application.Cells(10, 1).ColumnWidth = 10
ExcelNesne.Application.Cells(10, 1).Value = "ÜRÜN ADI"

'MÝKTAR
ExcelNesne.Application.Cells(10, 5).Font.Size = 10
ExcelNesne.Application.Cells(10, 5).Font.Bold = True
ExcelNesne.Application.Cells(10, 5).Font.Color = vbBlack
ExcelNesne.Application.Cells(10, 5).ColumnWidth = 10
ExcelNesne.Application.Cells(10, 5).Value = "MÝKTAR"

'BÝRÝM
ExcelNesne.Application.Cells(10, 6).Font.Size = 10
ExcelNesne.Application.Cells(10, 6).Font.Bold = True
ExcelNesne.Application.Cells(10, 6).Font.Color = vbBlack
ExcelNesne.Application.Cells(10, 6).ColumnWidth = 10
ExcelNesne.Application.Cells(10, 6).Value = "BÝRÝM"


i = i + 1

'# ÜRÜN ADLARI
ExcelNesne.Application.Cells(11, 1).Font.Size = 10
ExcelNesne.Application.Cells(12, 1).Font.Size = 10
ExcelNesne.Application.Cells(13, 1).Font.Size = 10
ExcelNesne.Application.Cells(14, 1).Font.Size = 10
ExcelNesne.Application.Cells(15, 1).Font.Size = 10
ExcelNesne.Application.Cells(16, 1).Font.Size = 10
ExcelNesne.Application.Cells(17, 1).Font.Size = 10
ExcelNesne.Application.Cells(18, 1).Font.Size = 10
ExcelNesne.Application.Cells(19, 1).Font.Size = 10

ExcelNesne.Application.Cells(11, 1).Value = Form1.Text4.Text
ExcelNesne.Application.Cells(12, 1).Value = Form1.Text6.Text
ExcelNesne.Application.Cells(13, 1).Value = Form1.Text9.Text
ExcelNesne.Application.Cells(14, 1).Value = Form1.Text12.Text
ExcelNesne.Application.Cells(15, 1).Value = Form1.Text15.Text
ExcelNesne.Application.Cells(16, 1).Value = Form1.Text18.Text
ExcelNesne.Application.Cells(17, 1).Value = Form1.Text21.Text
ExcelNesne.Application.Cells(18, 1).Value = Form1.Text24.Text
ExcelNesne.Application.Cells(19, 1).Value = Form1.Text27.Text

'# MÝKTAR
ExcelNesne.Application.Cells(11, 5).Columnheight = 5

ExcelNesne.Application.Cells(11, 5).Font.Size = 10
ExcelNesne.Application.Cells(12, 5).Font.Size = 10
ExcelNesne.Application.Cells(13, 5).Font.Size = 10
ExcelNesne.Application.Cells(14, 5).Font.Size = 10
ExcelNesne.Application.Cells(15, 5).Font.Size = 10
ExcelNesne.Application.Cells(16, 5).Font.Size = 10
ExcelNesne.Application.Cells(17, 5).Font.Size = 10
ExcelNesne.Application.Cells(18, 5).Font.Size = 10
ExcelNesne.Application.Cells(19, 5).Font.Size = 10

ExcelNesne.Application.Cells(11, 5).Value = Form1.Text1.Text
ExcelNesne.Application.Cells(12, 5).Value = Form1.Text7.Text
ExcelNesne.Application.Cells(13, 5).Value = Form1.Text10.Text
ExcelNesne.Application.Cells(14, 5).Value = Form1.Text13.Text
ExcelNesne.Application.Cells(15, 5).Value = Form1.Text16.Text
ExcelNesne.Application.Cells(16, 5).Value = Form1.Text19.Text
ExcelNesne.Application.Cells(17, 5).Value = Form1.Text22.Text
ExcelNesne.Application.Cells(18, 5).Value = Form1.Text25.Text
ExcelNesne.Application.Cells(19, 5).Value = Form1.Text28.Text

'# BÝRÝM
ExcelNesne.Application.Cells(11, 6).Font.Size = 10
ExcelNesne.Application.Cells(12, 6).Font.Size = 10
ExcelNesne.Application.Cells(13, 6).Font.Size = 10
ExcelNesne.Application.Cells(14, 6).Font.Size = 10
ExcelNesne.Application.Cells(15, 6).Font.Size = 10
ExcelNesne.Application.Cells(16, 6).Font.Size = 10
ExcelNesne.Application.Cells(17, 6).Font.Size = 10
ExcelNesne.Application.Cells(18, 6).Font.Size = 10
ExcelNesne.Application.Cells(19, 6).Font.Size = 10

ExcelNesne.Application.Cells(11, 6).Value = Form1.Combo1.Text

ExcelNesne.Application.Cells(12, 6).Value = Form1.Combo2.Text
ExcelNesne.Application.Cells(13, 6).Value = Form1.Combo3.Text
ExcelNesne.Application.Cells(14, 6).Value = Form1.Combo4.Text
ExcelNesne.Application.Cells(15, 6).Value = Form1.Combo5.Text
ExcelNesne.Application.Cells(16, 6).Value = Form1.Combo6.Text
ExcelNesne.Application.Cells(17, 6).Value = Form1.Combo7.Text
ExcelNesne.Application.Cells(18, 6).Value = Form1.Combo8.Text
ExcelNesne.Application.Cells(19, 6).Value = Form1.Combo9.Text

'# SON

ExcelNesne.Application.Cells(21, 1).Font.Size = 10
ExcelNesne.Application.Cells(21, 3).Font.Size = 10
ExcelNesne.Application.Cells(22, 5).Font.Size = 10
ExcelNesne.Application.Cells(21, 7).Font.Size = 10
ExcelNesne.Application.Cells(21, 1).Value = "Teslim Eden"
ExcelNesne.Application.Cells(21, 1).Font.Bold = True
ExcelNesne.Application.Cells(21, 4).Value = "Onaylayan"
ExcelNesne.Application.Cells(21, 4).Font.Bold = True
ExcelNesne.Application.Cells(22, 4).Value = Form1.DataGrid2.Text
ExcelNesne.Application.Cells(21, 7).Value = "Teslim Alan"
ExcelNesne.Application.Cells(21, 7).Font.Bold = True



'###EXCEL KAYDETME###
If "Sari Kaynak" = Form1.DataGrid1.Text Then
ExcelNesne.SaveCopyAs ("C:\HastemTutanakGecmisleri\" & Form1.DataGrid1.Text & "\" & "[" & Date & " - " & Format(Time, "hh.nn") & "] " & Form1.DataGrid1.Text & ".xls")
End If

If "Ahmet Melih Anadolu Lisesi" = Form1.DataGrid1.Text Then
ExcelNesne.SaveCopyAs ("C:\HastemTutanakGecmisleri\" & Form1.DataGrid1.Text & "\" & "[" & Date & " - " & Format(Time, "hh.nn") & "] " & Form1.DataGrid1.Text & ".xls")
End If

If "ALTINBAÞAK GIDA" = Form1.DataGrid1.Text Then
ExcelNesne.SaveCopyAs ("C:\HastemTutanakGecmisleri\" & Form1.DataGrid1.Text & "\" & "[" & Date & " - " & Format(Time, "hh.nn") & "] " & Form1.DataGrid1.Text & ".xls")
End If

If "Anadolu Öðretmen Lisesi" = Form1.DataGrid1.Text Then
ExcelNesne.SaveCopyAs ("C:\HastemTutanakGecmisleri\" & Form1.DataGrid1.Text & "\" & "[" & Date & " - " & Format(Time, "hh.nn") & "] " & Form1.DataGrid1.Text & ".xls")
End If

If "Aþ Evi" = Form1.DataGrid1.Text Then
ExcelNesne.SaveCopyAs ("C:\HastemTutanakGecmisleri\" & Form1.DataGrid1.Text & "\" & "[" & Date & " - " & Format(Time, "hh.nn") & "] " & Form1.DataGrid1.Text & ".xls")
End If

If "Aybilge" = Form1.DataGrid1.Text Then
ExcelNesne.SaveCopyAs ("C:\HastemTutanakGecmisleri\" & Form1.DataGrid1.Text & "\" & "[" & Date & " - " & Format(Time, "hh.nn") & "] " & Form1.DataGrid1.Text & ".xls")
End If

If "Basmacý Oðlu Otel" = Form1.DataGrid1.Text Then
ExcelNesne.SaveCopyAs ("C:\HastemTutanakGecmisleri\" & Form1.DataGrid1.Text & "\" & "[" & Date & " - " & Format(Time, "hh.nn") & "] " & Form1.DataGrid1.Text & ".xls")
End If

If "Çöpçü Restaurant" = Form1.DataGrid1.Text Then
ExcelNesne.SaveCopyAs ("C:\HastemTutanakGecmisleri\" & Form1.DataGrid1.Text & "\" & "[" & Date & " - " & Format(Time, "hh.nn") & "] " & Form1.DataGrid1.Text & ".xls")
End If

If "Davraz Yaþam Hastanesil" = Form1.DataGrid1.Text Then
ExcelNesne.SaveCopyAs ("C:\HastemTutanakGecmisleri\" & Form1.DataGrid1.Text & "\" & "[" & Date & " - " & Format(Time, "hh.nn") & "] " & Form1.DataGrid1.Text & ".xls")
End If

If "Deveci Ticaret" = Form1.DataGrid1.Text Then
ExcelNesne.SaveCopyAs ("C:\HastemTutanakGecmisleri\" & Form1.DataGrid1.Text & "\" & "[" & Date & " - " & Format(Time, "hh.nn") & "] " & Form1.DataGrid1.Text & ".xls")
End If

If "Doðum Evi Ana Okulu" = Form1.DataGrid1.Text Then
ExcelNesne.SaveCopyAs ("C:\HastemTutanakGecmisleri\" & Form1.DataGrid1.Text & "\" & "[" & Date & " - " & Format(Time, "hh.nn") & "] " & Form1.DataGrid1.Text & ".xls")
End If

If "Eðirdir Türem" = Form1.DataGrid1.Text Then
ExcelNesne.SaveCopyAs ("C:\HastemTutanakGecmisleri\" & Form1.DataGrid1.Text & "\" & "[" & Date & " - " & Format(Time, "hh.nn") & "] " & Form1.DataGrid1.Text & ".xls")
End If

If "Fen Lisesil" = Form1.DataGrid1.Text Then
ExcelNesne.SaveCopyAs ("C:\HastemTutanakGecmisleri\" & Form1.DataGrid1.Text & "\" & "[" & Date & " - " & Format(Time, "hh.nn") & "] " & Form1.DataGrid1.Text & ".xls")
End If

If "Gazi Lisesi" = Form1.DataGrid1.Text Then
ExcelNesne.SaveCopyAs ("C:\HastemTutanakGecmisleri\" & Form1.DataGrid1.Text & "\" & "[" & Date & " - " & Format(Time, "hh.nn") & "] " & Form1.DataGrid1.Text & ".xls")
End If

If "Gülbirlik - Rosense" = Form1.DataGrid1.Text Then
ExcelNesne.SaveCopyAs ("C:\HastemTutanakGecmisleri\" & Form1.DataGrid1.Text & "\" & "[" & Date & " - " & Format(Time, "hh.nn") & "] " & Form1.DataGrid1.Text & ".xls")
End If


If "Gülköy Et" = Form1.DataGrid1.Text Then
ExcelNesne.SaveCopyAs ("C:\HastemTutanakGecmisleri\" & Form1.DataGrid1.Text & "\" & "[" & Date & " - " & Format(Time, "hh.nn") & "] " & Form1.DataGrid1.Text & ".xls")
End If

If "Güzel Sanatlar Lisesi" = Form1.DataGrid1.Text Then
ExcelNesne.SaveCopyAs ("C:\HastemTutanakGecmisleri\" & Form1.DataGrid1.Text & "\" & "[" & Date & " - " & Format(Time, "hh.nn") & "] " & Form1.DataGrid1.Text & ".xls")
End If

If "Ýkbal Restaurant" = Form1.DataGrid1.Text Then
ExcelNesne.SaveCopyAs ("C:\HastemTutanakGecmisleri\" & Form1.DataGrid1.Text & "\" & "[" & Date & " - " & Format(Time, "hh.nn") & "] " & Form1.DataGrid1.Text & ".xls")
End If

If "Isparta Tabildot" = Form1.DataGrid1.Text Then
ExcelNesne.SaveCopyAs ("C:\HastemTutanakGecmisleri\" & Form1.DataGrid1.Text & "\" & "[" & Date & " - " & Format(Time, "hh.nn") & "] " & Form1.DataGrid1.Text & ".xls")
End If

If "Isparta Türem" = Form1.DataGrid1.Text Then
ExcelNesne.SaveCopyAs ("C:\HastemTutanakGecmisleri\" & Form1.DataGrid1.Text & "\" & "[" & Date & " - " & Format(Time, "hh.nn") & "] " & Form1.DataGrid1.Text & ".xls")
End If

If "Kaçýkoç Lisesi Okulu" = Form1.DataGrid1.Text Then
ExcelNesne.SaveCopyAs ("C:\HastemTutanakGecmisleri\" & Form1.DataGrid1.Text & "\" & "[" & Date & " - " & Format(Time, "hh.nn") & "] " & Form1.DataGrid1.Text & ".xls")
End If

If "Kývýlcým Medikal" = Form1.DataGrid1.Text Then
ExcelNesne.SaveCopyAs ("C:\HastemTutanakGecmisleri\" & Form1.DataGrid1.Text & "\" & "[" & Date & " - " & Format(Time, "hh.nn") & "] " & Form1.DataGrid1.Text & ".xls")
End If

If "MEHMET KIRMIZIBAYRAK" = Form1.DataGrid1.Text Then
ExcelNesne.SaveCopyAs ("C:\HastemTutanakGecmisleri\" & Form1.DataGrid1.Text & "\" & "[" & Date & " - " & Format(Time, "hh.nn") & "] " & Form1.DataGrid1.Text & ".xls")
End If

If "Mekke Eðitim Vakfý" = Form1.DataGrid1.Text Then
ExcelNesne.SaveCopyAs ("C:\HastemTutanakGecmisleri\" & Form1.DataGrid1.Text & "\" & "[" & Date & " - " & Format(Time, "hh.nn") & "] " & Form1.DataGrid1.Text & ".xls")
End If

If "Orman Fakültesi" = Form1.DataGrid1.Text Then
ExcelNesne.SaveCopyAs ("C:\HastemTutanakGecmisleri\" & Form1.DataGrid1.Text & "\" & "[" & Date & " - " & Format(Time, "hh.nn") & "] " & Form1.DataGrid1.Text & ".xls")
End If


If "Otogar Lokantasý" = Form1.DataGrid1.Text Then
ExcelNesne.SaveCopyAs ("C:\HastemTutanakGecmisleri\" & Form1.DataGrid1.Text & "\" & "[" & Date & " - " & Format(Time, "hh.nn") & "] " & Form1.DataGrid1.Text & ".xls")
End If


If "SDÜ Ana Okulu" = Form1.DataGrid1.Text Then
ExcelNesne.SaveCopyAs ("C:\HastemTutanakGecmisleri\" & Form1.DataGrid1.Text & "\" & "[" & Date & " - " & Format(Time, "hh.nn") & "] " & Form1.DataGrid1.Text & ".xls")
End If


If "SDÜ Týp Fakültesi" = Form1.DataGrid1.Text Then
ExcelNesne.SaveCopyAs ("C:\HastemTutanakGecmisleri\" & Form1.DataGrid1.Text & "\" & "[" & Date & " - " & Format(Time, "hh.nn") & "] " & Form1.DataGrid1.Text & ".xls")
End If


If "Senirkent EML" = Form1.DataGrid1.Text Then
ExcelNesne.SaveCopyAs ("C:\HastemTutanakGecmisleri\" & Form1.DataGrid1.Text & "\" & "[" & Date & " - " & Format(Time, "hh.nn") & "] " & Form1.DataGrid1.Text & ".xls")
End If


If "Senirkent Ýmam Hatip Lisesi" = Form1.DataGrid1.Text Then
ExcelNesne.SaveCopyAs ("C:\HastemTutanakGecmisleri\" & Form1.DataGrid1.Text & "\" & "[" & Date & " - " & Format(Time, "hh.nn") & "] " & Form1.DataGrid1.Text & ".xls")
End If

If "Süt Ofis" = Form1.DataGrid1.Text Then
ExcelNesne.SaveCopyAs ("C:\HastemTutanakGecmisleri\" & Form1.DataGrid1.Text & "\" & "[" & Date & " - " & Format(Time, "hh.nn") & "] " & Form1.DataGrid1.Text & ".xls")
End If

If "Þenol Kimya" = Form1.DataGrid1.Text Then
ExcelNesne.SaveCopyAs ("C:\HastemTutanakGecmisleri\" & Form1.DataGrid1.Text & "\" & "[" & Date & " - " & Format(Time, "hh.nn") & "] " & Form1.DataGrid1.Text & ".xls")
End If

If "Teras Park - Ýgsiad" = Form1.DataGrid1.Text Then
ExcelNesne.SaveCopyAs ("C:\HastemTutanakGecmisleri\" & Form1.DataGrid1.Text & "\" & "[" & Date & " - " & Format(Time, "hh.nn") & "] " & Form1.DataGrid1.Text & ".xls")
End If

If "Tutaþ Gýda" = Form1.DataGrid1.Text Then
ExcelNesne.SaveCopyAs ("C:\HastemTutanakGecmisleri\" & Form1.DataGrid1.Text & "\" & "[" & Date & " - " & Format(Time, "hh.nn") & "] " & Form1.DataGrid1.Text & ".xls")
End If

If "Uluborlu Lisesi" = Form1.DataGrid1.Text Then
ExcelNesne.SaveCopyAs ("C:\HastemTutanakGecmisleri\" & Form1.DataGrid1.Text & "\" & "[" & Date & " - " & Format(Time, "hh.nn") & "] " & Form1.DataGrid1.Text & ".xls")
End If

If "Yalvaç Turizim Otelcilik" = Form1.DataGrid1.Text Then
ExcelNesne.SaveCopyAs ("C:\HastemTutanakGecmisleri\" & Form1.DataGrid1.Text & "\" & "[" & Date & " - " & Format(Time, "hh.nn") & "] " & Form1.DataGrid1.Text & ".xls")
End If

If "Yalvaç Türem Pansiyonu" = Form1.DataGrid1.Text Then
ExcelNesne.SaveCopyAs ("C:\HastemTutanakGecmisleri\" & Form1.DataGrid1.Text & "\" & "\" & "[" & Date & " - " & Format(Time, "hh.nn") & "] " & Form1.DataGrid1.Text & ".xls")
End If


If "Ziraat Fakültesi" = Form1.DataGrid1.Text Then
ExcelNesne.SaveCopyAs ("C:\HastemTutanakGecmisleri\" & Form1.DataGrid1.Text & "\" & "[" & Date & " - " & Format(Time, "hh.nn") & "] " & Form1.DataGrid1.Text & ".xls")
End If











MsgBox "Microsoft Excel'e Aktarildi Bekleniyor...", vbInformation, "Bildiri;"
End Sub
