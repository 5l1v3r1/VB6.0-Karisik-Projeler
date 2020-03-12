VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Çocuðunuzun Ortalama Boyunu Bulun"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim AnneBoyu, BabaBoyu As String
Dim Cinsiyet As String
Dim Sonuc, Sonuc2, Sonuc3 As String
Dim KýzBoyOraný, ErkekBoyOraný As String
AnneBoyu = InputBox("Anne'nin boyu", Form1.Caption)
BabaBoyu = InputBox("Baba'nýn boyu", Form1.Caption)
Cinsiyet = InputBox("Çocuðun Cinsiyeti Nedir? Erkek/Kýz", "Deger3")
If Cinsiyet = "kýz" Then
Sonuc = Val(AnneBoyu) + Val(BabaBoyu)
Sonuc2 = Val(Sonuc) - Val("13")
Sonuc3 = Val(Sonuc2) / Val("2")
KýzBoyOraný = Sonuc3
MsgBox "Kýzýnýzýn Tahmini Boy Oraný : " & KýzBoyOraný & " cm", vbInformation, "Tahmin ;"
Show
Print "Annenin Boyu : " & AnneBoyu & " cm"
Print "Babanýn Boyu : " & BabaBoyu & " cm"
Print "Kýzýnýzýn Tahmini Boy Oraný : " & KýzBoyOraný & " cm"
End If
If Cinsiyet = "erkek" Then
Sonuc = Val(AnneBoyu) + Val(BabaBoyu)
Sonuc2 = Val(Sonuc) + Val("13")
Sonuc3 = Val(Sonuc2) / Val("2")
ErkekBoyOraný = Sonuc3
MsgBox "Oðlunuzun Tahmini Boy Oraný : " & ErkekBoyOraný & " cm", vbInformation, "Tahmin ;"
Show
Print "Annenin Boyu : " & AnneBoyu & " cm"
Print "Babanýn Boyu : " & BabaBoyu & " cm"
Print "Oðlunuzun Tahmini Boy Oraný : " & ErkekBoyOraný & " cm"
End If
End Sub
