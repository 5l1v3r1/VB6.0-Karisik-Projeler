VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "�ocu�unuzun Ortalama Boyunu Bulun"
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
Dim K�zBoyOran�, ErkekBoyOran� As String
AnneBoyu = InputBox("Anne'nin boyu", Form1.Caption)
BabaBoyu = InputBox("Baba'n�n boyu", Form1.Caption)
Cinsiyet = InputBox("�ocu�un Cinsiyeti Nedir? Erkek/K�z", "Deger3")
If Cinsiyet = "k�z" Then
Sonuc = Val(AnneBoyu) + Val(BabaBoyu)
Sonuc2 = Val(Sonuc) - Val("13")
Sonuc3 = Val(Sonuc2) / Val("2")
K�zBoyOran� = Sonuc3
MsgBox "K�z�n�z�n Tahmini Boy Oran� : " & K�zBoyOran� & " cm", vbInformation, "Tahmin ;"
Show
Print "Annenin Boyu : " & AnneBoyu & " cm"
Print "Baban�n Boyu : " & BabaBoyu & " cm"
Print "K�z�n�z�n Tahmini Boy Oran� : " & K�zBoyOran� & " cm"
End If
If Cinsiyet = "erkek" Then
Sonuc = Val(AnneBoyu) + Val(BabaBoyu)
Sonuc2 = Val(Sonuc) + Val("13")
Sonuc3 = Val(Sonuc2) / Val("2")
ErkekBoyOran� = Sonuc3
MsgBox "O�lunuzun Tahmini Boy Oran� : " & ErkekBoyOran� & " cm", vbInformation, "Tahmin ;"
Show
Print "Annenin Boyu : " & AnneBoyu & " cm"
Print "Baban�n Boyu : " & BabaBoyu & " cm"
Print "O�lunuzun Tahmini Boy Oran� : " & ErkekBoyOran� & " cm"
End If
End Sub
