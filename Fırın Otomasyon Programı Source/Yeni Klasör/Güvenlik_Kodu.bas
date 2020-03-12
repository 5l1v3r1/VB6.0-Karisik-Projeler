Attribute VB_Name = "Güvenlik_Kodu"
Private Function Güvenlik_Random_Kod()
Randomize
Dim harfler(3) As String
Dim sayýlar(3) As Integer
harfler(0) = Chr(Math.Round(Rnd() * (122 - 97) + 97))
harfler(1) = Chr(Math.Round(Rnd() * (122 - 97) + 97))
harfler(2) = Chr(Math.Round(Rnd() * (122 - 97) + 97))
sayýlar(0) = Math.Round(Rnd() * (9 - 0) + 0)
sayýlar(1) = Math.Round(Rnd() * (9 - 0) + 0)
sayýlar(2) = Math.Round(Rnd() * (9 - 0) + 0)
Form1.Label4.Caption = harfler(0) & harfler(1) & harfler(2) & sayýlar(0) & sayýlar(1) & sayýlar(2)
End Function
