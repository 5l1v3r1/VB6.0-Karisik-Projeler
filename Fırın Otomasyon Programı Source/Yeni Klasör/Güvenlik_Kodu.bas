Attribute VB_Name = "G�venlik_Kodu"
Private Function G�venlik_Random_Kod()
Randomize
Dim harfler(3) As String
Dim say�lar(3) As Integer
harfler(0) = Chr(Math.Round(Rnd() * (122 - 97) + 97))
harfler(1) = Chr(Math.Round(Rnd() * (122 - 97) + 97))
harfler(2) = Chr(Math.Round(Rnd() * (122 - 97) + 97))
say�lar(0) = Math.Round(Rnd() * (9 - 0) + 0)
say�lar(1) = Math.Round(Rnd() * (9 - 0) + 0)
say�lar(2) = Math.Round(Rnd() * (9 - 0) + 0)
Form1.Label4.Caption = harfler(0) & harfler(1) & harfler(2) & say�lar(0) & say�lar(1) & say�lar(2)
End Function
