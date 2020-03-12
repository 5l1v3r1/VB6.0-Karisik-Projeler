Attribute VB_Name = "Check_Göster_Gizle"
Public Function check_gizle() As Integer
Form1.CheckBox7.Enabled = False
Form1.CheckBox10.Enabled = False
Form1.CheckBox9.Enabled = False
Form1.CheckBox15.Enabled = False
Form1.CheckBox14.Enabled = False
Form1.CheckBox12.Enabled = False
Form1.CheckBox13.Enabled = False
Form1.CheckBox20.Enabled = False
Form1.CheckBox11.Enabled = False
Form1.CheckBox8.Enabled = False
Form1.CheckBox16.Enabled = False
End Function
Public Function check_göster() As Integer
Form1.CheckBox7.Enabled = True
Form1.CheckBox10.Enabled = True
Form1.CheckBox9.Enabled = True
Form1.CheckBox15.Enabled = True
Form1.CheckBox14.Enabled = True
Form1.CheckBox12.Enabled = True
Form1.CheckBox13.Enabled = True
Form1.CheckBox20.Enabled = True
Form1.CheckBox11.Enabled = True
Form1.CheckBox8.Enabled = True
Form1.CheckBox16.Enabled = True
End Function
