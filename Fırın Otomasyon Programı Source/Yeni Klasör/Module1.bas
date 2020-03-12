Attribute VB_Name = "Module1"
Global dbConn As ADODB.Connection
Global kayit As ADODB.Recordset
Global strpass As String

Sub Main()
Set dbConn = New ADODB.Connection
Set Db = New ADODB.Command
Set kayit = New ADODB.Recordset

strpass = "holocaust32"
With dbConn
.Provider = "Microsoft.Jet.OLEDB.4.0"
.Properties("Jet OLEDB:Database Password") = strpass
.Mode = adModeReadWrite
.Open App.Path & "\Database.mdb"
End With
Form3.Show 'Ýlk ekrana hangi formun açýlmasýný istiyorsanýz onu yazýn
End Sub
