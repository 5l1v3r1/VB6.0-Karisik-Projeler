Attribute VB_Name = "Module1"
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const REG_SZ = 1
Public Sub RegKaydiYaz(hKey As Long, Anahtar As String, DegerAdi As String, Deger As String)
Dim Ac 'Oluþturulacak anahtarýn adresi
RegCreateKey hKey, Anahtar, Ac 'Anahtarý oluþturduk
RegSetValueEx Ac, DegerAdi, 0, REG_SZ, ByVal Deger, Len(Deger) 'Anahtarýmýzýn "DegerAdi" isimli deðerine "Deger" parametresi ile gelen String deðeri atadýk.
RegCloseKey Ac 'Ve açtýðýmýz anahtarý kapattýk.
End Sub
