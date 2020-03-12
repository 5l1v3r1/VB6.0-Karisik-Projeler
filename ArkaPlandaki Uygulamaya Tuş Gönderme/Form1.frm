VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3630
   LinkTopic       =   "Form1"
   ScaleHeight     =   1530
   ScaleWidth      =   3630
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, _
    ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Const WM_KEYDOWN = &H100
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Private Sub Command1_Click()
Dim a, b
a = FindWindow("WordPad", vbNullString)
'Buradaki Notepad o pencerenin/uygulamanin class ismidir.
'Spy++ ile class ismini bulabilirsiniz.
'Class ismini bulmaniz daha iyi olur cunku pencere ismi degisebilir.
'Ancak siz yine de pencere ismini kullanmak istiyorsaniz, vbnullstring ile
'"Notepad" i yer degistirin ve programinizin pencere ismini "Notepad" bolumune yazin.

b = FindWindowEx(a, ByVal 0&, "Edit", vbNullString)
' Burasi uygulamanin icindeki her hangi bir kontrolun HWND sini bulmaya yarar.
'Tam anlamiyla verim almak icin kullanilmalidir.
'Buradaki "Edit" kontrolun class ismidir. Siz bunu Spy++ ile bulacaksiniz.

PostMessage b, WM_KEYDOWN, vbKeyZ, 0
' Evet api cagriliyor, buradaki vbKeyZ Z tusunun hex degerini belirtir,
'farkli tus kullanicaksaniz farkli tusu yazin ornegin G vbKeyG olabilir burasi.
'Eger vbkeyX serisinde belirtilmemisse oraya basina &H koyarak hex degerini girin.
'Microsoft un sitesinde hex degerleri var.
End Sub

