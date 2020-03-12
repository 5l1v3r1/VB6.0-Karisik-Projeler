VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "ActiveX.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   Caption         =   "Danger AdminPanel Scanner"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11565
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   11565
   StartUpPosition =   3  'Windows Default
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   8760
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Durdur"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   11
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Baþlat"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   480
      Top             =   5160
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4560
      TabIndex        =   7
      Text            =   "www.udesa.co.za"
      Top             =   120
      Width           =   2895
   End
   Begin VB.Frame Frame2 
      Height          =   3495
      Left            =   3840
      TabIndex        =   5
      Top             =   360
      Width           =   3615
      Begin VB.TextBox Text3 
         Height          =   2775
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Text            =   "Form1.frx":0000
         Top             =   600
         Width           =   3375
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1440
         TabIndex        =   9
         Text            =   "Text2"
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label312 
         AutoSize        =   -1  'True
         Caption         =   "Admin Panel :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1200
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "-Admin List"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      Begin VB.CommandButton Command2 
         Caption         =   "Listeyi Temizle"
         Height          =   255
         Left            =   1920
         TabIndex        =   4
         Top             =   3480
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Admin List Ekle"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   3480
         Width           =   1695
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2790
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3495
      End
      Begin XtremeSuiteControls.CommonDialog Cmd1 
         Left            =   2760
         Top             =   3960
         _Version        =   851968
         _ExtentX        =   423
         _ExtentY        =   423
         _StockProps     =   4
      End
      Begin VB.Label Label1 
         Caption         =   "Toplam Dork:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   3120
         Width           =   3495
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Adres :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3840
      TabIndex        =   6
      Top             =   120
      Width           =   630
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'* Dosya Seçme Penceresi Açar
Cmd1.ShowOpen
'* List1'i Tamamen Temizler
List1.Clear
'* Ýf Döngüsü Baþlangýç
If Dir(Cmd1.FileName) <> "" Then
If Cmd1.FileName = "" Then
'* Dosya Seçilmezse Bilgi Mesajý verir
MsgBox "Dosya Seçilmedi", vbInformation, "Uyarý ;"
'* Taramayý Baþlat Butonunu Aktifleþtirir
'* Deðilse
Else
'* Seçilen Dosyayý List1'e aktarýr
Open Cmd1.FileName For Input As #1
While Not EOF(1)
Input #1, a
List1.AddItem a
Wend
Close #1
'* Label'e Listbox'daki Toplam Deðeri Yansýtýr
Label1.Caption = "Toplam Dork: " & List1.ListCount
'* Taramayý Baþlat Butonunu Aktifleþtirir
End If
'* Ýf Döngüsü Bitiþ
End If
End Sub

Private Sub Command2_Click()
Command1.Enabled = True
List1.Clear
End Sub

Private Sub Command3_Click()
Timer1.Enabled = True
End Sub

Private Sub Command4_Click()
Timer1.Enabled = False
End Sub

Private Sub Form_Load()
Cmd1.DialogTitle = "Kaydedilecek Yer"
Cmd1.Filter = ".txt Dosyasý|*.txt"
If Dir(App.Path & "\AdminList.txt") <> "" Then
Command1.Enabled = False
Open App.Path & "\AdminList.txt" For Input As #1
While Not EOF(1)
Input #1, a
List1.AddItem a
'* Label'e Listbox'daki Toplam Deðeri Yansýtýr
Label1.Caption = "Toplam Dork: " & List1.ListCount
Wend
Close #1
Else
Command1.Enabled = True
MsgBox "AdminList.txt Dosyasý Dizinde Bulunamadý...", vbInformation, "Hata ;"
End If
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
With List1
'*Ýndex sayýsýný 1 arttýrýr
.ListIndex = .ListIndex + 1
'*Seçili elemanýn deðeri
Text2.Text = .List(.ListIndex)
'WebBrowser1.Navigate ("Http://" & Text1.Text & "/" & Text2.Text)
Text3.Text = Inet1.OpenURL("Http://" & Text1.Text & "/" & Text2.Text)
'* Listenin son elemanýna gelince bilgi mesajý verir
If .ListIndex = .ListCount - 1 Then
.ListIndex = -1
'* Timer1'i Durdurur
Timer1.Enabled = False
End If
End With
End Sub
