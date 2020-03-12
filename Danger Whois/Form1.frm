VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13290
   LinkTopic       =   "Form1"
   ScaleHeight     =   7665
   ScaleWidth      =   13290
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text26 
      Height          =   285
      Left            =   12360
      Locked          =   -1  'True
      TabIndex        =   38
      Text            =   "Kýsa Sorgu Sonucu"
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox Text25 
      Height          =   285
      Left            =   12360
      Locked          =   -1  'True
      TabIndex        =   37
      Text            =   $"Form1.frx":0000
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox Text24 
      Height          =   285
      Left            =   12360
      Locked          =   -1  'True
      TabIndex        =   36
      Text            =   $"Form1.frx":002E
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox Text23 
      Height          =   285
      Left            =   12360
      Locked          =   -1  'True
      TabIndex        =   35
      Text            =   $"Form1.frx":0045
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox Text22 
      Height          =   285
      Left            =   12360
      Locked          =   -1  'True
      TabIndex        =   34
      Text            =   $"Form1.frx":0061
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox Text21 
      Height          =   285
      Left            =   12360
      Locked          =   -1  'True
      TabIndex        =   33
      Text            =   $"Form1.frx":007D
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox Text20 
      Height          =   285
      Left            =   12360
      Locked          =   -1  'True
      TabIndex        =   32
      Text            =   $"Form1.frx":0095
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox Text19 
      Height          =   285
      Left            =   12360
      Locked          =   -1  'True
      TabIndex        =   31
      Text            =   $"Form1.frx":00AD
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Text18 
      Height          =   285
      Left            =   11280
      Locked          =   -1  'True
      TabIndex        =   30
      Text            =   $"Form1.frx":00BC
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox Text17 
      Height          =   285
      Left            =   11280
      Locked          =   -1  'True
      TabIndex        =   29
      Text            =   "Son Güncelleme Tarihi : "
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox Text16 
      Height          =   285
      Left            =   11280
      Locked          =   -1  'True
      TabIndex        =   28
      Text            =   $"Form1.frx":00CB
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox Text15 
      Height          =   285
      Left            =   11280
      Locked          =   -1  'True
      TabIndex        =   27
      Text            =   $"Form1.frx":00E9
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox Text14 
      Height          =   285
      Left            =   11280
      Locked          =   -1  'True
      TabIndex        =   26
      Text            =   $"Form1.frx":00FE
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox Text13 
      Height          =   285
      Left            =   11280
      Locked          =   -1  'True
      TabIndex        =   25
      Text            =   $"Form1.frx":0113
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   11280
      Locked          =   -1  'True
      TabIndex        =   24
      Text            =   $"Form1.frx":012F
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   11280
      Locked          =   -1  'True
      TabIndex        =   23
      Text            =   "Sorgulanan Alan Adý : "
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "göster"
      Height          =   315
      Left            =   11280
      TabIndex        =   8
      Top             =   2520
      Width           =   1575
   End
   Begin SHDocVwCtl.WebBrowser Web1 
      Height          =   855
      Left            =   7920
      TabIndex        =   7
      Top             =   120
      Width           =   3255
      ExtentX         =   5741
      ExtentY         =   1508
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.TextBox Text2 
      Height          =   1575
      Left            =   7920
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Text            =   "Form1.frx":014A
      Top             =   960
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      Height          =   5535
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   6855
      Begin VB.Frame Frame2 
         Caption         =   "-Whois Sonucu"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   120
         TabIndex        =   21
         Top             =   2400
         Width           =   6615
         Begin VB.TextBox Text10 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2655
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   22
            Top             =   240
            Width           =   6375
         End
      End
      Begin VB.TextBox Text9 
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
         Left            =   4680
         TabIndex        =   20
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   1680
         Width           =   4335
      End
      Begin VB.TextBox Text7 
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
         Left            =   2400
         TabIndex        =   16
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox Text6 
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
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox Text5 
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
         Left            =   4680
         TabIndex        =   11
         Top             =   480
         Width           =   2055
      End
      Begin VB.TextBox Text4 
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
         Left            =   2400
         TabIndex        =   10
         Top             =   480
         Width           =   2055
      End
      Begin VB.TextBox Text3 
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
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Kimindir Sunucusu"
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
         Left            =   4680
         TabIndex        =   18
         Top             =   840
         Width           =   1590
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Ýsim Sunucularý"
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
         TabIndex        =   17
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Durumu"
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
         Left            =   2400
         TabIndex        =   15
         Top             =   840
         Width           =   690
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Son Güncelleme Tarihi"
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
         TabIndex        =   14
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Bitiþ Tarihi"
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
         Left            =   4680
         TabIndex        =   13
         Top             =   240
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Oluþturulma Tarihi"
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
         Left            =   2400
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Sorgulanan Alan Adý"
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
         TabIndex        =   4
         Top             =   240
         Width           =   1740
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sorgula"
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
      Left            =   3360
      TabIndex        =   2
      Top             =   120
      Width           =   1095
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
      Left            =   720
      TabIndex        =   1
      Text            =   "cyber-warrior.org"
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Site :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Web1.Navigate ("http://kimindir.com/" & Text1.Text)
End Sub

Private Sub Command2_Click()
Text2.Text = Web1.Document.body.innerText
End Sub

Private Sub Text2_Change()
SorgulananAlanAdý
OlusturmaTarihi
BitisTarihi
SonGuncellemeTarihi
Durumu
ÝsimSunucularý
KimindirSunucusu
Whois
If InStr(1, Text2.Text, "alan adý kimse tarafýndan alýnmamýþ.") Then
MsgBox "[ " & Text1.Text & " ] alan adý kimse tarafýndan alýnmamýþ.", vbExclamation, "Uyarý;"
End If

End Sub

Private Function Whois()
On Error Resume Next
Dim AranacakYer, aranan As String
Dim ilk, son
AranacakYer = Text2.Text 'Verinin Aranacaðý Yer
deg1 = deg1 & AranacakYer
aranan = Text25.Text  '1. Deger
ilk = InStr(1, deg1, aranan) + Len(aranan)
If ilk <> Len(aranan) Then
son = InStr(ilk, deg1, Text26.Text) - ilk  ' " "  Arasýndaki 2. Deger
Text10.Text = Mid(deg1, ilk, son) 'Degerlerin arasýndaki veri
Exit Function
End If
End Function

Private Function KimindirSunucusu()
On Error Resume Next
Dim AranacakYer, aranan As String
Dim ilk, son
AranacakYer = Text2.Text 'Verinin Aranacaðý Yer
deg1 = deg1 & AranacakYer
aranan = Text23.Text  '1. Deger
ilk = InStr(1, deg1, aranan) + Len(aranan)
If ilk <> Len(aranan) Then
son = InStr(ilk, deg1, Text24.Text) - ilk  ' " "  Arasýndaki 2. Deger
Text9.Text = Mid(deg1, ilk, son) 'Degerlerin arasýndaki veri
Exit Function
End If
End Function

Private Function ÝsimSunucularý()
On Error Resume Next
Dim AranacakYer, aranan As String
Dim ilk, son
AranacakYer = Text2.Text 'Verinin Aranacaðý Yer
deg1 = deg1 & AranacakYer
aranan = Text21.Text  '1. Deger
ilk = InStr(1, deg1, aranan) + Len(aranan)
If ilk <> Len(aranan) Then
son = InStr(ilk, deg1, Text22.Text) - ilk  ' " "  Arasýndaki 2. Deger
Text8.Text = Mid(deg1, ilk, son) 'Degerlerin arasýndaki veri
Exit Function
End If
End Function

Private Function Durumu()
On Error Resume Next
Dim AranacakYer, aranan As String
Dim ilk, son
AranacakYer = Text2.Text 'Verinin Aranacaðý Yer
deg1 = deg1 & AranacakYer
aranan = Text19.Text  '1. Deger
ilk = InStr(1, deg1, aranan) + Len(aranan)
If ilk <> Len(aranan) Then
son = InStr(ilk, deg1, Text20.Text) - ilk  ' " "  Arasýndaki 2. Deger
Text7.Text = Mid(deg1, ilk, son) 'Degerlerin arasýndaki veri
Exit Function
End If
End Function

Private Function SonGuncellemeTarihi()
On Error Resume Next
Dim AranacakYer, aranan As String
Dim ilk, son
AranacakYer = Text2.Text 'Verinin Aranacaðý Yer
deg1 = deg1 & AranacakYer
aranan = Text17.Text  '1. Deger
ilk = InStr(1, deg1, aranan) + Len(aranan)
If ilk <> Len(aranan) Then
son = InStr(ilk, deg1, Text18.Text) - ilk  ' " "  Arasýndaki 2. Deger
Text6.Text = Mid(deg1, ilk, son) 'Degerlerin arasýndaki veri
Exit Function
End If
End Function

Private Function BitisTarihi()
On Error Resume Next
Dim AranacakYer, aranan As String
Dim ilk, son
AranacakYer = Text2.Text 'Verinin Aranacaðý Yer
deg1 = deg1 & AranacakYer
aranan = Text15.Text  '1. Deger
ilk = InStr(1, deg1, aranan) + Len(aranan)
If ilk <> Len(aranan) Then
son = InStr(ilk, deg1, Text16.Text) - ilk  ' " "  Arasýndaki 2. Deger
Text5.Text = Mid(deg1, ilk, son) 'Degerlerin arasýndaki veri
Exit Function
End If
End Function

Private Function OlusturmaTarihi()
On Error Resume Next
Dim AranacakYer, aranan As String
Dim ilk, son
AranacakYer = Text2.Text 'Verinin Aranacaðý Yer
deg1 = deg1 & AranacakYer
aranan = Text13.Text  '1. Deger
ilk = InStr(1, deg1, aranan) + Len(aranan)
If ilk <> Len(aranan) Then
son = InStr(ilk, deg1, Text14.Text) - ilk  ' " "  Arasýndaki 2. Deger
Text4.Text = Mid(deg1, ilk, son) 'Degerlerin arasýndaki veri
Exit Function
End If
End Function

Private Function SorgulananAlanAdý()
On Error Resume Next
Dim AranacakYer, aranan As String
Dim ilk, son
AranacakYer = Text2.Text 'Verinin Aranacaðý Yer
deg1 = deg1 & AranacakYer
aranan = Text11.Text  '1. Deger
ilk = InStr(1, deg1, aranan) + Len(aranan)
If ilk <> Len(aranan) Then
son = InStr(ilk, deg1, Text12.Text) - ilk  ' " "  Arasýndaki 2. Deger
Text3.Text = Mid(deg1, ilk, son) 'Degerlerin arasýndaki veri
Exit Function
End If
End Function


Private Sub Web1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
Text2.Text = Web1.Document.body.innerText

End Sub

