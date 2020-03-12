VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "ActiveX.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12090
   LinkTopic       =   "Form1"
   ScaleHeight     =   8535
   ScaleWidth      =   12090
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text23 
      Height          =   285
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   37
      Text            =   "</PRE></TD></TR></TBODY></TABLE></DIV></TD></TR></TBODY></TABLE><!-- js sekme-->"
      Top             =   5880
      Width           =   1815
   End
   Begin VB.TextBox Text22 
      Height          =   285
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   36
      Text            =   "<TD id=jx_ipwhrec_pre class=""ptd full"" colSpan=2><PRE>"
      Top             =   5640
      Width           =   1815
   End
   Begin VB.TextBox Text20 
      Height          =   285
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   32
      Text            =   $"Form1.frx":0000
      Top             =   5280
      Width           =   1815
   End
   Begin VB.TextBox Text19 
      Height          =   285
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   31
      Text            =   ".gif""> "
      Top             =   7080
      Width           =   1815
   End
   Begin VB.TextBox Text18 
      Height          =   285
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   30
      Text            =   $"Form1.frx":0020
      Top             =   6840
      Width           =   1815
   End
   Begin VB.TextBox Text17 
      Height          =   285
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   29
      Text            =   $"Form1.frx":0054
      Top             =   5880
      Width           =   1815
   End
   Begin VB.TextBox Text16 
      Height          =   285
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   28
      Text            =   $"Form1.frx":006F
      Top             =   5640
      Width           =   1815
   End
   Begin VB.TextBox Text15 
      Height          =   285
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   27
      Text            =   $"Form1.frx":009D
      Top             =   7440
      Width           =   1815
   End
   Begin VB.TextBox Text14 
      Height          =   285
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   26
      Text            =   $"Form1.frx":00BB
      Top             =   6480
      Width           =   1815
   End
   Begin VB.TextBox Text13 
      Height          =   285
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   25
      Text            =   $"Form1.frx":00D9
      Top             =   6240
      Width           =   1815
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   24
      Text            =   $"Form1.frx":0104
      Top             =   5280
      Width           =   1815
   End
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   23
      Text            =   $"Form1.frx":0122
      Top             =   5040
      Width           =   1815
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   $"Form1.frx":0147
      Top             =   5040
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Bilgileri Göster"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   480
      Top             =   7680
   End
   Begin XtremeSuiteControls.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   6240
      Width           =   7455
      _Version        =   851968
      _ExtentX        =   13150
      _ExtentY        =   450
      _StockProps     =   93
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.TextBox Text4 
      Height          =   2535
      Left            =   7680
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   11
      Text            =   "Form1.frx":019B
      Top             =   2400
      Width           =   4335
   End
   Begin XtremeSuiteControls.WebBrowser WebBrowser1 
      Height          =   2415
      Left            =   7680
      TabIndex        =   10
      Top             =   0
      Width           =   4335
      _Version        =   851968
      _ExtentX        =   7646
      _ExtentY        =   4260
      _StockProps     =   173
      BackColor       =   -2147483643
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
      Left            =   3480
      TabIndex        =   9
      Top             =   120
      Width           =   855
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
      Left            =   600
      TabIndex        =   8
      Text            =   "atem.k12.tr"
      Top             =   120
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   7455
      Begin XtremeSuiteControls.TabControl TabControl1 
         CausesValidation=   0   'False
         Height          =   3495
         Left            =   120
         TabIndex        =   33
         Top             =   2040
         Width           =   7215
         _Version        =   851968
         _ExtentX        =   12726
         _ExtentY        =   6165
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Color           =   8
         ItemCount       =   2
         Item(0).Caption =   "Registry Whois Kaydý"
         Item(0).ControlCount=   1
         Item(0).Control(0)=   "Text6"
         Item(1).Caption =   "IP Whois Kaydý"
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "Text21"
         Begin VB.TextBox Text21 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2895
            Left            =   -69880
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   35
            Top             =   480
            Visible         =   0   'False
            Width           =   6975
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
            Height          =   3015
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   34
            Top             =   360
            Width           =   6975
         End
      End
      Begin XtremeSuiteControls.ListBox ListBox1 
         Height          =   1500
         Left            =   2400
         TabIndex        =   22
         Top             =   480
         Width           =   2655
         _Version        =   851968
         _ExtentX        =   4683
         _ExtentY        =   2646
         _StockProps     =   77
         BackColor       =   -2147483643
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
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
         Height          =   285
         Left            =   5160
         TabIndex        =   20
         Top             =   1680
         Width           =   2175
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
         Left            =   120
         TabIndex        =   18
         Top             =   1680
         Width           =   2175
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
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   2175
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
         Left            =   5160
         TabIndex        =   6
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox Text2 
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
         Left            =   5160
         TabIndex        =   5
         Top             =   480
         Width           =   2175
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
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "NS Sunucusu"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2400
         TabIndex        =   21
         Top             =   240
         Width           =   1245
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Ülke"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5160
         TabIndex        =   19
         Top             =   1440
         Width           =   435
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Oluþturma Tarihi"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   1440
         Width           =   1635
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Bitiþ Tarihi"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   840
         Width           =   1050
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Kayýtlý Whois Sunucu"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5160
         TabIndex        =   3
         Top             =   840
         Width           =   2025
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "IP Adresi"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5160
         TabIndex        =   2
         Top             =   240
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Durum"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   645
      End
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Site :"
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
      TabIndex        =   7
      Top             =   120
      Width           =   465
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
WebBrowser1.Navigate ("http://www.whois.com.tr/?q=" & Text5.Text)
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text6.Text = ""
Text8.Text = ""
Text9.Text = ""
Text21.Text = ""
End Sub
Private Function SiteKodlarý()
On Error Resume Next
Text4.Text = WebBrowser1.Document.body.InnerHTML
End Function
Private Sub Command3_Click()
Durum
BitiþTarih
OluþturmaTarih
'NSsunucu
IPadress
KayýtlýWhoisSunucu
Ülke
RegistryWhoisKayýt
IPWhoisKayýt
End Sub
Private Function IPWhoisKayýt()
On Error Resume Next
Dim AranacakYer, aranan As String
Dim ilk, son
AranacakYer = Text4.Text
deg1 = deg1 & AranacakYer
aranan = Text22.Text
ilk = InStr(1, deg1, aranan) + Len(aranan)
If ilk <> Len(aranan) Then
son = InStr(ilk, deg1, Text23.Text) - ilk
Text21.Text = Mid(deg1, ilk, son)
Exit Function
End If
End Function
Private Function RegistryWhoisKayýt()
On Error Resume Next
Dim AranacakYer, aranan As String
Dim ilk, son
AranacakYer = Text4.Text
deg1 = deg1 & AranacakYer
aranan = Text7.Text
ilk = InStr(1, deg1, aranan) + Len(aranan)
If ilk <> Len(aranan) Then
son = InStr(ilk, deg1, Text20.Text) - ilk
Text6.Text = Mid(deg1, ilk, son)
Exit Function
End If
End Function
Private Function Ülke()
On Error Resume Next
Dim AranacakYer, aranan As String
Dim ilk, son
AranacakYer = Text4.Text 'Verinin Aranacaðý Yer
deg1 = deg1 & AranacakYer
aranan = Text19.Text
ilk = InStr(1, deg1, aranan) + Len(aranan)
If ilk <> Len(aranan) Then
son = InStr(ilk, deg1, Text18.Text) - ilk  ' " "  Arasýndaki 2. Deger
Text10.Text = Mid(deg1, ilk, son) 'Degerlerin arasýndaki veri
Exit Function
End If
End Function
Private Function KayýtlýWhoisSunucu()
On Error Resume Next
Dim AranacakYer, aranan As String
Dim ilk, son
AranacakYer = Text4.Text 'Verinin Aranacaðý Yer
deg1 = deg1 & AranacakYer
aranan = Text16.Text  '1. Deger
ilk = InStr(1, deg1, aranan) + Len(aranan)
If ilk <> Len(aranan) Then
son = InStr(ilk, deg1, Text17.Text) - ilk  ' " "  Arasýndaki 2. Deger
Text3.Text = Mid(deg1, ilk, son) 'Degerlerin arasýndaki veri
Exit Function
End If
End Function
Private Function IPadress()
On Error Resume Next
Dim AranacakYer, aranan As String
Dim ilk, son
AranacakYer = Text4.Text 'Verinin Aranacaðý Yer
deg1 = deg1 & AranacakYer
aranan = "<TD id=jx_ptd_ipup class=ptd>" '1. Deger
ilk = InStr(1, deg1, aranan) + Len(aranan)
If ilk <> Len(aranan) Then
son = InStr(ilk, deg1, Text15.Text) - ilk  ' " "  Arasýndaki 2. Deger
Text2.Text = Mid(deg1, ilk, son) 'Degerlerin arasýndaki veri
Exit Function
End If
End Function
Private Function NSsunucu()
On Error Resume Next
Dim AranacakYer, aranan As String
Dim ilk, son
AranacakYer = Text4.Text 'Verinin Aranacaðý Yer
deg1 = deg1 & AranacakYer
aranan = Text13.Text  '1. Deger
ilk = InStr(1, deg1, aranan) + Len(aranan)
If ilk <> Len(aranan) Then
son = InStr(ilk, deg1, Text14.Text) - ilk  ' " "  Arasýndaki 2. Deger
Text9.Text = Mid(deg1, ilk, son) 'Degerlerin arasýndaki veri
Exit Function
End If
End Function
Private Function OluþturmaTarih()
On Error Resume Next
Dim AranacakYer, aranan As String
Dim ilk, son
AranacakYer = Text4.Text 'Verinin Aranacaðý Yer
deg1 = deg1 & AranacakYer
aranan = Text13.Text  '1. Deger
ilk = InStr(1, deg1, aranan) + Len(aranan)
If ilk <> Len(aranan) Then
son = InStr(ilk, deg1, Text14.Text) - ilk  ' " "  Arasýndaki 2. Deger
Text9.Text = Mid(deg1, ilk, son) 'Degerlerin arasýndaki veri
Exit Function
End If
End Function
Private Function BitiþTarih()
On Error Resume Next
Dim AranacakYer, aranan As String
Dim ilk, son
AranacakYer = Text4.Text 'Verinin Aranacaðý Yer
deg1 = deg1 & AranacakYer
aranan = Text11.Text  '1. Deger
ilk = InStr(1, deg1, aranan) + Len(aranan)
If ilk <> Len(aranan) Then
son = InStr(ilk, deg1, Text12.Text) - ilk  ' " "  Arasýndaki 2. Deger
Text8.Text = Mid(deg1, ilk, son) 'Degerlerin arasýndaki veri
Exit Function
End If
End Function
Private Function Durum()
On Error Resume Next
Dim AranacakYer, aranan As String
Dim ilk, son
AranacakYer = Text4.Text 'Verinin Aranacaðý Yer
deg1 = deg1 & AranacakYer
aranan = "<P class=dom_not_avail>" '1. Deger
ilk = InStr(1, deg1, aranan) + Len(aranan)
If ilk <> Len(aranan) Then
son = InStr(ilk, deg1, "</P></TD></TR>") - ilk ' " "  Arasýndaki 2. Deger
Text1.Text = Mid(deg1, ilk, son) 'Degerlerin arasýndaki veri
Exit Function
End If
End Function



Private Sub Text4_Change()
Command3.Enabled = True
Durum
BitiþTarih
OluþturmaTarih
'NSsunucu
IPadress
KayýtlýWhoisSunucu
Ülke
RegistryWhoisKayýt
IPWhoisKayýt
End Sub

Private Sub Timer1_Timer()
  SiteKodlarý
  Timer1.Enabled = False
End Sub

Private Sub webBrowser1_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
        On Error Resume Next
        If Progress = -1 Then ProgressBar1.Value = 120
       Form1.Caption = "Done"
       Timer1.Enabled = True
        ProgressBar1.Visible = False
        If Progress > 0 And ProgressMax > 0 Then
            ProgressBar1.Visible = True
            ProgressBar1.Value = Progress * 100 / ProgressMax
             Form1.Caption = Int(Progress * 100 / ProgressMax) & "%"
        End If
    End Sub
