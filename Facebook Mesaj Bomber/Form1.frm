VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "ActiveX.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Danger facebook Mesaj Bomber // _DangerOusMaN_"
   ClientHeight    =   9675
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   8955
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9675
   ScaleWidth      =   8955
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.PushButton PushButton2 
      Height          =   255
      Left            =   4320
      TabIndex        =   24
      Top             =   2520
      Width           =   2055
      _Version        =   851968
      _ExtentX        =   3625
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "ID Numarasý Bulma"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   975
      Left            =   360
      TabIndex        =   21
      Top             =   3600
      Width           =   1695
      _Version        =   851968
      _ExtentX        =   2990
      _ExtentY        =   1720
      _StockProps     =   79
      Caption         =   "[ Sayaç ]"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   4
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Mesaj Gönderildi..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   600
         TabIndex        =   23
         Top             =   360
         Width           =   1050
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   180
      End
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit1 
      Height          =   255
      Left            =   2160
      TabIndex        =   20
      Top             =   3000
      Width           =   4215
      _Version        =   851968
      _ExtentX        =   7435
      _ExtentY        =   450
      _StockProps     =   77
      ForeColor       =   4210752
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      Text            =   "Baþlýk"
      Appearance      =   4
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.WebBrowser WebBrowser2 
      Height          =   3255
      Left            =   3360
      TabIndex        =   18
      Top             =   6000
      Width           =   3255
      _Version        =   851968
      _ExtentX        =   5741
      _ExtentY        =   5741
      _StockProps     =   173
      BackColor       =   -2147483643
   End
   Begin XtremeSuiteControls.WebBrowser WebBrowser1 
      Height          =   3255
      Left            =   120
      TabIndex        =   17
      Top             =   6000
      Width           =   3255
      _Version        =   851968
      _ExtentX        =   5741
      _ExtentY        =   5741
      _StockProps     =   173
      BackColor       =   -2147483643
   End
   Begin XtremeSuiteControls.PushButton PushButton8 
      Height          =   375
      Left            =   4920
      TabIndex        =   16
      Top             =   5280
      Width           =   1335
      _Version        =   851968
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Çýkýþ"
      BackColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton PushButton7 
      Height          =   375
      Left            =   3480
      TabIndex        =   15
      Top             =   5280
      Width           =   1335
      _Version        =   851968
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Hakkýnda"
      BackColor       =   65535
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton PushButton6 
      Height          =   375
      Left            =   1920
      TabIndex        =   14
      Top             =   5280
      Width           =   1335
      _Version        =   851968
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Durdur"
      BackColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton PushButton5 
      Height          =   375
      Left            =   480
      TabIndex        =   13
      Top             =   5280
      Width           =   1335
      _Version        =   851968
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Baþlat"
      BackColor       =   65280
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      UseVisualStyle  =   -1  'True
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1440
      Top             =   3120
   End
   Begin XtremeSuiteControls.PushButton PushButton3 
      Height          =   255
      Left            =   3000
      TabIndex        =   12
      Top             =   4800
      Width           =   975
      _Version        =   851968
      _ExtentX        =   1720
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Deðiþtir"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit Text3 
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   2520
      Width           =   2055
      _Version        =   851968
      _ExtentX        =   3625
      _ExtentY        =   450
      _StockProps     =   77
      ForeColor       =   255
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      Appearance      =   4
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   615
      Left            =   5400
      TabIndex        =   6
      Top             =   1680
      Width           =   975
      _Version        =   851968
      _ExtentX        =   1720
      _ExtentY        =   1085
      _StockProps     =   79
      Caption         =   "Giriþ Yap"
      BackColor       =   -2147483635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.CheckBox Check1 
      Height          =   255
      Left            =   4320
      TabIndex        =   5
      Top             =   2040
      Width           =   975
      _Version        =   851968
      _ExtentX        =   1720
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Göster"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
   End
   Begin XtremeSuiteControls.FlatEdit Text2 
      Height          =   255
      Left            =   2160
      TabIndex        =   4
      Top             =   2040
      Width           =   2055
      _Version        =   851968
      _ExtentX        =   3625
      _ExtentY        =   450
      _StockProps     =   77
      ForeColor       =   16711680
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PasswordChar    =   "*"
      Appearance      =   4
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit Text1 
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   1680
      Width           =   3135
      _Version        =   851968
      _ExtentX        =   5530
      _ExtentY        =   450
      _StockProps     =   77
      ForeColor       =   16711680
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   4
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit Text4 
      Height          =   1215
      Left            =   2160
      TabIndex        =   9
      Top             =   3360
      Width           =   4215
      _Version        =   851968
      _ExtentX        =   7435
      _ExtentY        =   2143
      _StockProps     =   77
      ForeColor       =   4210752
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      Text            =   "Mesaj"
      MultiLine       =   -1  'True
      ScrollBars      =   3
      Appearance      =   4
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit Text5 
      Height          =   255
      Left            =   2160
      TabIndex        =   11
      Top             =   4800
      Width           =   735
      _Version        =   851968
      _ExtentX        =   1296
      _ExtentY        =   450
      _StockProps     =   77
      ForeColor       =   12632256
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      Text            =   "1"
      Alignment       =   2
      MaxLength       =   5
      Appearance      =   4
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.Label Label9 
      Height          =   255
      Left            =   4080
      TabIndex        =   25
      Top             =   4800
      Width           =   2415
      _Version        =   851968
      _ExtentX        =   4260
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Saniye Girmeden Baþlatamazsýnýz."
      ForeColor       =   16777215
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Konu Baþlýðý               :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   360
      TabIndex        =   19
      Top             =   3000
      Width           =   1725
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   240
      X2              =   6480
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   240
      X2              =   6480
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "Zaman Ayarý (sn)    :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   240
      X2              =   6480
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "Mesaj                          :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   3300
      Width           =   1815
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   240
      X2              =   6480
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Kurbanýn ýd Adresi   :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Facebook Þifreniz    :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Facebook Adresiniz :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   7500
      Left            =   0
      Picture         =   "Form1.frx":F172
      Top             =   0
      Width           =   7500
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 1 Then
Text2.PasswordChar = ""
Check1.Caption = "Gizle"
Else
Text2.PasswordChar = "*"
Check1.Caption = "Göster"
End If
End Sub



Private Sub Form_Load()
WebBrowser1.Navigate "https://m.facebook.com/index.php"
End Sub

Private Sub Form_Resize()
'Form1.Height = 6300
'Form1.Width = 6840
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub PushButton1_Click()
On Error Resume Next
WebBrowser1.Document.All.Item("email").Value = Text1.Text
WebBrowser1.Document.All.Item("pass").Value = Text2.Text
WebBrowser1.Document.Forms(0).Submit
Text3.Enabled = True
Text4.Enabled = True
FlatEdit1.Enabled = True
Text5.Enabled = True
PushButton3.Enabled = True
End Sub

Private Sub PushButton2_Click()
Form2.Show
End Sub

Private Sub PushButton3_Click()
Timer1.Interval = Text5.Text & "000"
PushButton5.Enabled = True
PushButton3.Enabled = False
End Sub

Private Sub PushButton5_Click()
Timer1.Enabled = True
PushButton5.Enabled = False
PushButton6.Enabled = True
WebBrowser2.Navigate "https://www.facebook.com/messages/" & Text3.Text

End Sub

Private Sub PushButton6_Click()
Timer1.Enabled = False
PushButton6.Enabled = False
PushButton5.Enabled = True
PushButton3.Enabled = True
End Sub

Private Sub PushButton7_Click()
MsgBox "Facebook Mesaj Bomber _DangerOusMaN_ Tarafýndan Kodlanmýstýr....                                                          Ýletiþim = DangerOusMaN32@windowslive.com Tüm Haklarým Saklýdýr © 2011 Saldýrýda Oluþabilecek Sorunlarda Sorumluluk Kullanýcýya Aittir.!", vbInformation, "Hakkýnda ;"
End Sub

Private Sub PushButton8_Click()
cýkýs = MsgBox("Programdan Çýkmak mý Ýstiyorsunuz_?", vbInformation + vbYesNo, "Çýkýþ;")
If cýkýs = vbYes Then
Timer1.Enabled = False
End
End If
End Sub


Private Sub Timer1_Timer()
On Error Resume Next
WebBrowser2.Document.All.Item("subject").Value = FlatEdit1.Text
WebBrowser2.Document.Forms(0).Item("body").Value = Text4.Text
WebBrowser2.Document.Forms(0).Submit
WebBrowser2.Navigate "http://m.facebook.com/inbox/?compose&ids=" & Text3.Text & "&refid=17"

End Sub

Private Sub WebBrowser2_DownloadComplete()
Label8.Caption = Label8.Caption + 1
End Sub
