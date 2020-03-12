VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Masaüstü Sað Menü Program Ekleme"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   4425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   4425
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   120
      TabIndex        =   11
      Text            =   "Text4"
      Top             =   3480
      Width           =   4095
   End
   Begin MSComDlg.CommonDialog Cmd2 
      Left            =   840
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog Cmd1 
      Left            =   240
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Hakkýnda"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      TabIndex        =   10
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Çýkýþ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3360
      TabIndex        =   9
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Uygula"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   8
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Seç"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3720
      TabIndex        =   7
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Seç"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3720
      TabIndex        =   4
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2400
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   4320
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label3 
      Caption         =   "Ýcon Yolu Seç :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   4320
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   4320
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label2 
      Caption         =   "Eklemek Ýstediðiniz Programýn Yolu :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   2655
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4320
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label1 
      Caption         =   "Görünmesini Ýstediðiniz Ýsim :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Cmd1.ShowOpen ' Dosya Penceresi Açar
Text2.Text = Cmd1.FileName ' Seçilen Dosyanýn Yolunu Text2 Kaydeder
Text4.Text = "Directory\shell\" & Text1.Text
End Sub
Private Sub Command2_Click()
Cmd2.ShowOpen ' Dosya Penceresi Açar
Text3.Text = Cmd2.FileName ' Seçilen Dosyanýn Yolunu Text3 Kaydeder
End Sub
Private Sub Command3_Click()
Dim KayitDefteri As Object
Dim reg As Object
Set KayitDefteri = CreateObject("wscript.shell")
RegKaydiYaz HKEY_CLASSES_ROOT, Text4.Text & "\command", "", ""
KayitDefteri.regwrite "HKEY_CLASSES_ROOT\" & Text4.Text & "\command\", Text2.Text
KayitDefteri.regwrite "HKEY_CLASSES_ROOT\" & Text4.Text & "\command\" & "icon", Text3.Text
End Sub
Private Sub Command4_Click()
cýkýs = MsgBox("Programdan Çýkmak mý Ýstiyorsunuz ?", vbQuestion + vbYesNo, " Çýkýþ ;")
If cýkýs = vbYes Then
End
End If
End Sub
Private Sub Command5_Click()
MsgBox "_DangerOusMaN_ Tarafýndan  Kodlanmýstýr....                                                                      Ýletiþim = DangerOusMaN32@windowslive.com     Bütün Haklarým Saklýdýr  © 2011                                                         ''YAPILAN TÜM UYGULAMALARDA SORUMLULUK KÝÞÝYE AÝTTÝR.!''", 48, "Hakkýnda ;"
End Sub
Private Sub Form_Load()
Cmd1.DialogTitle = "Dosya Yolu Seçiniz : "
Cmd2.DialogTitle = "Ýcon Yolu Seçiniz : "
Cmd1.Filter = "Exe Dosyasý (.exe)|*.exe"
Cmd2.Filter = "Ýcon Simgesi (.ico)|*.ico"
End Sub
Private Sub Te()
Command3.Enabled = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
End
End Sub
Private Sub Text2_Change()
Command3.Enabled = True
End Sub
