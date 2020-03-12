VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "ActiveX.ocx"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Özel Klasör ve Dosya Seçimi"
   ClientHeight    =   2430
   ClientLeft      =   4560
   ClientTop       =   3015
   ClientWidth     =   5370
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   5370
   Begin XtremeSuiteControls.PushButton Command4 
      Height          =   375
      Left            =   3960
      TabIndex        =   8
      Top             =   2040
      Width           =   1215
      _Version        =   851968
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Ýptal"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton Command3 
      Height          =   375
      Left            =   2640
      TabIndex        =   7
      Top             =   2040
      Width           =   1215
      _Version        =   851968
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Tamam"
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
   Begin XtremeSuiteControls.PushButton Command2 
      Height          =   255
      Left            =   4320
      TabIndex        =   6
      Top             =   1560
      Width           =   975
      _Version        =   851968
      _ExtentX        =   1720
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Gözat"
      ForeColor       =   0
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
   Begin XtremeSuiteControls.FlatEdit Text2 
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   1560
      Width           =   3855
      _Version        =   851968
      _ExtentX        =   6800
      _ExtentY        =   450
      _StockProps     =   77
      ForeColor       =   8421504
      BackColor       =   -2147483643
      Enabled         =   0   'False
   End
   Begin XtremeSuiteControls.PushButton Command1 
      Height          =   255
      Left            =   4320
      TabIndex        =   4
      Top             =   840
      Width           =   975
      _Version        =   851968
      _ExtentX        =   1720
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Gözat"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.RadioButton Radio2 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   855
      _Version        =   851968
      _ExtentX        =   1508
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Dosya"
      ForeColor       =   0
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
   Begin XtremeSuiteControls.RadioButton Radio1 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1575
      _Version        =   851968
      _ExtentX        =   2778
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Sürücü ve Dizin"
      ForeColor       =   0
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
      Value           =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit Text1 
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   840
      Width           =   3855
      _Version        =   851968
      _ExtentX        =   6800
      _ExtentY        =   450
      _StockProps     =   77
      ForeColor       =   8421504
      BackColor       =   -2147483643
      Enabled         =   0   'False
   End
   Begin XtremeSuiteControls.CommonDialog cmd2 
      Left            =   5400
      Top             =   1560
      _Version        =   851968
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   4
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   5280
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label1 
      Caption         =   "Özel :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   162
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin XtremeSuiteControls.CommonDialog cmd1 
      Left            =   5400
      Top             =   840
      _Version        =   851968
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   4
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Sub Form_Initialize()
InitCommonControls
End Sub
Private Sub Command1_Click()
cmd1.ShowBrowseFolder
Text1.Text = cmd1.FileName
End Sub

Private Sub Command2_Click()
cmd2.ShowOpen
Text2.Text = cmd2.FileName
End Sub

Private Sub Command3_Click()
On Error Resume Next
'Sürücü ve Dizin
If Radio1.Enabled = True Then
If Radio1.Value = True Then
Dosya = cmd1.FileName & "\*.*"
Dim Dosyalar As ListItem
Set Dosyalar = Form1.ListView2.ListItems.Add
Dosyalar.SubItems(1) = Dosya
Dosyalar.SubItems(2) = Sürücü
End If
End If

'Dosya
If Radio2.Enabled = True Then
If Radio2.Value = True Then
Dosya = cmd2.FileName
Dim Dosyalar1 As ListItem
Set Dosyalar1 = Form1.ListView2.ListItems.Add
Dosyalar1.SubItems(1) = Dosya
Dosyalar1.SubItems(2) = Sürücü
End If
End If
'Kod Bitiþ
Form2.Hide
Form1.Enabled = True
Form1.Show
MsgBox "Seçtiðiniz Klasör veya Dosyalar Cleaner Seçeneðinin Windows Sekmesindeki 'Özel Ek Dosya ve Klasördeki Seçili Ögeleri Temizleme' Seçeneðine Týklamanýz Yeterli Olacaktýr.", vbInformation, "Bildiri ;"
End Sub

Private Sub Command4_Click()
Form1.Enabled = True
Form2.Hide
Form1.Show
End Sub

Private Sub Form_Load()
Form1.Enabled = False
Menu1
cmd1.DialogTitle = "Klasör veya Dizin :"
cmd2.DialogTitle = "Aç"
cmd2.Filter = "*.*|*.*"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.Enabled = True
Form2.Hide
End Sub

Private Sub Radio1_Click()
Command2.Enabled = False
Command1.Enabled = True
End Sub
Private Sub Radio2_Click()
Command1.Enabled = False
Command2.Enabled = True
End Sub

Private Sub Radio3_Click()
Text3.Enabled = False
Radio4.Enabled = False
End Sub

Private Sub Radio4_Click()
Radio3.Enabled = False
Text3.Enabled = True
End Sub

Private Sub Text1_Change()
Command3.Enabled = True
End Sub

Private Sub Text2_Change()
Command3.Enabled = True
End Sub
