VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "ActiveX.ocx"
Begin VB.Form Form3 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Pc Ýyileþtirmede Regedit Ayarlarý ve Girilen Deðerler Açýk Kaynak"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   8880
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   8880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.FlatEdit FlatEdit1 
      Height          =   5535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8895
      _Version        =   851968
      _ExtentX        =   15690
      _ExtentY        =   9763
      _StockProps     =   77
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   $"Form3.frx":0000
      MultiLine       =   -1  'True
      ScrollBars      =   3
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Form1.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.PushButton4.Enabled = True
Form1.Enabled = True
MsgBox "Regedit Ayarlarý Uzmanlar içindir.! Girilen Deðerlerde Bilgisayarý Olumsuz Yönde Etkilememiþtir ve Bizzat Denenmiþtir.! Not: Herhangi Bir Durumda Oluþabilecek Zararlardan Kullanýcý Sorumludur. Kodlayan Hiçbir Þekilde Sorumlu Tutulamaz...!", vbInformation, "UYARI.!"
End Sub
