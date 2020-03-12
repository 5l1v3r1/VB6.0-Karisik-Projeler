VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form12 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Firin Otomasyonu | Yükleniyor..."
   ClientHeight    =   1230
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   5040
   LinkTopic       =   "Form12"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1230
   ScaleWidth      =   5040
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   240
      Top             =   1680
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   840
      TabIndex        =   3
      Top             =   960
      Width           =   555
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Durum:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   660
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "FIRIN PROGRAMI YUKLENIYOR...."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4875
   End
   Begin VB.Menu yeniden_tara 
      Caption         =   "&Yeniden Tara"
   End
   Begin VB.Menu yedekle 
      Caption         =   "&Veritabani Yedekle"
   End
   Begin VB.Menu kapat 
      Caption         =   "&Programi Kapat"
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
ProgressBar1.Value = ProgressBar1.Min
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub kapat_Click()
Dim cýkýs As String
cýkýs = MsgBox("Programdan Çikis Yapilsin mi?", vbYesNo + vbInformation, "Kapatiliyor...")
If cýkýs = vbYes Then
Unload Me
End If
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
ProgressBar1.Value = ProgressBar1.Value + 10
If ProgressBar1.Value = 10 Then
Label3.Caption = "Lütfen Bekleyiniz..."
End If

If ProgressBar1.Value = 30 Then
Label3.Caption = "Program Aciliyor..."
End If

If ProgressBar1.Value = 50 Then
Label3.Caption = "Database.mdb Araniyor"
FileCopy App.Path & "\Databese.mdb", "c:\Database.mdb"

If Dir("c:\Database.mdb") = "" Then
    MsgBox "Dosya Bulunamiyor", vbExclamation, "Hata;"
    Label3.Caption = "Dosya Bulunamadi Hata... Tekrar Kontrol Ediniz."
    Timer1.Enabled = False
    ProgressBar1.Value = 0
Else
    Label3.Caption = "Dosya Bulunuyor. Bekleyiniz..."
End If
End If

If ProgressBar1.Value = 80 Then
Label3.Caption = "Program Acildi...."
End If

If ProgressBar1.Value >= ProgressBar1.Max Then
Timer1.Enabled = False
Form12.Hide
Form1.Show
End If
End Sub

Private Sub yedekle_Click()
Form12.Hide
Form13.Show
End Sub

Private Sub yeniden_tara_Click()
Timer1.Enabled = True
End Sub
