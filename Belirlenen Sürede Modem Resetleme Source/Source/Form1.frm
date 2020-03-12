VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Modem Resetleme // _DangerOusMaN_"
   ClientHeight    =   2955
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   ScaleHeight     =   2955
   ScaleWidth      =   6180
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Çýkýþ"
      Height          =   495
      Left            =   4560
      TabIndex        =   9
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Modemi Resetle"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   600
      Width           =   1455
   End
   Begin VB.Timer Timer2 
      Left            =   840
      Top             =   3240
   End
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   240
      Top             =   3240
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Deðiþtir"
      Height          =   255
      Left            =   4560
      TabIndex        =   7
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   5160
      MaxLength       =   5
      TabIndex        =   6
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "IP Bilgisini Göster"
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
      Left            =   4560
      TabIndex        =   5
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Gönder"
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
      Left            =   3240
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   2295
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   480
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   720
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Cyber-Warrior.org // _DangerOusMaN_"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   10
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Line Line5 
      X1              =   4440
      X2              =   6120
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line4 
      X1              =   6120
      X2              =   6120
      Y1              =   120
      Y2              =   1800
   End
   Begin VB.Line Line3 
      X1              =   4440
      X2              =   6120
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   4440
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line1 
      X1              =   4440
      X2              =   4440
      Y1              =   120
      Y2              =   2880
   End
   Begin VB.Label Label2 
      Caption         =   "Zaman :"
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
      Left            =   4560
      TabIndex        =   1
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Komut :"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 1 Then
Timer2.Enabled = True
Check1.Caption = "Ýptal Et"
Else
Check1.Caption = "Modemi Resetle"
Timer2.Enabled = False
End If
End Sub

Private Sub Command1_Click()
Text2.Text = ""
Komut (Text1.Text)
End Sub

Private Sub Command2_Click()
Komut ("ipconfig")
End Sub

Private Sub Command3_Click()
cýkýs = MsgBox("Programdan Cýkmak mý Ýstiyorsunuz_?", vbYesNo + 48, "Çýkýþ ;")
If cýkýs = vbYes Then
End
End If
End Sub

Private Sub Command4_Click()
Timer2.Ýnterval = Text3.Text
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Timer1_Timer() 'Timer1
    Dim lngBytesread As Long
    Dim strBuff As String * 2048
    If ReadFile(hReadPipe, strBuff, 2048, lngBytesread, 0&) <> 0 Then
   Text2.Text = Text2.Text & Left(strBuff, lngBytesread)
    Else
    CloseHandle (proc.hProcess)
    CloseHandle (proc.hThread)
    CloseHandle (hReadPipe)
    Timer1.Enabled = False
End If
End Sub
Private Sub Timer2_Timer()
Komut ("ipconfig /flushdns") 'DNS Önbelleðini Siler
Komut ("ipconfig /release") 'Ip Adresimizi Serbest Býrakýr
Komut ("ipconfig /renew") 'Ip Adresini Yeniler
End Sub
