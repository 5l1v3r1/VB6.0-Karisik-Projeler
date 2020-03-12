VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form9 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form9"
   ClientHeight    =   2100
   ClientLeft      =   6780
   ClientTop       =   4395
   ClientWidth     =   3225
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   3225
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1680
      TabIndex        =   7
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Kapat"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Kaydet"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1680
      TabIndex        =   5
      Top             =   1320
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3720
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Seç"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Hedef Dosya Yolu:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1770
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Kaynak Dosya Yolu:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1920
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error Resume Next
Command1.Caption = "Kaynak Seç"
Command2.Caption = "Hedef Seç"
Command3.Caption = "Dosya Kopyala"
Command4.Caption = "Kapat"
CommonDialog1.Filter = "Microsoft Access(.MDB)|*.mdb"
Form9.Caption = "Veritabani Yedekleme | " & Form4.Text1.Text
End Sub
Private Sub Form_Unload(Cancel As Integer)
Unload Me
Form2.Show
End Sub
Private Sub Command1_Click()
CommonDialog1.DialogTitle = "Lutfen Kopyalanacak Dosyayi Seciniz..."
CommonDialog1.ShowOpen
Text1.Text = CommonDialog1.FileName
End Sub
Private Sub Command2_Click()
CommonDialog1.DialogTitle = "Lutfen Kopyalanacak Yeri Seciniz..."
CommonDialog1.ShowSave
Text2.Text = CommonDialog1.FileName
End Sub
Private Sub Command3_Click()
On Error Resume Next
a = Len(Text2.Text)
b = Len(Text1.Text)
If Mid(Text2.Text, a - 3, 1) <> "." Then
Text2.Text = Text2.Text & Mid(Text1.Text, b - 3, 4)
End If
On Error GoTo hata
FileCopy Text1.Text, Text2.Text
MsgBox "Kopyalama islemi basarili bir sekilde yapilmistir...", vbInformation, "Tebrikler"
Exit Sub
hata:
MsgBox "Kopyalama islemi yapilamiyor... Ayni isimde dosya olabilir...", vbCritical, "Hata"
End Sub
Private Sub Command4_Click()
Unload Me
Form2.Show
End Sub


