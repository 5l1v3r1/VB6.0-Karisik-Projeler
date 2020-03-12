VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Program Ismi"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6630
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   6630
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   6000
      Top             =   120
   End
   Begin ComctlLib.StatusBar Status1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   2505
      Width           =   6630
      _ExtentX        =   11695
      _ExtentY        =   661
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Personel/Çalisan Listesi"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3360
      TabIndex        =   3
      Top             =   1560
      Width           =   3135
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Firma/Dükkan/Toptanci Hesaplari"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Gider Tablosu"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3360
      TabIndex        =   1
      Top             =   600
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Gelir Tablosu"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   3135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1050
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
On Error Resume Next
Form3.Hide
Form4.Hide
Form5.Hide
Form6.Hide
Form8.Hide
Form2.Hide
Form7.Hide
Form9.Hide
Form10.Show
End Sub

Private Sub Command2_Click()
On Error Resume Next
Form3.Hide
Form4.Hide
Form5.Hide
Form6.Hide
Form8.Hide
Form2.Hide
Form7.Hide
Form9.Hide
Form10.Hide
Form11.Show
End Sub

Private Sub Command3_Click()
On Error Resume Next
Form3.Hide
Form4.Hide
Form5.Hide
Form6.Hide
Form9.Hide
Form8.Hide
Form2.Hide
Form7.Show
End Sub
Private Sub Command4_Click()
On Error Resume Next
Form3.Hide
Form2.Hide
Form4.Hide
Form5.Hide
Form9.Hide
Form7.Hide
Form8.Hide
Form6.Show
End Sub
Private Sub Form_Load()
Form2.Caption = "Genel Ana Panel | " & Form4.Text1.Text
Label1.Caption = Form4.Text1.Text
End Sub
Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub
Private Sub Timer1_Timer()
On Error Resume Next
Status1.SimpleText = "Saat: " & TimeValue(Now) & " | Tarih: " & Format(Date, "dd.mmmm.yyyy") & " | Günlerden: " & WeekdayName(Weekday(Now) - 1)
End Sub
