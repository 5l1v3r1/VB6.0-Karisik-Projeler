VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "ActiveX.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   ScaleHeight     =   8085
   ScaleWidth      =   10560
   StartUpPosition =   3  'Windows Default
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   840
      Top             =   5880
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AccessType      =   1
      URL             =   "http://"
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   4920
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   615
      Left            =   1320
      TabIndex        =   15
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "-Modem Tarayýcý ve Marka/Model Bulucu"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8055
      Left            =   4080
      TabIndex        =   10
      Top             =   0
      Width           =   6375
      Begin VB.TextBox Text6 
         Height          =   1695
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Text            =   "Form1.frx":0000
         Top             =   6240
         Width           =   6135
      End
      Begin XtremeSuiteControls.WebBrowser WebBrowser1 
         Height          =   5415
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   6135
         _Version        =   851968
         _ExtentX        =   10821
         _ExtentY        =   9551
         _StockProps     =   173
         BackColor       =   -2147483643
         ScriptErrorsSuppressed=   -1  'True
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
         Left            =   480
         TabIndex        =   12
         Text            =   "127.0.0.1"
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "IP :"
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
         TabIndex        =   11
         Top             =   360
         Width           =   315
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   600
      Top             =   4920
   End
   Begin VB.Frame Frame1 
      Caption         =   "-IP Oluþturucu"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2400
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   3735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Oluþtur"
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
         Left            =   2880
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         MaxLength       =   3
         TabIndex        =   4
         Text            =   "255"
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   3
         Text            =   "7"
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         MaxLength       =   3
         TabIndex        =   2
         Text            =   "183"
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         MaxLength       =   3
         TabIndex        =   1
         Text            =   "78"
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "."
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
         Index           =   2
         Left            =   2160
         TabIndex        =   7
         Top             =   480
         Width           =   60
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "."
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
         Index           =   1
         Left            =   1440
         TabIndex        =   6
         Top             =   480
         Width           =   60
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "."
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
         Index           =   0
         Left            =   720
         TabIndex        =   5
         Top             =   480
         Width           =   60
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
List1.Clear
Bytes = "0"
For X = 1 To Text4.Text
Bytes = Bytes + 1
List1.AddItem Text1.Text & "." & Text2.Text & "." & Text3.Text & "." & Bytes & vbNewLine
pause 0.000001 'without the pause it will freeze if you do over 1000
Next X
End Sub
Sub pause(interval)
    Current = Timer
    Do While Timer - Current < Val(interval)
        DoEvents
    Loop
End Sub

Private Sub Command2_Click()
Timer2.Enabled = True
End Sub

Private Sub Text1_Change()
If Text1.Text > 255 Then
MsgBox "255'den fazla olamaz", vbCritical, "Hata;"
End If
End Sub

Private Sub Text2_Change()
If Text2.Text > 255 Then
MsgBox "255'den fazla olamaz", vbCritical, "Hata;"
End If
End Sub

Private Sub Text3_Change()
If Text3.Text > 255 Then
MsgBox "255'den fazla olamaz", vbCritical, "Hata;"
End If
End Sub

Private Sub Text4_Change()
If Text4.Text > 255 Then
MsgBox "255'den fazla olamaz", vbCritical, "Hata;"
End If
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Text6.Text = WebBrowser1.Document.bOdy.InnerHTML
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
With List1
'*Ýndex sayýsýný 1 arttýrýr
.ListIndex = .ListIndex + 1
'*Seçili elemanýn deðeri
Text5.Text = .List(.ListIndex)
WebBrowser1.Navigate "Http:\\" & Text5.Text
'Text6.Text = Inet1.OpenURL(Text5.Text)
'* Listenin son elemanýna gelince bilgi mesajý verir
If .ListIndex = .ListCount - 1 Then
MsgBox "Sýra Ýle Okuma Ýþlemi Yerine Getirildi..!!", vbInformation, "ÝÞLEM YAPILDI"
.ListIndex = -1
'* Timer1 Durdurur
Timer2.Enabled = False
End If
End With
End Sub

