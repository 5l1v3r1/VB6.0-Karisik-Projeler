VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8580
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   8580
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   5520
      TabIndex        =   17
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Timer Timer3 
      Interval        =   2000
      Left            =   5160
      Top             =   1560
   End
   Begin VB.Timer Timer2 
      Interval        =   2000
      Left            =   5160
      Top             =   1080
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5040
      Top             =   240
   End
   Begin VB.Frame Frame2 
      Height          =   6375
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   3855
      Begin VB.Frame Frame4 
         Caption         =   "-Genel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Width           =   3615
         Begin VB.TextBox Text4 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1440
            TabIndex        =   18
            Text            =   "50"
            Top             =   320
            Width           =   615
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Oto Intihar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   1200
            TabIndex        =   19
            Top             =   360
            Width           =   255
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "-Oto Pot"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   3615
         Begin VB.ComboBox Combo2 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "Form2.frx":0000
            Left            =   2160
            List            =   "Form2.frx":0013
            TabIndex        =   14
            Text            =   "Seçiniz"
            Top             =   600
            Width           =   1335
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Oto MP"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   650
            Width           =   975
         End
         Begin VB.TextBox Text3 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1440
            MaxLength       =   2
            TabIndex        =   12
            Text            =   "50"
            Top             =   600
            Width           =   615
         End
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "Form2.frx":0030
            Left            =   2160
            List            =   "Form2.frx":0043
            TabIndex        =   9
            Text            =   "Seçiniz"
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1440
            TabIndex        =   8
            Text            =   "50"
            Top             =   240
            Width           =   615
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Oto HP"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   260
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   1200
            TabIndex        =   11
            Top             =   650
            Width           =   255
         End
         Begin VB.Label Label2 
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   1200
            TabIndex        =   10
            Top             =   270
            Width           =   255
         End
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Oyun Acilip Karakter Goruldugunde Buraya Tikla"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   100
      TabIndex        =   4
      Top             =   120
      Width           =   3865
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   6720
      Width           =   3855
      Begin VB.CommandButton Command1 
         Caption         =   "Gönder"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
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
         Left            =   1800
         TabIndex        =   2
         Text            =   "4800"
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Paket Gönderme :"
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
         Left            =   120
         TabIndex        =   1
         Top             =   260
         Width           =   1575
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ptest As Long
Private Declare Function RDWORD Lib "DLL.dll" (ByVal addy As Long) As Long
Private Declare Function CreateP Lib "DLL.dll" (ByVal Direct As String, ByVal DLL As String, ByVal pHANDLE As Long) As Long
Private Declare Function GetInfo Lib "DLL.dll" (ByVal InfoType As Long) As Long
Private Declare Function PartyInfo Lib "DLL.dll" () As PARTY_INFORMATION
Private Declare Function InvInfo Lib "DLL.dll" () As INV_INFORMATION
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Const NEAR As Long = 1
Private Const MEID As Long = 2
Private Const TID As Long = 3
Private Const NT As Long = 4
Private Const MAXMP As Long = 5
Private Const MP As Long = 6
Private Const MaxHP As Long = 7
Private Const HP As Long = 8
Private Const CLASS As Long = 9
Private Const LVL As Long = 10
Private Const GOLD As Long = 11
Private Const EXP As Long = 12
Private Const MAXEXP As Long = 13
Private Const ZONE As Long = 14
Private Const X As Long = 15
Private Const Y As Long = 16
Private Const Z As Long = 17


Private Sub Check3_Click()
If Check3.Value = 1 And Cinfo.HP < ((Cinfo.MaxHP * Text4.Text) / 100) Then
PacketSend "290103"
PacketSend "1200"
End If
End Sub

Private Sub Command1_Click()
Form2.Caption = KarakterID
PacketSend Text1.Text
End Sub

Private Sub Command2_Click()
Form2.Caption = Mid(AlignDWORD(Cinfo.MEID), 1, 4)
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Timer2_Timer()
Cinfo = CharInfo()
If Check1.Value = 1 Then
If Cinfo.HP < ((Cinfo.MaxHP * Text2.Text) / 100) Then
Select Case Combo1.Text
Case "720"
PotBas ("1E")
Case "360"
PotBas ("1D")
Case "180"
PotBas ("1C")
Case "90"
PotBas ("1B")
Case "50"
PotBas ("1A")
End Select
End If
End If
End Sub

Private Sub Timer3_Timer()
Cinfo = CharInfo()
If Check2.Value = 1 Then
If Cinfo.MP < ((Cinfo.MAXMP * Text3.Text) / 100) Then
Select Case Combo2.Text
Case "1920"
PotBas (24)
Case "960"
PotBas (23)
Case "480"
PotBas (22)
Case "180"
PotBas (21)
Case "90"
PotBas (20)
End Select
End If
End If
End Sub
