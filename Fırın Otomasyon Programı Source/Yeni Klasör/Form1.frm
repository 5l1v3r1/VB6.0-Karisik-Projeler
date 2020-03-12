VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Giris Paneli"
   ClientHeight    =   2970
   ClientLeft      =   5505
   ClientTop       =   4320
   ClientWidth     =   3720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   3720
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   120
      TabIndex        =   13
      Text            =   "Text4"
      Top             =   4560
      Width           =   3255
   End
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   2400
      Width           =   3495
      Begin VB.Label Label6 
         Caption         =   "Durum: Bekliyor..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   165
         Width           =   3135
      End
      Begin VB.Label Label5 
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
         Left            =   120
         TabIndex        =   11
         Top             =   170
         Width           =   855
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Programi Kapat"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Giris Yap"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin VB.TextBox Text3 
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
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   6
         Top             =   1300
         Width           =   1695
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
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   600
         Width           =   1695
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
         Left            =   1680
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "000000"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   270
         Left            =   1680
         TabIndex        =   7
         Top             =   960
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Kullanici Adi :"
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
         Top             =   240
         Width           =   1320
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Parola :"
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
         TabIndex        =   2
         Top             =   600
         Width           =   750
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Güvenlik Kodu :"
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
         Top             =   970
         Width           =   1515
      End
   End
   Begin VB.Menu güvenlikyenile 
      Caption         =   "&Güvenlik Kod Yenile"
   End
   Begin VB.Menu hakkinda 
      Caption         =   "Hakkinda"
   End
   Begin VB.Menu cýkýs 
      Caption         =   "Çikis"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cn As New ADODB.Connection, strCNString As String
Dim rs As New ADODB.Recordset
Dim Txt As String
Private Sub Command1_Click()
On Error GoTo ErrHandler
strCNString = "Data Source=" & "c:\Database.mdb" 'database adýný yazýn.
cn.Provider = "Microsoft Jet 4.0 OLE DB Provider"
cn.ConnectionString = strCNString
cn.Properties("Jet OLEDB:Database Password") = "holocaust32" 'Database þifreli ise buraya þifreyi yazýn
cn.Open
If Label4.Caption = Text3.Text Then
With rs
.Open "Select * from Personel where Kullanici_Adi='" & Text1.Text & "' and Parola='" & Text2.Text & "'", cn, adOpenDynamic, adLockOptimistic
If .EOF Then
MsgBox "Kullanici Adi veya Parola Yanlis!", vbOKOnly + vbCritical, "Uyari"
Label6.Caption = "Kullanici Adi veya Parola Yanlis!"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Güvenlik_Random_Kod
Text1.SetFocus
cn.Close
Else
Txt = "" & " " & UCase$(Text1.Text) & ""
Text4.Text = ""
Text4.Text = "[" & Time & "/" & Date & "] saatinde ve tarihinde [" & Txt & " ] kullanicisi baglandi."
MsgBox Txt & " adli kisi baglaniyor...", vbOKOnly + vbExclamation, "Giris"
Label6.Caption = Txt & " adlý kiþi baðlandý."
cn.Close
Form1.Hide
Form2.Show
End If
End With
Exit Sub
ErrHandler:
MsgBox Err.Description, vbCritical, "Giris"
cn.Close
Else
MsgBox "Guvenlik Kodunda Hata Var!.", vbCritical, "Hata"
Label6.Caption = "Guvenlik Kodunda Hata Var."
Text3.Text = ""
cn.Close
Güvenlik_Random_Kod
End If
End Sub
Private Sub Command2_Click()
Dim cýkýs As String
cýkýs = MsgBox("Programdan Çikis Yapilsin mi?", vbYesNo + vbInformation, "Kapatiliyor...")
If cýkýs = vbYes Then
Unload Me
End If
End Sub
Private Sub Form_Load()
Güvenlik_Random_Kod
Text4.Text = ""
End Sub
Private Sub cýkýs_Click()
Dim cýkýs As String
cýkýs = MsgBox("Programdan Çikis Yapilsin mi?", vbYesNo + vbInformation, "Kapatiliyor...")
If cýkýs = vbYes Then
Unload Me
End If
End Sub
Private Function Güvenlik_Random_Kod()
On Error Resume Next
Randomize
Dim harfler(3) As String
Dim sayýlar(3) As Integer
harfler(0) = Chr(Math.Round(Rnd() * (122 - 97) + 97))
harfler(1) = Chr(Math.Round(Rnd() * (122 - 97) + 97))
harfler(2) = Chr(Math.Round(Rnd() * (122 - 97) + 97))
sayýlar(0) = Math.Round(Rnd() * (9 - 0) + 0)
sayýlar(1) = Math.Round(Rnd() * (9 - 0) + 0)
sayýlar(2) = Math.Round(Rnd() * (9 - 0) + 0)
Form1.Label4.Caption = harfler(0) & harfler(1) & sayýlar(0) & sayýlar(1) & harfler(2) & sayýlar(2)
End Function
Private Sub Form_Resize()
Form1.Height = 3720
Form1.Width = 3810
End Sub
Private Sub güvenlikyenile_Click()
Güvenlik_Random_Kod
Text4.Text = ""
End Sub

Private Sub hakkinda_Click()
MsgBox "Otomasyon Programý 'Osman Yavuz' Tarafindan Kodlanmýþtýr...", vbExclamation, "Hakkinda;"
End Sub
