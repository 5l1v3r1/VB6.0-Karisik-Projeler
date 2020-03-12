VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form14 
   BackColor       =   &H80000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Çýktý Tutanak Geçmiþi"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   8085
   LinkTopic       =   "Form14"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Tutanak Sil"
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
      Left            =   6720
      TabIndex        =   11
      Top             =   3890
      Width           =   1215
   End
   Begin VB.Timer Timer4 
      Interval        =   10000
      Left            =   6600
      Top             =   120
   End
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   4680
      Top             =   5520
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   8520
      Top             =   1680
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\hastem_db.mdb"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\hastem_db.mdb"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Musteri"
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   9600
      Top             =   2040
   End
   Begin VB.TextBox Text2 
      DataField       =   "Sayin"
      DataSource      =   "Adodc2"
      Height          =   285
      Left            =   8520
      TabIndex        =   8
      Text            =   "Text2"
      Top             =   960
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "HastemTutanakGecmisleri"
      Height          =   255
      Left            =   8520
      TabIndex        =   7
      Top             =   720
      Width           =   2415
   End
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3240
      Left            =   3600
      TabIndex        =   4
      Top             =   600
      Width           =   4335
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   8520
      Top             =   1320
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\hastem_db.mdb"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\hastem_db.mdb"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Musteri"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   8880
      Top             =   2040
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form14.frx":0000
      Height          =   3465
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   6112
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483648
      ColumnHeaders   =   0   'False
      ForeColor       =   128
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1055
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1055
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   975
      Left            =   120
      TabIndex        =   10
      Top             =   4560
      Width           =   7815
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   5640
      Width           =   3135
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3720
      TabIndex        =   6
      Top             =   4200
      Width           =   525
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      Caption         =   "Kayýtlý Tutanaklar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   3600
      TabIndex        =   5
      Top             =   240
      Width           =   1365
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   3
      Top             =   4200
      Width           =   525
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000000&
      Caption         =   "Ara: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Yol
Dim Excel

Private Sub Command1_Click()
On Error Resume Next
MkDir "c:/HastemTutanakGecmisleri"
End Sub

Private Sub Command2_Click()
On Error Resume Next
sil = MsgBox(File1.FileName & " Ýsimli tutanak silinsin mi?", vbYesNo + vbCritical, "Silme Ýþlemi;")
If sil = vbYes Then
Kill "C:\HastemTutanakGecmisleri\" & DataGrid1.Text & "\" & File1.FileName
DataGrid1.Refresh
File1.Refresh
End If
End Sub

Private Sub File1_DblClick()
On Error Resume Next
Excel = "C:\Program Files\Microsoft Office\OFFICE11\EXCEL.exe "
Yol = "C:\HastemTutanakGecmisleri\" & DataGrid1.Text & "\" & File1.FileName

Label6.Caption = Excel & Yol

Dim ac As Excel.Application
Set ac = New Excel.Application
ac.Workbooks.Open Yol
ac.Application.Visible = True

End Sub

Private Sub Form_Load()
On Error Resume Next
File1.Pattern = "*.xls"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Form14.Hide
Form1.Show
End Sub

Private Sub Text1_Change()
On Error Resume Next
Dim sql As String
sql = "select * from Musteri where Sayin like '%" & Text1.Text & "%'"
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = sql
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Label2.Caption = "Toplam Kayýt: " & Adodc1.Recordset.RecordCount
DataGrid1.Refresh
Label4.Caption = "Toplam Tutanak: " & File1.ListCount
Label5.Caption = DataGrid1.Text
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
Adodc2.Recordset.MoveNext
MkDir "c:/HastemTutanakGecmisleri/" & Text2.Text
End Sub

Private Sub Timer3_Timer()
On Error Resume Next
If Label5.Caption = "Ahmet Melih Anadolu Lisesi" Then
File1.Path = "C:\HastemTutanakGecmisleri\" & Label5.Caption
End If
If Label5.Caption = "Anadolu Öðretmen Lisesi" Then
File1.Path = "C:\HastemTutanakGecmisleri\" & Label5.Caption
End If
If Label5.Caption = "Aþ Evi" Then
File1.Path = "C:\HastemTutanakGecmisleri\" & Label5.Caption
End If
If Label5.Caption = "Aybilge" Then
File1.Path = "C:\HastemTutanakGecmisleri\" & Label5.Caption
End If
If Label5.Caption = "Basmacý Oðlu Otel" Then
File1.Path = "C:\HastemTutanakGecmisleri\" & Label5.Caption
End If
If Label5.Caption = "Billur Kýz Öðrenci Yurdu" Then
File1.Path = "C:\HastemTutanakGecmisleri\" & Label5.Caption
End If
If Label5.Caption = "Çöpçü Restaurant" Then
File1.Path = "C:\HastemTutanakGecmisleri\" & Label5.Caption
End If
If Label5.Caption = "Davraz Yaþam Hastanesi" Then
File1.Path = "C:\HastemTutanakGecmisleri\" & Label5.Caption
End If
If Label5.Caption = "Deveci Ticaret" Then
File1.Path = "C:\HastemTutanakGecmisleri\" & Label5.Caption
End If
If Label5.Caption = "Doðum Evi Ana Okulu" Then
File1.Path = "C:\HastemTutanakGecmisleri\" & Label5.Caption
End If


If Label5.Caption = "Eðirdir Türem" Then
File1.Path = "C:\HastemTutanakGecmisleri\" & Label5.Caption
End If
If Label5.Caption = "Fen Lisesi" Then
File1.Path = "C:\HastemTutanakGecmisleri\" & Label5.Caption
End If
If Label5.Caption = "Gazi Lisesi" Then
File1.Path = "C:\HastemTutanakGecmisleri\" & Label5.Caption
End If
If Label5.Caption = "Gülbirlik - Rosense" Then
File1.Path = "C:\HastemTutanakGecmisleri\" & Label5.Caption
End If
If Label5.Caption = "Gülköy Et" Then
File1.Path = "C:\HastemTutanakGecmisleri\" & Label5.Caption
End If
If Label5.Caption = "Güzel Sanatlar Lisesi" Then
File1.Path = "C:\HastemTutanakGecmisleri\" & Label5.Caption
End If
If Label5.Caption = "Ýkbal Restaurant" Then
File1.Path = "C:\HastemTutanakGecmisleri\" & Label5.Caption
End If
If Label5.Caption = "Isparta Tabildot" Then
File1.Path = "C:\HastemTutanakGecmisleri\" & Label5.Caption
End If
If Label5.Caption = "Isparta Türem" Then
File1.Path = "C:\HastemTutanakGecmisleri\" & Label5.Caption
End If
If Label5.Caption = "Kaçýkoç Lisesi Okulu" Then
File1.Path = "C:\HastemTutanakGecmisleri\" & Label5.Caption
End If
If Label5.Caption = "Kaçýkoç Lisesi Pansiyonu" Then
File1.Path = "C:\HastemTutanakGecmisleri\" & Label5.Caption
End If
If Label5.Caption = "Kývýlcým Medikal" Then
File1.Path = "C:\HastemTutanakGecmisleri\" & Label5.Caption
End If

If Label5.Caption = "MEHMET KIRMIZIBAYRAK" Then
File1.Path = "C:\HastemTutanakGecmisleri\" & Label5.Caption
End If
If Label5.Caption = "Mekke Eðitim Vakfý" Then
File1.Path = "C:\HastemTutanakGecmisleri\" & Label5.Caption
End If
If Label5.Caption = "Orman Fakültesi" Then
File1.Path = "C:\HastemTutanakGecmisleri\" & Label5.Caption
End If
If Label5.Caption = "Otogar Lokantasý" Then
File1.Path = "C:\HastemTutanakGecmisleri\" & Label5.Caption
End If
If Label5.Caption = "Sari Kaynak" Then
File1.Path = "C:\HastemTutanakGecmisleri\" & Label5.Caption
End If

If Label5.Caption = "SDÜ Ana Okulu" Then
File1.Path = "C:\HastemTutanakGecmisleri\" & Label5.Caption
End If
If Label5.Caption = "SDÜ Týp Fakültesi" Then
File1.Path = "C:\HastemTutanakGecmisleri\" & Label5.Caption
End If
If Label5.Caption = "Senirkent EML" Then
File1.Path = "C:\HastemTutanakGecmisleri\" & Label5.Caption
End If

If Label5.Caption = "Þenol Kimya" Then
File1.Path = "C:\HastemTutanakGecmisleri\" & Label5.Caption
End If
If Label5.Caption = "Teras Park - Ýgsiad" Then
File1.Path = "C:\HastemTutanakGecmisleri\" & Label5.Caption
End If

If Label5.Caption = "Tutaþ Gýda" Then
File1.Path = "C:\HastemTutanakGecmisleri\" & Label5.Caption
End If
If Label5.Caption = "Uluborlu Lisesi" Then
File1.Path = "C:\HastemTutanakGecmisleri\" & Label5.Caption
End If

If Label5.Caption = "Yalvaç Turizim Otelcilik" Then
File1.Path = "C:\HastemTutanakGecmisleri\" & Label5.Caption
End If
If Label5.Caption = "Yalvaç Türem Pansiyonu" Then
File1.Path = "C:\HastemTutanakGecmisleri\" & Label5.Caption
End If

If Label5.Caption = "Ziraat Fakültesi" Then
File1.Path = "C:\HastemTutanakGecmisleri\" & Label5.Caption
End If


End Sub

Private Sub Timer4_Timer()
DataGrid1.Refresh
File1.Refresh
End Sub
