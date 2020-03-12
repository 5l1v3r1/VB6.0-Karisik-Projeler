VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form6 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form6"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5730
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Microsoft Excel'e Aktar"
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
      TabIndex        =   31
      Top             =   6120
      Width           =   2655
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2400
      Top             =   7680
   End
   Begin VB.TextBox Text8 
      DataField       =   "ID"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   600
      TabIndex        =   28
      Text            =   "Text8"
      Top             =   7440
      Width           =   855
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   480
      Top             =   6720
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1085
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Database.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Database.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Calisan_Personel"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame4 
      Height          =   495
      Left            =   120
      TabIndex        =   25
      Top             =   5560
      Width           =   3015
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "0"
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
         Left            =   1560
         TabIndex        =   27
         Top             =   180
         Width           =   135
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Toplam Çalisan:"
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
         TabIndex        =   26
         Top             =   180
         Width           =   1395
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "-Kayit Arama"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   545
      Left            =   3240
      TabIndex        =   22
      Top             =   5520
      Width           =   2415
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
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
         Left            =   720
         TabIndex        =   24
         Top             =   200
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Isim:"
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
         TabIndex        =   23
         Top             =   220
         Width           =   735
      End
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Kayit Silme"
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
      Left            =   3240
      TabIndex        =   21
      Top             =   5115
      Width           =   2415
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Kaydet"
      Enabled         =   0   'False
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
      Left            =   4440
      TabIndex        =   20
      Top             =   4725
      Width           =   1210
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Ilk Kayit"
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
      Left            =   3240
      TabIndex        =   19
      Top             =   3600
      Width           =   2415
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Ileri >"
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
      Left            =   4440
      TabIndex        =   18
      Top             =   3960
      Width           =   1210
   End
   Begin VB.CommandButton Command9 
      Caption         =   "< Geri"
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
      Left            =   3240
      TabIndex        =   17
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "En Son Kayit"
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
      Left            =   3240
      TabIndex        =   16
      Top             =   4350
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Yeni Kayit"
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
      Left            =   3240
      TabIndex        =   15
      Top             =   4725
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "-Kayit Islemleri"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2100
      Left            =   120
      TabIndex        =   2
      Top             =   3480
      Width           =   3015
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         DataSource      =   "Adodc1"
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
         Left            =   960
         TabIndex        =   14
         Top             =   1740
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         DataSource      =   "Adodc1"
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
         Left            =   960
         TabIndex        =   13
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         DataSource      =   "Adodc1"
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
         Left            =   960
         TabIndex        =   12
         Top             =   1140
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         DataSource      =   "Adodc1"
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
         Left            =   960
         TabIndex        =   11
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         DataSource      =   "Adodc1"
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
         Left            =   960
         TabIndex        =   10
         Top             =   540
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataSource      =   "Adodc1"
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
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Maasi:"
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
         Left            =   360
         TabIndex        =   8
         Top             =   1740
         Width           =   555
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Cep Tel:"
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
         Left            =   160
         TabIndex        =   7
         Top             =   1140
         Width           =   735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Görevi:"
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
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   645
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Soyadi:"
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
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ismi:"
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
         Left            =   440
         TabIndex        =   4
         Top             =   540
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ID:"
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
         Left            =   600
         TabIndex        =   3
         Top             =   240
         Width           =   285
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3015
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   5535
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Form6.frx":0000
         Height          =   2655
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   4683
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
   End
   Begin VB.CommandButton Command2 
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
      Height          =   375
      Left            =   4200
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Label10"
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
      TabIndex        =   30
      Top             =   200
      Width           =   660
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
On Error Resume Next
Dim i As Integer
If Adodc1.Recordset.RecordCount = 0 Then
MsgBox "KAYIT YOK"
Exit Sub
End If
Dim ExcelNesne As Object
Set ExcelNesne = CreateObject("Excel.SHEET")
ExcelNesne.Application.Visible = True

 ExcelNesne.Application.Cells(1, 1).Font.Size = 11
ExcelNesne.Application.Cells(1, 1).Font.Bold = True
ExcelNesne.Application.Cells(1, 1).Font.Underline = True
   ExcelNesne.Application.Cells(1, 8).Font.Color = vbBlack
ExcelNesne.Application.Cells(1, 1).ColumnWidth = 5
 ExcelNesne.Application.Cells(1, 1).Value = "ID"
 
 ExcelNesne.Application.Cells(1, 2).Font.Size = 11
ExcelNesne.Application.Cells(1, 2).Font.Bold = True
ExcelNesne.Application.Cells(1, 2).Font.Underline = True
   ExcelNesne.Application.Cells(1, 2).Font.Color = vbBlack
ExcelNesne.Application.Cells(1, 2).ColumnWidth = 15
 ExcelNesne.Application.Cells(1, 2).Value = "Isim"
 
 ExcelNesne.Application.Cells(1, 3).Font.Size = 11
ExcelNesne.Application.Cells(1, 3).Font.Bold = True
ExcelNesne.Application.Cells(1, 3).Font.Underline = True
   ExcelNesne.Application.Cells(1, 3).Font.Color = vbBlack
ExcelNesne.Application.Cells(1, 3).ColumnWidth = 15
 ExcelNesne.Application.Cells(1, 3).Value = "Soyadý"
 
 ExcelNesne.Application.Cells(1, 4).Font.Size = 11
ExcelNesne.Application.Cells(1, 4).Font.Bold = True
ExcelNesne.Application.Cells(1, 4).Font.Underline = True
   ExcelNesne.Application.Cells(1, 4).Font.Color = vbBlack
ExcelNesne.Application.Cells(1, 4).ColumnWidth = 15
 ExcelNesne.Application.Cells(1, 4).Value = "Cep Telefonu"
 
 ExcelNesne.Application.Cells(1, 5).Font.Size = 11
ExcelNesne.Application.Cells(1, 5).Font.Bold = True
ExcelNesne.Application.Cells(1, 5).Font.Underline = True
   ExcelNesne.Application.Cells(1, 5).Font.Color = vbBlack
ExcelNesne.Application.Cells(1, 5).ColumnWidth = 15
 ExcelNesne.Application.Cells(1, 5).Value = "Görevi"
 
 ExcelNesne.Application.Cells(1, 6).Font.Size = 11
ExcelNesne.Application.Cells(1, 6).Font.Bold = True
ExcelNesne.Application.Cells(1, 6).Font.Underline = True
   ExcelNesne.Application.Cells(1, 6).Font.Color = vbBlack
ExcelNesne.Application.Cells(1, 6).ColumnWidth = 15
 ExcelNesne.Application.Cells(1, 6).Value = "Maasi"

 
i = 1
Adodc1.Recordset.MoveFirst
Do While Not Adodc1.Recordset.EOF = True
i = i + 1

ExcelNesne.Application.Cells(i, 1).Value = Adodc1.Recordset.Fields("ID")
ExcelNesne.Application.Cells(i, 2).Value = Adodc1.Recordset.Fields("Isim")
ExcelNesne.Application.Cells(i, 3).Value = Adodc1.Recordset.Fields("Soyad")
ExcelNesne.Application.Cells(i, 4).Value = Adodc1.Recordset.Fields("Cep_Tel")
ExcelNesne.Application.Cells(i, 5).Value = Adodc1.Recordset.Fields("Görevi")
ExcelNesne.Application.Cells(i, 6).Value = Adodc1.Recordset.Fields("Maasi")

Adodc1.Recordset.MoveNext
Loop
MsgBox "Microsoft Excel'e Aktarildi Bekleniyor...", vbInformation, "Bildiri;"
Adodc1.Recordset.MoveLast
End Sub

Private Sub Command2_Click()
Unload Me
Form2.Show
End Sub
Private Sub Command3_Click()
Text1.Text = Text8.Text + 1
Command6.Enabled = True
Command3.Enabled = False
Frame2.Enabled = True
End Sub
Private Sub Command4_Click()
On Error Resume Next
Adodc1.Recordset.MoveFirst
End Sub
Private Sub Command5_Click()
On Error Resume Next
Adodc1.Recordset.MoveLast
End Sub
Private Sub Command6_Click()
On Error Resume Next
kaydet = MsgBox("[" & Text2.Text & "] isimli yeni kullanici kaydedilsin mi?", vbYesNo + vbInformation, "Kaydedilsin mi?")
If kaydet = vbYes Then
Adodc1.Recordset.AddNew
Adodc1.Recordset!ID = Text1.Text
Adodc1.Recordset!Isim = Text2.Text
Adodc1.Recordset!Soyad = Text3.Text
Adodc1.Recordset!Cep_Tel = Text4.Text
Adodc1.Recordset!Görevi = Text5.Text
Adodc1.Recordset!Maasi = Text6.Text
DataGrid1.Refresh
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Command3.Enabled = True
Command6.Enabled = False
Frame2.Enabled = False
Else
Adodc1.Recordset.CancelUpdate
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
DataGrid1.Refresh
Command3.Enabled = True
Command6.Enabled = False
Frame2.Enabled = False
End If
End Sub
Private Sub Command7_Click()
If Adodc1.Recordset.EOF = True Then
    Adodc1.Recordset.MoveFirst
Else
    Adodc1.Recordset.MoveNext
End If
End Sub
Private Sub Command8_Click()
On Error Resume Next
silme = MsgBox("Seçili kullanici silinsin mi?", vbYesNo + vbInformation, "Silme Ýþlemi")
If silme = vbYes Then
Adodc1.Recordset.Delete
Adodc1.Recordset.MovePrevious
End If
End Sub
Private Sub Command9_Click()
If Adodc1.Recordset.BOF = True Then
    Adodc1.Recordset.MoveLast
Else
    Adodc1.Recordset.MovePrevious
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
Unload Me
Form2.Show
End Sub
Private Sub Form_Load()
Form6.Caption = "Personel/Çalisan Listesi | " & Form4.Text1.Text
Label10.Caption = Form4.Text1.Text & "|Çalisan Listesi"
End Sub
Private Sub Text7_Change()
On Error Resume Next
Dim sql As String
sql = "select * from Calisan_Personel where Isim like '%" & Text7.Text & "%'"
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = sql
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
End Sub
Private Sub Timer1_Timer()
Text8.Text = Adodc1.Recordset.RecordCount
Label9.Caption = Text8.Text
End Sub
