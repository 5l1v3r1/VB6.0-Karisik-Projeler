VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Yetki Düzenleme"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6720
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   6720
   Begin VB.CommandButton Command8 
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
      Left            =   4440
      TabIndex        =   31
      Top             =   5880
      Width           =   2175
   End
   Begin VB.CommandButton Command6 
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
      TabIndex        =   30
      Top             =   5880
      Width           =   2175
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   480
      Top             =   6720
      Width           =   1935
      _ExtentX        =   3413
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Database.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Database.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Personel"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form3.frx":0000
      Height          =   2895
      Left            =   240
      TabIndex        =   17
      Top             =   360
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   5106
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   15
      RowDividerStyle =   6
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
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
      Caption         =   "Personel"
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "ID"
         Caption         =   "ID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1055
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Kullanici_Adi"
         Caption         =   "Kullanici Adi"
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
      BeginProperty Column02 
         DataField       =   "Parola"
         Caption         =   "Parola"
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
      BeginProperty Column03 
         DataField       =   "Yetki"
         Caption         =   "Yetki"
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
            Alignment       =   2
            ColumnWidth     =   494,929
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1500,095
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1500,095
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1500,095
         EndProperty
      EndProperty
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
      Left            =   4080
      TabIndex        =   18
      Top             =   4850
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5280
      Top             =   6600
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
      Left            =   4080
      TabIndex        =   13
      Top             =   4470
      Width           =   2415
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
      Left            =   4080
      TabIndex        =   11
      Top             =   4080
      Width           =   1215
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
      Left            =   5280
      TabIndex        =   10
      Top             =   4080
      Width           =   1210
   End
   Begin VB.Frame Frame2 
      Caption         =   "-Islemler"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   1
      Top             =   3480
      Width           =   6495
      Begin VB.CommandButton Command1 
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
         Left            =   3960
         TabIndex        =   29
         Top             =   220
         Width           =   2415
      End
      Begin VB.Frame Frame4 
         Caption         =   "-Yeni Kayit"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Visible         =   0   'False
         Width           =   3375
         Begin VB.TextBox Text9 
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
            Left            =   1440
            TabIndex        =   28
            Top             =   1110
            Width           =   1575
         End
         Begin VB.TextBox Text8 
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
            Left            =   1440
            TabIndex        =   27
            Top             =   805
            Width           =   1815
         End
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
            Left            =   1440
            TabIndex        =   26
            Top             =   510
            Width           =   1815
         End
         Begin VB.TextBox Text6 
            Alignment       =   2  'Center
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
            Left            =   1440
            TabIndex        =   25
            Top             =   210
            Width           =   615
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "ID :"
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
            Index           =   1
            Left            =   1080
            TabIndex        =   24
            Top             =   240
            Width           =   345
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Kullanici Adi :"
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
            Index           =   1
            Left            =   240
            TabIndex        =   23
            Top             =   550
            Width           =   1170
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Parola :"
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
            Index           =   1
            Left            =   720
            TabIndex        =   22
            Top             =   830
            Width           =   675
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Yetki :"
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
            Index           =   1
            Left            =   840
            TabIndex        =   21
            Top             =   1120
            Width           =   555
         End
      End
      Begin VB.CommandButton Command4 
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
         Left            =   5160
         TabIndex        =   19
         Top             =   1360
         Width           =   1210
      End
      Begin VB.Frame Frame3 
         Height          =   540
         Left            =   120
         TabIndex        =   14
         Top             =   1600
         Width           =   3375
         Begin VB.TextBox Text5 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            DataField       =   "ID"
            DataSource      =   "Adodc2"
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
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   180
            Width           =   1335
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Toplam Kullanici :"
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
            TabIndex        =   15
            Top             =   180
            Width           =   1725
         End
      End
      Begin VB.CommandButton Command2 
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
         Left            =   3960
         TabIndex        =   12
         Top             =   1740
         Width           =   2415
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         DataField       =   "Yetki"
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1280
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         DataField       =   "Parola"
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   950
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         DataField       =   "Kullanici_Adi"
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   620
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "ID"
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   280
         Width           =   735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Yetki :"
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
         Index           =   0
         Left            =   840
         TabIndex        =   5
         Top             =   1280
         Width           =   615
      End
      Begin VB.Label Label3 
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
         Index           =   0
         Left            =   720
         TabIndex        =   4
         Top             =   950
         Width           =   750
      End
      Begin VB.Label Label2 
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
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   620
         Width           =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ID :"
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
         Index           =   0
         Left            =   1080
         TabIndex        =   2
         Top             =   280
         Width           =   345
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "-Kayitli Kullanicilar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
Adodc1.Recordset.MoveFirst
End Sub
Private Sub Command2_Click()
On Error Resume Next
silme = MsgBox("[" & Text2.Text & "] isimli kullanici silinsin mi?", vbYesNo + vbInformation, "Silme Ýþlemi")
If silme = vbYes Then
Adodc1.Recordset.Delete
Adodc1.Recordset.MovePrevious
End If
End Sub
Private Sub Command3_Click()
On Error Resume Next
Frame4.Visible = True
Frame4.Height = 1455
Frame4.Left = 120
Frame4.Top = 240
Frame4.Width = 3375
Text6.Text = Text5.Text + 1
Command4.Enabled = True
Command3.Enabled = False
End Sub
Private Sub Command4_Click()
On Error Resume Next
kaydet = MsgBox("[" & Text7.Text & "] isimli yeni kullanici kaydedilsin mi?", vbYesNo + vbInformation, "Kaydedilsin mi?")
If kaydet = vbYes Then
Adodc1.Recordset.AddNew
Adodc1.Recordset!ID = Text6.Text
Adodc1.Recordset!Kullanici_Adi = Text7.Text
Adodc1.Recordset!Parola = Text8.Text
Adodc1.Recordset!Yetki = Text9.Text
DataGrid1.Refresh
Frame4.Visible = False
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Command3.Enabled = True
Command4.Enabled = False
Else
Adodc1.Recordset.CancelUpdate
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
DataGrid1.Refresh
Frame4.Visible = False
Command3.Enabled = True
Command4.Enabled = False
End If
End Sub
Private Sub Command5_Click()
On Error Resume Next
Adodc1.Recordset.MoveLast
End Sub

Private Sub Command6_Click()
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
   ExcelNesne.Application.Cells(1, 1).Font.Color = vbBlack
ExcelNesne.Application.Cells(1, 1).ColumnWidth = 5
 ExcelNesne.Application.Cells(1, 1).Value = "ID"
 
 ExcelNesne.Application.Cells(1, 2).Font.Size = 11
ExcelNesne.Application.Cells(1, 2).Font.Bold = True
ExcelNesne.Application.Cells(1, 2).Font.Underline = True
   ExcelNesne.Application.Cells(1, 2).Font.Color = vbBlack
ExcelNesne.Application.Cells(1, 2).ColumnWidth = 15
 ExcelNesne.Application.Cells(1, 2).Value = "Kullanici Adi"
 
 ExcelNesne.Application.Cells(1, 3).Font.Size = 11
ExcelNesne.Application.Cells(1, 3).Font.Bold = True
ExcelNesne.Application.Cells(1, 3).Font.Underline = True
   ExcelNesne.Application.Cells(1, 3).Font.Color = vbBlack
ExcelNesne.Application.Cells(1, 3).ColumnWidth = 15
 ExcelNesne.Application.Cells(1, 3).Value = "Parola"
 
 ExcelNesne.Application.Cells(1, 4).Font.Size = 11
ExcelNesne.Application.Cells(1, 4).Font.Bold = True
ExcelNesne.Application.Cells(1, 4).Font.Underline = True
   ExcelNesne.Application.Cells(1, 4).Font.Color = vbBlack
ExcelNesne.Application.Cells(1, 4).ColumnWidth = 15
 ExcelNesne.Application.Cells(1, 4).Value = "Yetki"

i = 1
Adodc1.Recordset.MoveFirst
Do While Not Adodc1.Recordset.EOF = True
i = i + 1

ExcelNesne.Application.Cells(i, 1).Value = Adodc1.Recordset.Fields("ID")
ExcelNesne.Application.Cells(i, 2).Value = Adodc1.Recordset.Fields("Kullanici_Adi")
ExcelNesne.Application.Cells(i, 3).Value = Adodc1.Recordset.Fields("Parola")
ExcelNesne.Application.Cells(i, 4).Value = Adodc1.Recordset.Fields("Yetki")

Adodc1.Recordset.MoveNext
Loop
MsgBox "Microsoft Excel'e Aktarildi Bekleniyor...", vbInformation, "Bildiri;"
Adodc1.Recordset.MoveLast
End Sub

Private Sub Command7_Click()
If Adodc1.Recordset.EOF = True Then
    Adodc1.Recordset.MoveFirst
Else
    Adodc1.Recordset.MoveNext
End If
End Sub
Private Sub Command8_Click()
Unload Me
Form2.Show
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
Private Sub Timer1_Timer()
On Error Resume Next
Text5.Text = Adodc1.Recordset.RecordCount
End Sub
Public Function Text_Kapat()
Text1.Locked = False
Text2.Locked = False
Text3.Locked = False
Text4.Locked = False
End Function
Public Sub Text_Ac()
Form3.Text1.Locked = True
Form3.Text1.Locked = True
Form3.Text2.Locked = True
Form3.Text3.Locked = True
Form3.Text4.Locked = True
End Sub
Public Sub Text_Temizle()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
End Sub
