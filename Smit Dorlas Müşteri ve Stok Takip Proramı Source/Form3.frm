VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "ActiveX.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form3 
   BackColor       =   &H00404040&
   Caption         =   "Müþteri Kaydý Düzenleme | Smit & Dorlas"
   ClientHeight    =   6960
   ClientLeft      =   4260
   ClientTop       =   1605
   ClientWidth     =   7245
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6960
   ScaleWidth      =   7245
   Begin XtremeSuiteControls.PushButton PushButton3 
      Height          =   495
      Left            =   240
      TabIndex        =   28
      Top             =   6360
      Width           =   1695
      _Version        =   851968
      _ExtentX        =   2990
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Yeni Kayýt"
      BackColor       =   33023
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton PushButton4 
      Height          =   495
      Left            =   2040
      TabIndex        =   16
      Top             =   6360
      Width           =   1695
      _Version        =   851968
      _ExtentX        =   2990
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Kayýt Sil"
      BackColor       =   33023
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      DataField       =   "Aciklama"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   3600
      Width           =   2175
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      DataField       =   "Musteri_Kayit_Tarihi"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   3240
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      DataField       =   "Musteri_Adres"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Top             =   3960
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      DataField       =   "Musteri_Tel"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   3600
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      DataField       =   "Musteri_Adi"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Top             =   3240
      Width           =   2175
   End
   Begin XtremeSuiteControls.PushButton PushButton2 
      Height          =   495
      Left            =   5640
      TabIndex        =   0
      Top             =   6360
      Width           =   1455
      _Version        =   851968
      _ExtentX        =   2566
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Kapat"
      BackColor       =   33023
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   495
      Left            =   3840
      TabIndex        =   1
      Top             =   6360
      Width           =   1695
      _Version        =   851968
      _ExtentX        =   2990
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Güncelleþtir"
      BackColor       =   33023
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox GroupBox3 
      Height          =   1935
      Left            =   240
      TabIndex        =   7
      Top             =   4320
      Width           =   6735
      _Version        =   851968
      _ExtentX        =   11880
      _ExtentY        =   3413
      _StockProps     =   79
      Caption         =   "*Kahve Detay"
      BackColor       =   33023
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Begin VB.TextBox Text11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   35
         Text            =   "Text11"
         Top             =   360
         Width           =   1215
      End
      Begin XtremeSuiteControls.FlatEdit FlatEdit2 
         Height          =   255
         Left            =   6000
         TabIndex        =   33
         Top             =   1080
         Visible         =   0   'False
         Width           =   735
         _Version        =   851968
         _ExtentX        =   1296
         _ExtentY        =   450
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "FlatEdit2"
      End
      Begin VB.TextBox Text10 
         DataField       =   "Musteri_Tuketilen_Bardak"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   6120
         TabIndex        =   31
         Text            =   "Text10"
         Top             =   600
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   6240
         Top             =   1440
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         DataField       =   "Musteri_KahveVerilenTarih"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2400
         TabIndex        =   30
         Text            =   "Text9"
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2400
         TabIndex        =   22
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "Musteri_Verilen_Kahve"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2400
         TabIndex        =   8
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label15 
         BackColor       =   &H000080FF&
         Caption         =   "Günceleme Tarihi:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   34
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label14 
         BackColor       =   &H000080FF&
         Caption         =   "Kahve Verilen Tarih :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   29
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Kalan Kahve (Bardak) : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   430
         TabIndex        =   27
         Top             =   1460
         Width           =   1965
      End
      Begin VB.Label Label13 
         BackColor       =   &H000080FF&
         Caption         =   "Açýklama: Verilen kahveden kalan bardak sayýsý."
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
         Left            =   3600
         TabIndex        =   26
         Top             =   1440
         Width           =   2895
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Tüketilen Kahve (Bardak) :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   100
         TabIndex        =   24
         Top             =   1100
         Width           =   2250
      End
      Begin VB.Label Label12 
         BackColor       =   &H000080FF&
         Caption         =   "Açýklama: Tüketilen bardak miktarýný giriniz."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3600
         TabIndex        =   23
         Top             =   1020
         Width           =   2895
      End
      Begin VB.Label Label11 
         BackColor       =   &H000080FF&
         Caption         =   "Açýklama: 1 adet 1000 gr'dýr."
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
         Left            =   3600
         TabIndex        =   10
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Verilen Kahve (Adet) :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   500
         TabIndex        =   9
         Top             =   720
         Width           =   1845
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   3015
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   6975
      _Version        =   851968
      _ExtentX        =   12303
      _ExtentY        =   5318
      _StockProps     =   79
      Caption         =   "Müþteri Listesi"
      BackColor       =   33023
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   6600
         Top             =   4200
      End
      Begin XtremeSuiteControls.FlatEdit FlatEdit1 
         Height          =   255
         Left            =   1320
         TabIndex        =   18
         Top             =   2640
         Width           =   2055
         _Version        =   851968
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Appearance      =   1
         UseVisualStyle  =   0   'False
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Form3.frx":0000
         Height          =   2295
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   4048
         _Version        =   393216
         AllowUpdate     =   0   'False
         ColumnHeaders   =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "Musteri_Adi"
            Caption         =   "Müþteri Adý"
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
            DataField       =   "Musteri_Tel"
            Caption         =   "Telefon"
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
            DataField       =   "Musteri_Adres"
            Caption         =   "Adres"
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
            DataField       =   "Musteri_Kayit_Tarihi"
            Caption         =   "Kayýt Tarihi"
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
         BeginProperty Column04 
            DataField       =   "Aciklama"
            Caption         =   "Açýklama"
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
               ColumnWidth     =   2009,764
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   3509,858
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   2009,764
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   3014,929
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Müþteri Ara :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   2640
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Toplam Kayýt :"
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
         Left            =   5160
         TabIndex        =   20
         Top             =   2640
         Width           =   1395
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   10560
      Top             =   1800
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1296
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\SmitDorlasKahve_db.mdb"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\SmitDorlasKahve_db.mdb"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Musteriler_db"
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
   Begin VB.Label Label16 
      Caption         =   "Label16"
      DataField       =   "ID"
      DataSource      =   "Adodc1"
      Height          =   135
      Left            =   120
      TabIndex        =   32
      Top             =   3960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "Açýklama : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   195
      Left            =   3900
      TabIndex        =   15
      Top             =   3600
      Width           =   915
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "Kayýt Tarihi :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   195
      Left            =   3720
      TabIndex        =   14
      Top             =   3240
      Width           =   1050
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "Adresi :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   195
      Left            =   645
      TabIndex        =   13
      Top             =   3960
      Width           =   630
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   " Telefon :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   195
      Left            =   525
      TabIndex        =   12
      Top             =   3600
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "Müþteri Adý : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   3240
      Width           =   1095
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VerilenKahve, ToplamBardak, KalanBardak, TüketilenToplam, Tüketilen As String

Private Sub FlatEdit1_Change()
On Error Resume Next
Dim SQL1 As String
SQL1 = "select * from Musteriler_db where Musteri_Adi like '%" & FlatEdit1.Text & "%'"
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = SQL1
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
End Sub

Private Sub Form_Activate()
DataGrid1.Refresh
Adodc1.Refresh
Adodc1.Recordset.MoveLast
Timer2.Enabled = True
Form3.Left = 4140
Form3.Top = 1155
End Sub

Private Sub Form_Load()
On Error Resume Next
Form1.Adodc1.Recordset.CancelUpdate
Form1.Adodc1.Refresh
Form1.DataGrid1.Refresh
Adodc1.Recordset.CancelUpdate
Adodc1.Refresh
DataGrid1.Refresh
Adodc1.Recordset.MoveLast
End Sub

Private Sub Form_Resize()
On Error Resume Next
Form3.Height = 7530
Form3.Width = 7485
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Form1.Adodc1.Recordset.CancelUpdate
Adodc1.Recordset.CancelUpdate
Adodc1.Refresh
DataGrid1.Refresh
Unload Me
Form1.Adodc1.Refresh
Form1.DataGrid1.Refresh
Form1.Show
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
End Sub

Private Sub PushButton1_Click()
On Error Resume Next
kaydet = MsgBox("[" & Text1.Text & "] isimli müþteri güncellensin mi?", vbYesNo + vbInformation, "Güncelleme;?")
If kaydet = vbYes Then
Timer2.Enabled = False
Güncellestir
Güncellestir
Form1.Show
Unload Me
Else
Adodc1.Recordset.CancelUpdate
DataGrid1.Refresh
End If
End Sub
Private Function Güncellestir()
With Adodc1.Recordset
        .Update
        .Fields("Musteri_KalanKahve").Value = Text8.Text
        .Fields("Musteri_Tuketilen_Bardak").Value = FlatEdit2.Text
        .Fields("Musteri_Adi").Value = Text1.Text
        .Fields("Musteri_Tel").Value = Text2.Text
        .Fields("Musteri_Adres").Value = Text3.Text
        .Fields("Musteri_Kayit_Tarihi").Value = Text4.Text
        .Fields("Aciklama").Value = Text5.Text
        .Fields("Musteri_Verilen_Kahve").Value = Text6.Text
        .Fields("Musteri_KahveVerilenTarih").Value = Text9.Text
        .Fields("ID").Value = Label16.Caption
        .Fields("GuncelKahve").Value = Text11.Text
End With
End Function
Private Sub PushButton2_Click()
On Error Resume Next
Adodc1.Refresh
DataGrid1.Refresh
Unload Me
Form1.Show
End Sub

Private Sub PushButton3_Click()
On Error Resume Next
Unload Me
Form1.Hide
Form2.Show
End Sub

Private Sub PushButton4_Click()
On Error Resume Next
silme = MsgBox(Text1.Text & " isimli müþteri silinsin mi?", vbYesNo + vbInformation, "Silme Ýþlemi")
If silme = vbYes Then
Adodc1.Recordset.Delete
Adodc1.Recordset.MoveLast
Adodc1.Recordset.MoveLast
End If
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Label2.Caption = "Toplam Kayýt : " & Adodc1.Recordset.RecordCount
End Sub

Private Sub Timer2_Timer()
Dim ToplamKahve, ToplamTüketilen, KalanKahve As String
ToplamKahve = Val(Text6.Text) * (150)

ToplamTüketilen = Val(Text7.Text) + Val(Text10.Text)

KalanKahve = Val(ToplamKahve) - Val(ToplamTüketilen)

Text8.Text = KalanKahve

FlatEdit2.Text = ToplamTüketilen
Text11.Text = Date
End Sub
