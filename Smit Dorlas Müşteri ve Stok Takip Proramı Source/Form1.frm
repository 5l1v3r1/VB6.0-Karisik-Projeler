VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "ActiveX.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00404040&
   Caption         =   "Smit & Dorlas Müþteri ve Stok Takip Programý"
   ClientHeight    =   9795
   ClientLeft      =   1815
   ClientTop       =   945
   ClientWidth     =   15045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9795
   ScaleWidth      =   15045
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   10800
      TabIndex        =   38
      Top             =   9120
      Width           =   495
   End
   Begin VB.TextBox Text12 
      DataField       =   "Musteri_Tel"
      DataSource      =   "Adodc2"
      Height          =   285
      Left            =   9720
      TabIndex        =   37
      Text            =   "telefon"
      Top             =   9120
      Width           =   735
   End
   Begin VB.TextBox Text11 
      DataField       =   "ID"
      DataSource      =   "Adodc2"
      Height          =   285
      Left            =   9000
      TabIndex        =   36
      Text            =   "ýd"
      Top             =   9120
      Width           =   615
   End
   Begin VB.Timer Timer2 
      Interval        =   200
      Left            =   13320
      Top             =   6240
   End
   Begin VB.TextBox Text10 
      DataField       =   "Musteri_KalanKahve"
      DataSource      =   "Adodc2"
      Height          =   285
      Left            =   8160
      TabIndex        =   35
      Text            =   "kalan"
      Top             =   9120
      Width           =   735
   End
   Begin VB.TextBox Text9 
      DataField       =   "Musteri_Adi"
      DataSource      =   "Adodc2"
      Height          =   285
      Left            =   7200
      TabIndex        =   34
      Text            =   "isim"
      Top             =   9120
      Width           =   735
   End
   Begin XtremeSuiteControls.GroupBox GroupBox4 
      Height          =   3735
      Left            =   7320
      TabIndex        =   30
      Top             =   4920
      Width           =   5655
      _Version        =   851968
      _ExtentX        =   9975
      _ExtentY        =   6588
      _StockProps     =   79
      Caption         =   "*Kahve Takip"
      ForeColor       =   0
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
      Begin XtremeSuiteControls.PushButton PushButton5 
         Height          =   495
         Left            =   2400
         TabIndex        =   41
         Top             =   480
         Width           =   2895
         _Version        =   851968
         _ExtentX        =   5106
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Kahvesi Azalan Müþterileri Bul"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.GroupBox GroupBox5 
         Height          =   855
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   1935
         _Version        =   851968
         _ExtentX        =   3413
         _ExtentY        =   1508
         _StockProps     =   79
         Caption         =   "Filtre"
         BackColor       =   33023
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.RadioButton RadioButton1 
            Height          =   255
            Left            =   360
            TabIndex        =   33
            Top             =   240
            Width           =   1215
            _Version        =   851968
            _ExtentX        =   2143
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Kalan Kahve"
            BackColor       =   33023
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit FlatEdit2 
            Height          =   255
            Left            =   600
            TabIndex        =   32
            ToolTipText     =   "Girilen deðerden düþük olanlarý listeler."
            Top             =   480
            Width           =   855
            _Version        =   851968
            _ExtentX        =   1508
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
            Enabled         =   0   'False
            Text            =   "150"
            Alignment       =   2
         End
      End
      Begin XtremeSuiteControls.ListView ListView1 
         Height          =   2055
         Left            =   120
         TabIndex        =   39
         Top             =   1200
         Width           =   5415
         _Version        =   851968
         _ExtentX        =   9551
         _ExtentY        =   3625
         _StockProps     =   77
         ForeColor       =   128
         BackColor       =   33023
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   33023
         Appearance      =   1
         UseVisualStyle  =   0   'False
         Arrange         =   2
      End
      Begin VB.Label Label14 
         BackColor       =   &H000080FF&
         Caption         =   "Label14"
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
         Left            =   2640
         TabIndex        =   40
         Top             =   3360
         Width           =   2895
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   13560
      Top             =   2280
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
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   3735
      Left            =   120
      TabIndex        =   1
      Top             =   4920
      Width           =   7095
      _Version        =   851968
      _ExtentX        =   12515
      _ExtentY        =   6588
      _StockProps     =   79
      Caption         =   "Müþteri Detay"
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
         Left            =   4800
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         Text            =   "Form1.frx":0000
         Top             =   1350
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
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "Text4"
         Top             =   960
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
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "Text3"
         Top             =   1680
         Width           =   3375
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
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "Text2"
         Top             =   1320
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
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   960
         Width           =   2175
      End
      Begin XtremeSuiteControls.GroupBox GroupBox3 
         Height          =   1575
         Left            =   120
         TabIndex        =   15
         Top             =   2040
         Width           =   6855
         _Version        =   851968
         _ExtentX        =   12091
         _ExtentY        =   2778
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
         Begin VB.TextBox Text8 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            DataField       =   "Musteri_KalanKahve"
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
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   29
            Text            =   "Text8"
            Top             =   1080
            Width           =   1095
         End
         Begin VB.TextBox Text7 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            DataField       =   "Musteri_Tuketilen_Bardak"
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
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   28
            Text            =   "Text7"
            Top             =   720
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
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   27
            Text            =   "Text6"
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label13 
            BackColor       =   &H000080FF&
            Caption         =   "Açýklama: Verilen ve tüketilen kahveden kalan bardak sayýsý."
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
            Left            =   3840
            TabIndex        =   21
            Top             =   1080
            Width           =   2895
         End
         Begin VB.Label Label12 
            BackColor       =   &H000080FF&
            Caption         =   "Açýklama: Tüketilen bardak miktarý."
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
            Left            =   3840
            TabIndex        =   20
            Top             =   750
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
            Left            =   3840
            TabIndex        =   19
            Top             =   360
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
            TabIndex        =   18
            Top             =   360
            Width           =   1845
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
            Left            =   450
            TabIndex        =   17
            Top             =   1110
            Width           =   1965
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
            Left            =   120
            TabIndex        =   16
            Top             =   720
            Width           =   2250
         End
      End
      Begin XtremeSuiteControls.PushButton PushButton4 
         Height          =   375
         Left            =   4920
         TabIndex        =   14
         Top             =   360
         Width           =   1455
         _Version        =   851968
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "En Son Kayýt >|"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton3 
         Height          =   375
         Left            =   3480
         TabIndex        =   13
         Top             =   360
         Width           =   1335
         _Version        =   851968
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Sonraki Kayýt >"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton2 
         Height          =   375
         Left            =   2040
         TabIndex        =   12
         Top             =   360
         Width           =   1335
         _Version        =   851968
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "< Önceki Kayýt"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   375
         Left            =   600
         TabIndex        =   11
         Top             =   360
         Width           =   1335
         _Version        =   851968
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "|< Ýlk Kayýt"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3780
         TabIndex        =   10
         Top             =   1320
         Width           =   915
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3600
         TabIndex        =   9
         Top             =   960
         Width           =   1050
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   525
         TabIndex        =   8
         Top             =   1680
         Width           =   630
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   405
         TabIndex        =   7
         Top             =   1320
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   1095
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12855
      _Version        =   851968
      _ExtentX        =   22675
      _ExtentY        =   8281
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
         Left            =   12000
         Top             =   3480
      End
      Begin XtremeSuiteControls.FlatEdit FlatEdit1 
         Height          =   255
         Left            =   2040
         TabIndex        =   3
         Top             =   4320
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
         Bindings        =   "Form1.frx":0006
         Height          =   3975
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   12615
         _ExtentX        =   22251
         _ExtentY        =   7011
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
         ColumnCount     =   6
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
         BeginProperty Column05 
            DataField       =   "GuncelKahve"
            Caption         =   "Güncelleme"
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
               ColumnWidth     =   1800
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   2805,166
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   2009,764
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   2805,166
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   2009,764
            EndProperty
         EndProperty
      End
      Begin XtremeSuiteControls.FlatEdit FlatEdit3 
         Height          =   255
         Left            =   6720
         TabIndex        =   43
         Top             =   4320
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
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Müþteri Telefonuyla Ara :"
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
         Left            =   4440
         TabIndex        =   42
         Top             =   4320
         Width           =   2115
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
         Left            =   11040
         TabIndex        =   4
         Top             =   4320
         Width           =   1395
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Müþteri Adýyla Ara :"
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
         Left            =   240
         TabIndex        =   2
         Top             =   4320
         Width           =   1650
      End
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   495
      Left            =   13320
      Top             =   3960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\SmitDorlasKahve_db.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\SmitDorlasKahve_db.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Musteriler_db"
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
   Begin VB.Menu ac 
      Caption         =   "Smit Dorlas Programý"
      Visible         =   0   'False
   End
   Begin VB.Menu musteri_yapilandýrma 
      Caption         =   "Müþteri Yapýlandýrma"
      Begin VB.Menu musteri_kayýt 
         Caption         =   "Müþteri Ekleme"
      End
      Begin VB.Menu musteri_duzenleme 
         Caption         =   "Müþteri Düzenleme"
      End
      Begin VB.Menu bos 
         Caption         =   "-"
      End
      Begin VB.Menu musteriyazdýr 
         Caption         =   "Müþteri Listesini Yazdýr"
      End
   End
   Begin VB.Menu yenile 
      Caption         =   "Yenile"
   End
   Begin VB.Menu kapat 
      Caption         =   "Programý Kapat"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub ac_Click()
Form1.Show
End Sub

Private Sub Command1_Click()
'Dim Dosyalar As ListItem
Set Dosyalar = ListView1.ListItems.Add
Dosyalar.SubItems(1) = Text11.Text
Dosyalar.SubItems(2) = Text9.Text
Dosyalar.SubItems(3) = Text12.Text
Dosyalar.SubItems(4) = Text10.Text
End Sub

Private Sub FlatEdit1_Change()
On Error Resume Next
Dim SQL1 As String
SQL1 = "select * from Musteriler_db where Musteri_Adi like '%" & FlatEdit1.Text & "%'"
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = SQL1
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1

End Sub


Private Sub FlatEdit3_Change()
On Error Resume Next
Dim SQL1 As String
SQL1 = "select * from Musteriler_db where Musteri_Tel like '%" & FlatEdit3.Text & "%'"
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = SQL1
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
End Sub

Private Sub Form_Activate()
Adodc1.Refresh
Adodc2.Refresh
DataGrid1.Refresh
Adodc1.Recordset.MoveLast
Timer2.Enabled = True
If RadioButton1.Value = True Then
FlatEdit2.Enabled = True
End If
Form1.Top = 195
Form1.Left = 1695
End Sub

Private Sub Form_Load()
Adodc1.Refresh
DataGrid1.Refresh
Adodc1.Recordset.MoveLast

ListView1.View = lvwReport
ListView1.ColumnHeaders.Add , , ""
ListView1.ColumnHeaders.Add , , "ID"
ListView1.ColumnHeaders.Add , , "Müþteri Adý"
ListView1.ColumnHeaders.Add , , "Müþteri Telefon"
ListView1.ColumnHeaders.Add , , "Kalan Kahve"
ListView1.ColumnHeaders(1).Width = 0
ListView1.ColumnHeaders(2).Width = 500
ListView1.ColumnHeaders(3).Width = 1500
ListView1.ColumnHeaders(4).Width = 1500
ListView1.ColumnHeaders(5).Width = 1200
End Sub

Private Sub Form_Resize()
On Error Resume Next
Adodc1.Recordset.MoveLast
Form1.Height = 9350
Form1.Width = 13365
End Sub


Private Sub ListBox1_Click()
On Error Resume Next
Dim SQL1 As String
SQL1 = "select * from Musteriler_db where Musteri_Adi like '%" & ListBox1.Text & "%'"
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = SQL1
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
End Sub


Private Sub kapat_Click()
Dim kapat As Integer
kapat = MsgBox("Programý kapatmak mý istiyorsunuz?", vbYesNo + vbInformation, "Kapatýlýyor...")
If kapat = vbYes Then
End
End If
End Sub

Private Sub musteri_duzenleme_Click()
On Error Resume Next
Form3.Show
Form1.Hide
ListViewSil
End Sub

Private Sub musteri_kayýt_Click()
On Error Resume Next
Form2.Show
Form1.Hide
ListViewSil
End Sub

Private Sub PushButton1_Click()
On Error Resume Next
Adodc1.Recordset.MoveFirst
End Sub

Private Sub PushButton2_Click()
On Error Resume Next
If Adodc1.Recordset.BOF = True Then
    Adodc1.Recordset.MoveLast
Else
    Adodc1.Recordset.MovePrevious
End If
End Sub

Private Sub PushButton3_Click()
On Error Resume Next
If Adodc1.Recordset.EOF = True Then
    Adodc1.Recordset.MoveFirst
Else
    Adodc1.Recordset.MoveNext
End If
End Sub

Private Sub PushButton4_Click()
On Error Resume Next
Adodc1.Recordset.MoveLast
End Sub
Private Sub ListViewSil()
On Error Resume Next
ListView1.ListItems.Remove (ListView1.ListItems.Count)
ListView1.ListItems.Remove (ListView1.ListItems.Count)
ListView1.ListItems.Remove (ListView1.ListItems.Count)
ListView1.ListItems.Remove (ListView1.ListItems.Count)
ListView1.ListItems.Remove (ListView1.ListItems.Count)
ListView1.ListItems.Remove (ListView1.ListItems.Count)
ListView1.ListItems.Remove (ListView1.ListItems.Count)
ListView1.ListItems.Remove (ListView1.ListItems.Count)
ListView1.ListItems.Remove (ListView1.ListItems.Count)
ListView1.ListItems.Remove (ListView1.ListItems.Count)
ListView1.ListItems.Remove (ListView1.ListItems.Count)
ListView1.ListItems.Remove (ListView1.ListItems.Count)
ListView1.ListItems.Remove (ListView1.ListItems.Count)
ListView1.ListItems.Remove (ListView1.ListItems.Count)
ListView1.ListItems.Remove (ListView1.ListItems.Count)
ListView1.ListItems.Remove (ListView1.ListItems.Count)
ListView1.ListItems.Remove (ListView1.ListItems.Count)
ListView1.ListItems.Remove (ListView1.ListItems.Count)
ListView1.ListItems.Remove (ListView1.ListItems.Count)
ListView1.ListItems.Remove (ListView1.ListItems.Count)
ListView1.ListItems.Remove (ListView1.ListItems.Count)
ListView1.ListItems.Remove (ListView1.ListItems.Count)
ListView1.ListItems.Remove (ListView1.ListItems.Count)
ListView1.ListItems.Remove (ListView1.ListItems.Count)
ListView1.ListItems.Remove (ListView1.ListItems.Count)
ListView1.ListItems.Remove (ListView1.ListItems.Count)
ListView1.ListItems.Remove (ListView1.ListItems.Count)
ListView1.ListItems.Remove (ListView1.ListItems.Count)
End Sub
Private Sub PushButton5_Click()
On Error Resume Next
Timer2.Enabled = True
ListViewSil
End Sub

Private Sub RadioButton1_Click()
FlatEdit2.Enabled = True
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Label2.Caption = "Toplam Kayýt : " & Adodc1.Recordset.RecordCount
Label14.Caption = "Kahvesi Azalan (" & ListView1.ListItems.Count & ") müþteri bulunmakta."

End Sub

Private Sub Timer2_Timer()
On Error Resume Next
Adodc2.Recordset.MoveNext
If RadioButton1.Value = True Then

If Text10.Text < FlatEdit2.Text + 1 Then

Set Dosyalar = ListView1.ListItems.Add
Dosyalar.SubItems(1) = Text11.Text
Dosyalar.SubItems(2) = Text9.Text
Dosyalar.SubItems(3) = Text12.Text
Dosyalar.SubItems(4) = Text10.Text

End If

End If

If Text11.Text = Adodc1.Recordset.RecordCount Then
Timer2.Enabled = False
Adodc2.Recordset.MoveFirst
End If
End Sub
Private Sub yenile_Click()
Adodc1.Refresh
DataGrid1.Refresh
Adodc1.Recordset.MoveLast
ListViewSil
FlatEdit3.Text = ""
FlatEdit1.Text = ""
End Sub
