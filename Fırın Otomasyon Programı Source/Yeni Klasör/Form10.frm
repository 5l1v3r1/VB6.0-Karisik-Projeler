VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form10 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gelir Tablosu"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10440
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   10440
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   480
      Top             =   6960
      Width           =   2535
      _ExtentX        =   4471
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Database.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Database.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Gelir_Tablosu"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command11 
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
      Left            =   7920
      TabIndex        =   52
      Top             =   6240
      Width           =   2295
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
      Left            =   8520
      TabIndex        =   50
      Top             =   120
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4920
      Top             =   7560
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   3720
      TabIndex        =   31
      Text            =   "Text9"
      Top             =   8040
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "-Gelir Tablosu"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   10215
      Begin VB.CommandButton Command12 
         Caption         =   "Gider Tablosunu Temizle"
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
         Left            =   7800
         TabIndex        =   53
         Top             =   4680
         Width           =   2295
      End
      Begin VB.TextBox Text17 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8640
         TabIndex        =   48
         Top             =   3600
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "T�m Alacaklarin Hesapla"
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
         Left            =   6120
         TabIndex        =   47
         Top             =   3600
         Width           =   2415
      End
      Begin VB.Frame Frame5 
         Caption         =   "-Kayit Islemi"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2520
         Left            =   120
         TabIndex        =   32
         Top             =   3020
         Visible         =   0   'False
         Width           =   3375
         Begin VB.TextBox Text16 
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
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   46
            Top             =   2160
            Width           =   2295
         End
         Begin VB.TextBox Text15 
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
            Left            =   960
            TabIndex        =   45
            Top             =   1860
            Width           =   1215
         End
         Begin VB.TextBox Text14 
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
            Left            =   960
            TabIndex        =   44
            Top             =   1560
            Width           =   735
         End
         Begin VB.TextBox Text13 
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
            Left            =   960
            TabIndex        =   43
            Top             =   1260
            Width           =   735
         End
         Begin VB.TextBox Text12 
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
            Left            =   960
            TabIndex        =   42
            Top             =   960
            Width           =   1695
         End
         Begin VB.TextBox Text11 
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
            Left            =   960
            TabIndex        =   41
            Top             =   660
            Width           =   2295
         End
         Begin VB.TextBox Text10 
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
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   40
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Alacagin:"
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
            Left            =   120
            TabIndex        =   39
            Top             =   1860
            Width           =   795
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Adet:"
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
            Left            =   420
            TabIndex        =   38
            Top             =   1260
            Width           =   465
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Agirligi:"
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
            TabIndex        =   37
            Top             =   1560
            Width           =   660
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "�r�n:"
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
            Left            =   420
            TabIndex        =   36
            Top             =   960
            Width           =   480
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Firma:"
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
            Left            =   360
            TabIndex        =   35
            Top             =   660
            Width           =   555
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
            Index           =   1
            Left            =   600
            TabIndex        =   34
            Top             =   360
            Width           =   285
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Zaman:"
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
            TabIndex        =   33
            Top             =   2160
            Width           =   675
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "-Islemleri"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2460
         Left            =   120
         TabIndex        =   17
         Top             =   3050
         Width           =   3375
         Begin VB.TextBox Text8 
            Appearance      =   0  'Flat
            DataField       =   "S�re"
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
            TabIndex        =   54
            Top             =   2040
            Width           =   2295
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
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            DataField       =   "Firma"
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
            TabIndex        =   22
            Top             =   540
            Width           =   2295
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            DataField       =   "�r�n"
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
            TabIndex        =   21
            Top             =   840
            Width           =   1695
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            DataField       =   "Adet"
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
            TabIndex        =   20
            Top             =   1140
            Width           =   615
         End
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            DataField       =   "Agirlik"
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
            TabIndex        =   19
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
            DataField       =   "Alacagin"
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
            TabIndex        =   18
            Top             =   1740
            Width           =   855
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Zaman:"
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
            Index           =   0
            Left            =   240
            TabIndex        =   30
            Top             =   2040
            Width           =   675
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
            Index           =   0
            Left            =   600
            TabIndex        =   29
            Top             =   240
            Width           =   285
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Firma:"
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
            Index           =   0
            Left            =   360
            TabIndex        =   28
            Top             =   540
            Width           =   555
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "�r�n:"
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
            Index           =   0
            Left            =   420
            TabIndex        =   27
            Top             =   840
            Width           =   480
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Agirligi:"
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
            Index           =   0
            Left            =   240
            TabIndex        =   26
            Top             =   1440
            Width           =   660
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Adet:"
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
            Index           =   0
            Left            =   420
            TabIndex        =   25
            Top             =   1140
            Width           =   465
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Alacagin:"
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
            Index           =   0
            Left            =   120
            TabIndex        =   24
            Top             =   1740
            Width           =   795
         End
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
         Left            =   3600
         TabIndex        =   16
         Top             =   4245
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
         Left            =   3600
         TabIndex        =   15
         Top             =   3870
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
         Left            =   3600
         TabIndex        =   14
         Top             =   3480
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
         Left            =   4800
         TabIndex        =   13
         Top             =   3480
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
         Left            =   3600
         TabIndex        =   12
         Top             =   3100
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
         Left            =   4800
         TabIndex        =   11
         Top             =   4245
         Width           =   1210
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
         Left            =   3600
         TabIndex        =   10
         Top             =   4635
         Width           =   2415
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
         Left            =   3600
         TabIndex        =   7
         Top             =   5040
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
            TabIndex        =   8
            Top             =   200
            Width           =   1575
         End
         Begin VB.Label Label7 
            Caption         =   "Firma:"
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
            TabIndex        =   9
            Top             =   220
            Width           =   735
         End
      End
      Begin VB.Frame Frame4 
         Height          =   495
         Left            =   6120
         TabIndex        =   4
         Top             =   3050
         Width           =   1815
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Toplam Girdi:"
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
            TabIndex        =   6
            Top             =   180
            Width           =   1170
         End
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
            Left            =   1440
            TabIndex        =   5
            Top             =   195
            Width           =   135
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Form10.frx":0000
         Height          =   2775
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   4895
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
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
         Caption         =   "Gelir Tablosu"
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
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "TL"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9720
         TabIndex        =   49
         Top             =   3645
         Width           =   240
      End
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Firma/D�kkan/Toptanci Ekle"
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
      TabIndex        =   1
      Top             =   600
      Width           =   2775
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Gelir Tablosuna Girdi Ekle"
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
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label13 
      Caption         =   "Label13"
      Height          =   255
      Left            =   8760
      TabIndex        =   55
      Top             =   600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "Label12"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   51
      Top             =   360
      Width           =   5415
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Command1_Click()
ToplamAlacak
End Sub
Public Function dblRoundOff(ByVal X As Double, ByVal N As Integer) As Double
dblRoundOff = CDbl(Int((X * 10 ^ N) + 0.5) / 10 ^ N)
End Function
Private Sub ToplamAlacak()
For i = 1 To Adodc1.Recordset.RecordCount
Adodc1.Recordset.AbsolutePosition = i
T1 = Round((Adodc1.Recordset.Fields("Alacagin") * 10000), 2) * 100 'yaz�lan toplam
T1 = dblRoundOff(T1, 2)
T2 = Val(T2) + Val(T1)
Toplam = Round((Val(T2) / 10000), 2) / 100
Toplam = dblRoundOff(Toplam, 2)
Next
Text17.Text = Toplam
End Sub

Private Sub Command10_Click()
On Error Resume Next
Text10.Text = Text9.Text + 1
Text16.Text = "Saat: " & TimeValue(Now) & " | Tarih: " & Format(Date, "dd.mmmm.yyyy")
Command6.Enabled = True
Command3.Enabled = False
Frame2.Enabled = False
Frame5.Visible = True
End Sub

Private Sub Command11_Click()
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

ExcelNesne.Application.Cells(1, 1).Font.Color = vbBlack
ExcelNesne.Application.Cells(1, 1).ColumnWidth = 5
 ExcelNesne.Application.Cells(1, 1).Value = "ID"
 
 ExcelNesne.Application.Cells(1, 2).Font.Size = 11
ExcelNesne.Application.Cells(1, 2).Font.Bold = True
ExcelNesne.Application.Cells(1, 2).Font.Underline = True
 ExcelNesne.Application.Cells(1, 2).Font.Color = vbBlack
ExcelNesne.Application.Cells(1, 2).ColumnWidth = 20
 ExcelNesne.Application.Cells(1, 2).Value = "Firma"
 
 ExcelNesne.Application.Cells(1, 3).Font.Size = 11
ExcelNesne.Application.Cells(1, 3).Font.Bold = True
ExcelNesne.Application.Cells(1, 3).Font.Underline = True
 ExcelNesne.Application.Cells(1, 3).Font.Color = vbBlack
ExcelNesne.Application.Cells(1, 3).ColumnWidth = 20
 ExcelNesne.Application.Cells(1, 3).Value = "�r�n"
 
 ExcelNesne.Application.Cells(1, 4).Font.Size = 11
ExcelNesne.Application.Cells(1, 4).Font.Bold = True
ExcelNesne.Application.Cells(1, 4).Font.Underline = True
 ExcelNesne.Application.Cells(1, 4).Font.Color = vbBlack
ExcelNesne.Application.Cells(1, 4).ColumnWidth = 10
 ExcelNesne.Application.Cells(1, 4).Value = "Adet"
 
 ExcelNesne.Application.Cells(1, 5).Font.Size = 11
ExcelNesne.Application.Cells(1, 5).Font.Bold = True
ExcelNesne.Application.Cells(1, 5).Font.Underline = True
 ExcelNesne.Application.Cells(1, 5).Font.Color = vbBlack
ExcelNesne.Application.Cells(1, 5).ColumnWidth = 12
 ExcelNesne.Application.Cells(1, 5).Value = "Agirligi(kg)"
 
 ExcelNesne.Application.Cells(1, 6).Font.Size = 11
ExcelNesne.Application.Cells(1, 6).Font.Bold = True
ExcelNesne.Application.Cells(1, 6).Font.Underline = True
 ExcelNesne.Application.Cells(1, 6).Font.Color = vbBlack
ExcelNesne.Application.Cells(1, 6).ColumnWidth = 15
 ExcelNesne.Application.Cells(1, 6).Value = "Alacagin(TL)"
 
 ExcelNesne.Application.Cells(1, 7).Font.Size = 11
ExcelNesne.Application.Cells(1, 7).Font.Bold = True
ExcelNesne.Application.Cells(1, 7).Font.Underline = True
 ExcelNesne.Application.Cells(1, 7).Font.Color = vbBlack
ExcelNesne.Application.Cells(1, 7).ColumnWidth = 30
 ExcelNesne.Application.Cells(1, 7).Value = "Zaman"

' ExcelNesne.Application.Cells(1, 8).Font.Size = 11
'ExcelNesne.Application.Cells(1, 8).Font.Bold = True
'ExcelNesne.Application.Cells(1, 8).Font.Underline = True
 '  ExcelNesne.Application.Cells(1, 8).Font.Color = vbBlack
ExcelNesne.Application.Cells(1, 8).ColumnWidth = 20
 'ExcelNesne.Application.Cells(1, 8).Value = "Toplam Alacagin"

i = 1
Adodc1.Recordset.MoveFirst
a = Label9.Caption
Do While Not Adodc1.Recordset.EOF = True
i = i + 1
b = a + 5

ExcelNesne.Application.Cells(i, 1).Value = Adodc1.Recordset.Fields("ID")
ExcelNesne.Application.Cells(i, 2).Value = Adodc1.Recordset.Fields("Firma")
ExcelNesne.Application.Cells(i, 3).Value = Adodc1.Recordset.Fields("�r�n")
ExcelNesne.Application.Cells(i, 4).Value = Adodc1.Recordset.Fields("Adet")
ExcelNesne.Application.Cells(i, 5).Value = Adodc1.Recordset.Fields("Agirlik")
ExcelNesne.Application.Cells(i, 6).Value = Adodc1.Recordset.Fields("Alacagin")
ExcelNesne.Application.Cells(i, 7).Value = Adodc1.Recordset.Fields("S�re")

ExcelNesne.Application.Cells(b, 8).Value = "Toplam Alacagin: " & Text17.Text & " TL"
ExcelNesne.Application.Cells(b, 1).Value = "GELIR TABLOSU | " & Form4.Text1.Text

Adodc1.Recordset.MoveNext
Loop
MsgBox "Microsoft Excel'e Aktarildi Bekleniyor...", vbInformation, "Bildiri;"
Adodc1.Recordset.MoveLast
End Sub

Private Sub Command12_Click()
On Error Resume Next
If Adodc1.Recordset.RecordCount = 0 Then
MsgBox "KAYIT YOK"
End If
msg = MsgBox("B�t�n Kayitlar Silinsin mi?" & " Toplam Kayit: " & Label9.Caption, vbInformation + vbYesNo, "Silme �slemi;")
If msg = vbYes Then
For a = 1 To 30
Adodc1.Recordset.Delete
Adodc1.Recordset.MovePrevious
Next a
End If
End Sub

Private Sub Command2_Click()
On Error Resume Next
Unload Me
Form2.Show
End Sub

Private Sub Command20_Click()
On Error Resume Next
Form2.Hide
Form3.Hide
Form4.Hide
Form5.Hide
Form6.Hide
Form8.Hide
Form9.Hide
Form7.Show
End Sub

Private Sub Form_Load()
On Error Resume Next
Form10.Caption = "Gelir Tablosu | " & Form4.Text1.Text
Label12.Caption = Form4.Text1.Text & " | Gelir Tablosu"
ToplamAlacak
Label13.Caption = Label9.Caption + 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Unload Me
Form2.Show
End Sub
Private Sub Command3_Click()
On Error Resume Next
Text10.Text = Text9.Text + 1
Text16.Text = "Saat: " & TimeValue(Now) & " | Tarih: " & Format(Date, "dd.mmmm.yyyy")
Command6.Enabled = True
Command3.Enabled = False
Frame2.Enabled = False
Frame5.Visible = True
End Sub
Private Sub Command4_Click()
On Error Resume Next
Adodc1.Recordset.MoveFirst
End Sub
Private Sub Command5_Click()
On Error Resume Next
Adodc1.Recordset.MoveLast
End Sub
Private Sub Text7_Change()
On Error Resume Next
Dim sql As String
sql = "select * from Gelir_Tablosu where Firma like '%" & Text7.Text & "%'"
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = sql
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
End Sub
Private Sub Timer1_Timer()
On Error Resume Next
Text9.Text = Adodc1.Recordset.RecordCount
Label9.Caption = Text9.Text
End Sub

Private Sub Command6_Click()
On Error Resume Next
kaydet = MsgBox("[" & Text11.Text & "] isimli firma kaydedilsin mi?", vbYesNo + vbInformation, "Kaydedilsin mi?")
If kaydet = vbYes Then
Adodc1.Recordset.AddNew
Adodc1.Recordset!ID = Text10.Text
Adodc1.Recordset!Firma = Text11.Text
Adodc1.Recordset!�r�n = Text12.Text
Adodc1.Recordset!Adet = Text13.Text
Adodc1.Recordset!Agirlik = Text14.Text
Adodc1.Recordset!Alacagin = Text15.Text
Adodc1.Recordset!S�re = Text16.Text

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Command3.Enabled = True
Command6.Enabled = False
Frame5.Visible = False
'DataGrid1.Refresh
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
Frame5.Visible = False
End If
End Sub
Private Sub Command7_Click()
On Error Resume Next
If Adodc1.Recordset.EOF = True Then
    Adodc1.Recordset.MoveFirst
Else
    Adodc1.Recordset.MoveNext
End If
End Sub
Private Sub Command8_Click()
On Error Resume Next
silme = MsgBox("Se�ili kullanici silinsin mi?", vbYesNo + vbInformation, "Silme ��lemi")
If silme = vbYes Then
Adodc1.Recordset.Delete
Adodc1.Recordset.MovePrevious
ToplamAlacak
End If
End Sub
Private Sub Command9_Click()
On Error Resume Next
If Adodc1.Recordset.BOF = True Then
    Adodc1.Recordset.MoveLast
Else
    Adodc1.Recordset.MovePrevious
End If
End Sub


