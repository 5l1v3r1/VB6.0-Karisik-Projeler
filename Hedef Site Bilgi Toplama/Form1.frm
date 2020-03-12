VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "ActiveX.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   Caption         =   "Danger: Hedef Site Bilgi Toplama | _DangerOusMaN_"
   ClientHeight    =   10530
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   16890
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   9
      Charset         =   162
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10530
   ScaleWidth      =   16890
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Caption         =   "URL Root"
      Height          =   1935
      Left            =   240
      TabIndex        =   118
      Top             =   9840
      Width           =   2175
      Begin VB.TextBox Text60 
         Height          =   735
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   120
         Text            =   "Form1.frx":8D25A
         Top             =   1080
         Width           =   1935
      End
      Begin XtremeSuiteControls.WebBrowser web5 
         Height          =   735
         Left            =   120
         TabIndex        =   119
         Top             =   240
         Width           =   1935
         _Version        =   851968
         _ExtentX        =   3413
         _ExtentY        =   1296
         _StockProps     =   173
         BackColor       =   -2147483643
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Exploit Bulucu"
      Height          =   2175
      Left            =   8640
      TabIndex        =   103
      Top             =   7560
      Width           =   3855
      Begin VB.TextBox Text59 
         Height          =   330
         Left            =   2400
         TabIndex        =   107
         Text            =   $"Form1.frx":8D261
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox Text58 
         Height          =   330
         Left            =   2400
         TabIndex        =   106
         Text            =   "http://www.exploit-db.com/exploits/"
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox Text57 
         Enabled         =   0   'False
         Height          =   735
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   105
         Text            =   "Form1.frx":8D28D
         Top             =   1200
         Width           =   2175
      End
      Begin XtremeSuiteControls.WebBrowser WebBrowser1 
         Height          =   735
         Left            =   120
         TabIndex        =   104
         Top             =   360
         Width           =   2175
         _Version        =   851968
         _ExtentX        =   3836
         _ExtentY        =   1296
         _StockProps     =   173
         BackColor       =   -2147483643
      End
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   12120
      Top             =   5760
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   7920
      Top             =   7800
   End
   Begin XtremeSuiteControls.GroupBox GroupBox13 
      Height          =   2175
      Left            =   120
      TabIndex        =   89
      Top             =   7560
      Width           =   8415
      _Version        =   851968
      _ExtentX        =   14843
      _ExtentY        =   3836
      _StockProps     =   79
      Caption         =   "Site Yazýlým Dili ve Script"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Begin VB.Timer Timer5 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   7800
         Top             =   1200
      End
      Begin VB.Timer Timer4 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   7800
         Top             =   720
      End
      Begin VB.TextBox Text56 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5160
         TabIndex        =   98
         Text            =   "</"
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox Text55 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5160
         TabIndex        =   97
         Text            =   """>SMF"
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox Text54 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5160
         TabIndex        =   96
         Text            =   """ />"
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox Text53 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5160
         TabIndex        =   95
         Text            =   """Joomla"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox Text52 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5160
         TabIndex        =   94
         Text            =   """ />"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox Text51 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5160
         TabIndex        =   93
         Text            =   """vBulletin"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Text50 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   2280
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   92
         Text            =   "Form1.frx":8D294
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox Text49 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   91
         Text            =   "Form1.frx":8D29B
         Top             =   1080
         Width           =   2055
      End
      Begin XtremeSuiteControls.WebBrowser Web4 
         Height          =   735
         Left            =   120
         TabIndex        =   90
         Top             =   240
         Width           =   2055
         _Version        =   851968
         _ExtentX        =   3625
         _ExtentY        =   1296
         _StockProps     =   173
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ScriptErrorsSuppressed=   -1  'True
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox9 
      Height          =   255
      Left            =   8640
      TabIndex        =   67
      Top             =   1800
      Width           =   3015
      _Version        =   851968
      _ExtentX        =   5318
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "GroupBox9"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
   End
   Begin VB.Frame Frame3 
      Caption         =   "Port Tarama"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   12960
      TabIndex        =   49
      Top             =   2760
      Width           =   4095
      Begin VB.TextBox Text48 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2520
         TabIndex        =   66
         Text            =   "</TD></TR><!--  -->"
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox Text47 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2520
         TabIndex        =   65
         Text            =   "rltext>"
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox Text46 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   64
         Text            =   $"Form1.frx":8D2A2
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox Text45 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   63
         Text            =   $"Form1.frx":8D2B8
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox Text44 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   62
         Text            =   "</SPA"
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox Text43 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   61
         Text            =   "FONT-WEIGHT: bold"">"
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox Text42 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   60
         Text            =   $"Form1.frx":8D2D7
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox Text41 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   59
         Text            =   $"Form1.frx":8D2F9
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox Text40 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   58
         Text            =   $"Form1.frx":8D3FC
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox Text39 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   57
         Text            =   $"Form1.frx":8D40C
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox Text38 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3360
         TabIndex        =   56
         Text            =   "Text38"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox Text37 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3360
         TabIndex        =   55
         Text            =   "Text37"
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox Text36 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2520
         TabIndex        =   54
         Text            =   "Text36"
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox Text35 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2520
         TabIndex        =   53
         Text            =   "Text35"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox Text34 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2520
         TabIndex        =   52
         Text            =   "Text34"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox Text33 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   51
         Text            =   "Form1.frx":8D42B
         Top             =   960
         Width           =   2295
      End
      Begin XtremeSuiteControls.WebBrowser Web3 
         Height          =   735
         Left            =   120
         TabIndex        =   50
         Top             =   240
         Width           =   2295
         _Version        =   851968
         _ExtentX        =   4048
         _ExtentY        =   1296
         _StockProps     =   173
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ScriptErrorsSuppressed=   -1  'True
      End
   End
   Begin VB.TextBox Text32 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   12960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   47
      Top             =   8400
      Width           =   2175
   End
   Begin VB.Frame Frame2 
      Caption         =   "Reserve IP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   12960
      TabIndex        =   40
      Top             =   0
      Width           =   4095
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   3120
         Top             =   2040
      End
      Begin VB.TextBox Text31 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2760
         TabIndex        =   46
         Text            =   "DISABLED0"
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox Text30 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   45
         Text            =   "Form1.frx":8D432
         Top             =   1440
         Width           =   2535
      End
      Begin VB.TextBox Text29 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   44
         Text            =   "Home Reverse DNS IP Neighborhoods Selecting Hosting Keyword Rankings"
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox Text28 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   43
         Text            =   $"Form1.frx":8D439
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox Text27 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   1440
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   42
         Text            =   "Form1.frx":8D45B
         Top             =   240
         Width           =   1215
      End
      Begin XtremeSuiteControls.WebBrowser Web2 
         Height          =   1095
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   1931
         _StockProps     =   173
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ScriptErrorsSuppressed=   -1  'True
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Whois"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   12960
      TabIndex        =   20
      Top             =   5520
      Width           =   4095
      Begin VB.TextBox Text26 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   39
         Text            =   "Kýsa Sorgu Sonucu"
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox Text25 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   38
         Text            =   "Sorgu Sonucu /"
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox Text24 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   37
         Text            =   "Diðer Uzantýlar"
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox Text23 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   36
         Text            =   $"Form1.frx":8D462
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox Text22 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   35
         Text            =   $"Form1.frx":8D47E
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox Text21 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   34
         Text            =   $"Form1.frx":8D49A
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox Text20 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   33
         Text            =   $"Form1.frx":8D4B2
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox Text19 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   32
         Text            =   $"Form1.frx":8D4CA
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox Text18 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   31
         Text            =   $"Form1.frx":8D4D9
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox Text17 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   30
         Text            =   "Son Güncelleme Tarihi : "
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox Text16 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   29
         Text            =   $"Form1.frx":8D4E8
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox Text15 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   28
         Text            =   $"Form1.frx":8D506
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox Text14 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   $"Form1.frx":8D51B
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox Text13 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   26
         Text            =   $"Form1.frx":8D530
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox Text12 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   $"Form1.frx":8D54C
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox Text11 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "Sorgulanan Alan Adý : "
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "göster"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   23
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Text            =   "Form1.frx":8D567
         Top             =   1440
         Width           =   1935
      End
      Begin XtremeSuiteControls.WebBrowser Web1 
         Height          =   975
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   1935
         _Version        =   851968
         _ExtentX        =   3413
         _ExtentY        =   1720
         _StockProps     =   173
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ScriptErrorsSuppressed=   -1  'True
      End
   End
   Begin XtremeSuiteControls.PushButton Command1 
      Height          =   255
      Left            =   5160
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
      _Version        =   851968
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Sorgula"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit Text1 
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   1320
      Width           =   4095
      _Version        =   851968
      _ExtentX        =   7223
      _ExtentY        =   450
      _StockProps     =   77
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   11415
      _Version        =   851968
      _ExtentX        =   20135
      _ExtentY        =   8493
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoResizeClient=   0   'False
      Appearance      =   6
      Color           =   16
      ItemCount       =   4
      SelectedItem    =   2
      Item(0).Caption =   "Whois [Sunucu Bilgileri]"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "GroupBox1"
      Item(1).Caption =   "ReserveIP / Ping ve Port Taramasý"
      Item(1).ControlCount=   3
      Item(1).Control(0)=   "GroupBox3"
      Item(1).Control(1)=   "GroupBox5"
      Item(1).Control(2)=   "GroupBox4"
      Item(2).Caption =   "Hedef Site Web Analizi"
      Item(2).ControlCount=   4
      Item(2).Control(0)=   "GroupBox10"
      Item(2).Control(1)=   "GroupBox16"
      Item(2).Control(2)=   "PushButton1"
      Item(2).Control(3)=   "PushButton2"
      Item(3).Caption =   "Hakkýnda"
      Item(3).ControlCount=   1
      Item(3).Control(0)=   "Image2"
      Begin XtremeSuiteControls.PushButton PushButton2 
         Height          =   375
         Left            =   9360
         TabIndex        =   122
         Top             =   4320
         Width           =   1935
         _Version        =   851968
         _ExtentX        =   3413
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Programý Kapat"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   2
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   375
         Left            =   4320
         TabIndex        =   121
         Top             =   4320
         Width           =   1935
         _Version        =   851968
         _ExtentX        =   3413
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Bilgileri Kaydet"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   2
      End
      Begin XtremeSuiteControls.GroupBox GroupBox16 
         Height          =   3855
         Left            =   4320
         TabIndex        =   113
         Top             =   360
         Width           =   6975
         _Version        =   851968
         _ExtentX        =   12303
         _ExtentY        =   6800
         _StockProps     =   79
         Appearance      =   2
         Begin XtremeSuiteControls.GroupBox GroupBox17 
            Height          =   3135
            Left            =   120
            TabIndex        =   114
            Top             =   600
            Width           =   6735
            _Version        =   851968
            _ExtentX        =   11880
            _ExtentY        =   5530
            _StockProps     =   79
            Caption         =   "-Sitenin URL Adresleri"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Begin XtremeSuiteControls.ListBox ListBox2 
               Height          =   2775
               Left            =   120
               TabIndex        =   115
               Top             =   240
               Width           =   6495
               _Version        =   851968
               _ExtentX        =   11456
               _ExtentY        =   4895
               _StockProps     =   77
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   162
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Appearance      =   2
               UseVisualStyle  =   0   'False
            End
         End
         Begin VB.Label Label20 
            Caption         =   $"Form1.frx":8D56D
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   720
            TabIndex        =   117
            Top             =   150
            Width           =   6225
         End
         Begin VB.Label Label19 
            Caption         =   "Bilgi :"
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
            Left            =   120
            TabIndex        =   116
            Top             =   150
            Width           =   615
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox10 
         Height          =   4335
         Left            =   120
         TabIndex        =   83
         Top             =   360
         Width           =   4095
         _Version        =   851968
         _ExtentX        =   7223
         _ExtentY        =   7646
         _StockProps     =   79
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   2
         Begin XtremeSuiteControls.GroupBox GroupBox14 
            Height          =   3015
            Left            =   120
            TabIndex        =   99
            Top             =   1200
            Width           =   3855
            _Version        =   851968
            _ExtentX        =   6800
            _ExtentY        =   5318
            _StockProps     =   79
            Caption         =   "-Varsa Önceden Bulunan Scriptin Açýklarý"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   2
            Begin XtremeSuiteControls.GroupBox GroupBox15 
               Height          =   2655
               Left            =   120
               TabIndex        =   100
               Top             =   240
               Width           =   3615
               _Version        =   851968
               _ExtentX        =   6376
               _ExtentY        =   4683
               _StockProps     =   79
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   162
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Appearance      =   2
               Begin XtremeSuiteControls.ListBox ListBox1 
                  Height          =   2115
                  Left            =   120
                  TabIndex        =   101
                  Top             =   240
                  Width           =   3375
                  _Version        =   851968
                  _ExtentX        =   5953
                  _ExtentY        =   3731
                  _StockProps     =   77
                  BackColor       =   -2147483643
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   162
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Appearance      =   2
                  UseVisualStyle  =   0   'False
               End
               Begin VB.Label Label18 
                  Caption         =   "Bulunan Exploit :"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   162
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   120
                  TabIndex        =   102
                  Top             =   2380
                  Width           =   3375
               End
            End
         End
         Begin XtremeSuiteControls.GroupBox GroupBox12 
            Height          =   975
            Left            =   120
            TabIndex        =   84
            Top             =   120
            Width           =   3855
            _Version        =   851968
            _ExtentX        =   6800
            _ExtentY        =   1720
            _StockProps     =   79
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   2
            Begin VB.CommandButton Command3 
               Caption         =   "Manuel Ara"
               Height          =   255
               Left            =   2520
               TabIndex        =   108
               Top             =   240
               Width           =   1215
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit5 
               Height          =   255
               Left            =   1320
               TabIndex        =   88
               Top             =   600
               Width           =   2415
               _Version        =   851968
               _ExtentX        =   4260
               _ExtentY        =   450
               _StockProps     =   77
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   162
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Appearance      =   2
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit4 
               Height          =   255
               Left            =   1320
               TabIndex        =   87
               Top             =   240
               Width           =   735
               _Version        =   851968
               _ExtentX        =   1296
               _ExtentY        =   450
               _StockProps     =   77
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   162
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Alignment       =   2
               Appearance      =   2
               UseVisualStyle  =   0   'False
            End
            Begin VB.Label Label17 
               Caption         =   "Script Adý   :"
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
               TabIndex        =   86
               Top             =   600
               Width           =   1200
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               Caption         =   "Yazýlým Dili :"
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
               TabIndex        =   85
               Top             =   240
               Width           =   1065
            End
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   4335
         Left            =   -69880
         TabIndex        =   3
         Top             =   360
         Visible         =   0   'False
         Width           =   11175
         _Version        =   851968
         _ExtentX        =   19711
         _ExtentY        =   7646
         _StockProps     =   79
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   2
         Begin XtremeSuiteControls.GroupBox GroupBox2 
            Height          =   2775
            Left            =   120
            TabIndex        =   18
            Top             =   1440
            Width           =   10935
            _Version        =   851968
            _ExtentX        =   19288
            _ExtentY        =   4895
            _StockProps     =   79
            Caption         =   "-Whois Sonucu"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   2
            BorderStyle     =   1
            Begin XtremeSuiteControls.FlatEdit Text10 
               Height          =   2535
               Left            =   0
               TabIndex        =   19
               Top             =   240
               Width           =   10935
               _Version        =   851968
               _ExtentX        =   19288
               _ExtentY        =   4471
               _StockProps     =   77
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   162
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MultiLine       =   -1  'True
               ScrollBars      =   2
               Appearance      =   2
               UseVisualStyle  =   0   'False
            End
         End
         Begin XtremeSuiteControls.FlatEdit Text9 
            Height          =   255
            Left            =   4800
            TabIndex        =   16
            Top             =   1080
            Width           =   2175
            _Version        =   851968
            _ExtentX        =   3836
            _ExtentY        =   450
            _StockProps     =   77
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit Text7 
            Height          =   255
            Left            =   2400
            TabIndex        =   15
            Top             =   1080
            Width           =   2175
            _Version        =   851968
            _ExtentX        =   3836
            _ExtentY        =   450
            _StockProps     =   77
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit Text6 
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   1080
            Width           =   2055
            _Version        =   851968
            _ExtentX        =   3625
            _ExtentY        =   450
            _StockProps     =   77
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit Text5 
            Height          =   255
            Left            =   4800
            TabIndex        =   13
            Top             =   480
            Width           =   2175
            _Version        =   851968
            _ExtentX        =   3836
            _ExtentY        =   450
            _StockProps     =   77
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit Text4 
            Height          =   255
            Left            =   2400
            TabIndex        =   12
            Top             =   480
            Width           =   2175
            _Version        =   851968
            _ExtentX        =   3836
            _ExtentY        =   450
            _StockProps     =   77
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit Text3 
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   480
            Width           =   2055
            _Version        =   851968
            _ExtentX        =   3625
            _ExtentY        =   450
            _StockProps     =   77
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit Text8 
            Height          =   855
            Left            =   7200
            TabIndex        =   17
            Top             =   480
            Width           =   3855
            _Version        =   851968
            _ExtentX        =   6800
            _ExtentY        =   1508
            _StockProps     =   77
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MultiLine       =   -1  'True
            ScrollBars      =   2
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Sorgulanan Alan Adý"
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
            TabIndex        =   10
            Top             =   240
            Width           =   1740
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Oluþturulma Tarihi"
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
            Left            =   2400
            TabIndex        =   9
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Bitiþ Tarihi"
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
            Left            =   4800
            TabIndex        =   8
            Top             =   240
            Width           =   900
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Son Güncelleme Tarihi"
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
            TabIndex        =   7
            Top             =   840
            Width           =   1935
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Durumu"
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
            Left            =   2400
            TabIndex        =   6
            Top             =   840
            Width           =   690
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Ýsim Sunucularý"
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
            Left            =   7200
            TabIndex        =   5
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Kimindir Sunucusu"
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
            Left            =   4800
            TabIndex        =   4
            Top             =   840
            Width           =   1590
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox3 
         Height          =   2295
         Left            =   -64360
         TabIndex        =   68
         Top             =   360
         Visible         =   0   'False
         Width           =   5655
         _Version        =   851968
         _ExtentX        =   9975
         _ExtentY        =   4048
         _StockProps     =   79
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   2
         Begin XtremeSuiteControls.GroupBox GroupBox8 
            Height          =   2055
            Left            =   120
            TabIndex        =   69
            Top             =   120
            Width           =   5415
            _Version        =   851968
            _ExtentX        =   9551
            _ExtentY        =   3625
            _StockProps     =   79
            Caption         =   "-Port Tarama Sonuçlarý "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   2
            Begin XtremeSuiteControls.ListView ListView1 
               Height          =   1695
               Left            =   120
               TabIndex        =   70
               Top             =   240
               Width           =   5175
               _Version        =   851968
               _ExtentX        =   9128
               _ExtentY        =   2990
               _StockProps     =   77
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   162
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               View            =   3
               Appearance      =   2
               UseVisualStyle  =   0   'False
            End
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
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
            TabIndex        =   71
            Top             =   240
            Width           =   60
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox5 
         Height          =   4335
         Left            =   -69880
         TabIndex        =   72
         Top             =   360
         Visible         =   0   'False
         Width           =   5415
         _Version        =   851968
         _ExtentX        =   9551
         _ExtentY        =   7646
         _StockProps     =   79
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   2
         Begin VB.Timer Timer2 
            Enabled         =   0   'False
            Interval        =   20
            Left            =   7080
            Top             =   3840
         End
         Begin XtremeSuiteControls.GroupBox GroupBox7 
            Height          =   700
            Left            =   120
            TabIndex        =   73
            Top             =   3550
            Width           =   5175
            _Version        =   851968
            _ExtentX        =   9128
            _ExtentY        =   1235
            _StockProps     =   79
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Begin XtremeSuiteControls.FlatEdit FlatEdit2 
               Height          =   255
               Left            =   120
               TabIndex        =   74
               Top             =   360
               Width           =   4935
               _Version        =   851968
               _ExtentX        =   8705
               _ExtentY        =   450
               _StockProps     =   77
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   162
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Appearance      =   2
               UseVisualStyle  =   0   'False
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "Hedefin Tahmini Ýþletim Sistemi :"
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
               TabIndex        =   75
               Top             =   150
               Width           =   2850
            End
         End
         Begin XtremeSuiteControls.GroupBox GroupBox6 
            Height          =   3015
            Left            =   120
            TabIndex        =   76
            Top             =   120
            Width           =   5175
            _Version        =   851968
            _ExtentX        =   9128
            _ExtentY        =   5318
            _StockProps     =   79
            Caption         =   "-Sonuc"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   2
            Begin XtremeSuiteControls.FlatEdit FlatEdit1 
               Height          =   2295
               Left            =   120
               TabIndex        =   77
               Top             =   240
               Width           =   4935
               _Version        =   851968
               _ExtentX        =   8705
               _ExtentY        =   4048
               _StockProps     =   77
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   162
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MultiLine       =   -1  'True
               ScrollBars      =   2
            End
            Begin XtremeSuiteControls.FlatEdit FlatEdit3 
               Height          =   255
               Left            =   1440
               TabIndex        =   78
               Top             =   2640
               Width           =   1935
               _Version        =   851968
               _ExtentX        =   3413
               _ExtentY        =   450
               _StockProps     =   77
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   162
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Alignment       =   2
               Appearance      =   2
               UseVisualStyle  =   0   'False
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "Site IP Adresi: "
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
               TabIndex        =   80
               Top             =   2640
               Width           =   1305
            End
            Begin VB.Label Label13 
               Caption         =   "TTL="
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
               Left            =   4080
               TabIndex        =   79
               Top             =   2640
               Width           =   795
            End
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Bilgi :"
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
            Left            =   120
            TabIndex        =   82
            Top             =   3130
            Width           =   540
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "TTL Deðeri hedef sistemin uzaklýðý ve iþletim sistemi  hakkýnda bilgi verir."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   81
            Top             =   3340
            Width           =   5220
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox4 
         Height          =   1935
         Left            =   -64360
         TabIndex        =   109
         Top             =   2760
         Visible         =   0   'False
         Width           =   5655
         _Version        =   851968
         _ExtentX        =   9975
         _ExtentY        =   3413
         _StockProps     =   79
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   2
         Begin XtremeSuiteControls.GroupBox GroupBox11 
            Height          =   1695
            Left            =   120
            TabIndex        =   110
            Top             =   120
            Width           =   5415
            _Version        =   851968
            _ExtentX        =   9551
            _ExtentY        =   2990
            _StockProps     =   79
            Caption         =   "-ReserveIP [Sunucudaki Diðer Siteler]"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Begin XtremeSuiteControls.ListBox List1 
               Height          =   1180
               Left            =   120
               TabIndex        =   111
               Top             =   240
               Width           =   5175
               _Version        =   851968
               _ExtentX        =   9128
               _ExtentY        =   2081
               _StockProps     =   77
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   162
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Appearance      =   2
               UseVisualStyle  =   0   'False
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Sunucudaki Toplam Siteler:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   162
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   120
               TabIndex        =   112
               Top             =   1440
               Width           =   1920
            End
         End
      End
      Begin VB.Image Image2 
         Height          =   4350
         Left            =   -69880
         Picture         =   "Form1.frx":8D626
         Top             =   360
         Visible         =   0   'False
         Width           =   9870
      End
   End
   Begin XtremeSuiteControls.CommonDialog Cmd1 
      Left            =   12240
      Top             =   6960
      _Version        =   851968
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   4
   End
   Begin VB.Label Label14 
      Caption         =   "asd"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   15240
      TabIndex        =   48
      Top             =   8400
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   7500
      Left            =   0
      Picture         =   "Form1.frx":A42E1
      Top             =   0
      Width           =   11700
   End
   Begin VB.Menu siteaç 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu SiteyiAc 
         Caption         =   "&Siteyi Aç"
      End
   End
   Begin VB.Menu siteaç2 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu SiteyiAc2 
         Caption         =   "&Siteyi Aç"
      End
   End
   Begin VB.Menu siteaç3 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu SiteyiAc3 
         Caption         =   "&Exploit'i aç"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Command1_Click()
On Error Resume Next
Dim ArananKelime As String
Dim KelimeninYeri, AramayaBasla As Integer
Kill App.Path & "\ReserveIP.txt"
ListBox1.Clear
ListBox2.Clear
If InStr(1, Text1, "http://") Then
On Error GoTo hata
ArananKelime = "http://" 'text2 içindeki kelimeyi arayacaðýz
AramayaBasla = Text1.SelStart + Text1.SelLength 'arama yapýlacak metin uzunluðunda arama yapacaðýz
If AramayaBasla = 0 Or AramayaBasla = Len(Text1.Text) Then AramayaBasla = 1 'aranan kelime bulunmazsa baþa döneceðiz
KelimeninYeri = InStr(AramayaBasla, Text1.Text, ArananKelime, vbTextCompare)
Text1.SetFocus 'kelime bulunduðunda iþaretliyoruz
Text1.SelStart = KelimeninYeri - 1
Text1.SelLength = Len(ArananKelime)
Text1.SelText = ""
Command1_Click
Command1_Click
Exit Sub
ElseIf InStr(1, Text1, "/") Then
ArananKelime = "/" 'text2 içindeki kelimeyi arayacaðýz
AramayaBasla = Text1.SelStart + Text1.SelLength 'arama yapýlacak metin uzunluðunda arama yapacaðýz
If AramayaBasla = 0 Or AramayaBasla = Len(Text1.Text) Then AramayaBasla = 1 'aranan kelime bulunmazsa baþa döneceðiz
KelimeninYeri = InStr(AramayaBasla, Text1.Text, ArananKelime, vbTextCompare)
Text1.SetFocus 'kelime bulunduðunda iþaretliyoruz
Text1.SelStart = KelimeninYeri - 1
Text1.SelLength = Len(ArananKelime)
Text1.SelText = ""
Exit Sub
hata: 'arama metni sonuna geldiðimizde baþtan bir daha baþlýyoruz
Text1.SelStart = 1
End If
Web1.navigate ("http://kimindir.com/" & Text1.Text)
Web2.navigate ("http://www.websiteneighbors.com/results.php?output=php&ip_host=" & Text1.Text)
Web4.navigate (Text1.Text)
web5.navigate (Text1.Text)
FlatEdit1.Text = ""
FlatEdit2.Text = ""
FlatEdit3.Text = ""
FlatEdit4.Text = ""
FlatEdit5.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text49.Text = ""
'Text50.Text = ""
'List1.Clear
Label1.Caption = "Sunucudaki Toplam Siteler:"
Label13.Caption = "TTL="
Komut ("ping " & Text1.Text)
Timer3.Enabled = False
End Sub

Private Sub Command2_Click()
Text2.Text = Web1.document.body.innerText
End Sub





Private Sub Command3_Click()
ListBox1.Clear
WebBrowser1.navigate "http://www.exploit-db.com/search/?action=search&filter_page=1&filter_description=" & FlatEdit5.Text
End Sub

Private Sub FlatEdit3_Change()
Web3.navigate "https://w3dt.net/tools/portscan/?host=" & FlatEdit3.Text & "&t=gnrl&submit=Scan%21&clean_opt=1"
End Sub






Private Sub Form_Load()
On Error Resume Next
Kill App.Path & "\ReserveIP.txt"
Kill App.Path & "\URLRoot.txt"

ListView1.ColumnHeaders.Add 1, , "Host / IP"
ListView1.ColumnHeaders.Add 2, , "Port No"
ListView1.ColumnHeaders.Add 3, , "Durum"
ListView1.ColumnHeaders.Add 4, , "Hizmet Adý"
ListView1.ColumnHeaders.Add 5, , "Port Hakkýnda Ek Bilgi"
' Þimdi de Kolonlarý Biçimlendiriyoruz.
ListView1.ColumnHeaders(1).Width = 1800
ListView1.ColumnHeaders(2).Width = 1200
ListView1.ColumnHeaders(3).Width = 1200
ListView1.ColumnHeaders(4).Width = 1200
ListView1.ColumnHeaders(5).Width = 3000

Cmd1.DialogTitle = "Kaydetmek Ýstediðiniz Dizini Seçiniz"
End Sub

Private Sub Form_Resize()
On Error Resume Next
Form1.Height = 7920
Form1.Width = 11805
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Kill App.Path & "\ReserveIP.txt"
Kill App.Path & "\URLRoot.txt"
End
End Sub

Private Sub Label13_Change()
Dim Sayi As Integer
Sayi = Label14.Caption
Select Case Sayi
Case Is < 60
FlatEdit2.Text = "Windows 95 / 98 / 98SE / Me / NT 4"
Case Is > 60
FlatEdit2.Text = "Linux Kernel 2.2.x+"
Case Is > 120
FlatEdit2.Text = "Windows 2000 ve üzeri"
Case Is > 245
FlatEdit2.Text = "BSD,BSDI,Solaris,AIX vb Unix türevleri"
End Select
If Label4.Caption = "asd" Then
FlatEdit2.Text = "Belirsiz"
End If
End Sub

Private Sub ListBox1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then PopupMenu siteaç3
End Sub

Private Sub ListBox2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then PopupMenu siteaç2
End Sub

Private Sub PushButton1_Click()
On Error Resume Next
Dim KayýtYolu, Kayýt As String
kaydetme = MsgBox("Toplanan Bilgilerin Hepsini Kaydetmek mi Ýstiyorsunuz", vbYesNo + vbInformation, "Bildiri;")
If kaydetme = vbYes Then
Cmd1.ShowBrowseFolder
KayýtYolu = Cmd1.FileName
Kayýt = KayýtYolu & "\" & Text1.Text & " - Sonuçlarý"
MkDir Kayýt
'Whois [Sunucu Bilgileri]
Open Kayýt & "\Whois [Sunucu Bilgileri].txt" For Output As #1
Print #1, "WHOIS [Sunucu Bilgileri]" & vbCrLf
Print #1, "Sorgulanan Alan Adý : " & Text3.Text; vbCrLf
Print #1, "Oluþturma Tarihi : " & Text4.Text; vbCrLf
Print #1, "Bitiþ Tarihi : " & Text5.Text; vbCrLf
Print #1, "Son Güncelleme Tarihi : " & Text6.Text; vbCrLf
Print #1, "Durumu : " & Text7.Text; vbCrLf
Print #1, "Kimindir Sunucusu : " & Text8.Text; vbCrLf
Print #1, "Ýsim Sunucularý : " & Text9.Text; vbCrLf
Print #1, "Whois Sonucu : "
Print #1, Text10.Text; vbCrLf
Print #1, "Danger: Hedef Site Bilgi Toplama | _DangerOusMaN_"
Close #1
'Reserve IP
Open Kayýt & "\ReserveIP [Komþu IP Adresleri].txt" For Output As #1
Print #1, "RESERVE-IP [Komþu IP Adresleri]" & vbCrLf
Print #1, Label1.Caption; vbCrLf
Print #1, Text30.Text; vbCrLf
Print #1, "Danger: Hedef Site Bilgi Toplama | _DangerOusMaN_"
Close #1
'Ping Taramasý
Open Kayýt & "\Ping Sonuclarý.txt" For Output As #1
Print #1, "PING Sonuclarý" & vbCrLf
Print #1, FlatEdit1.Text; vbCrLf
Print #1, "Site IP Adresi : " & FlatEdit3.Text & Space(10) & Label13.Caption; vbCrLf
Print #1, Label10.Caption & " " & FlatEdit2.Text; vbCrLf
Print #1, "Danger: Hedef Site Bilgi Toplama | _DangerOusMaN_"
Close #1
'Port Taramasý
Open Kayýt & "\Port Tarama Sonuclarý.txt" For Output As #1
Print #1, "PORT TARAMA SONUCLARI" & vbCrLf
Print #1, "Danger: Hedef Site Bilgi Toplama | _DangerOusMaN_"
Close #1
'
Open Kayýt & "\Hedef Site Web Analizi.txt" For Output As #1
Print #1, "HEDEF SÝTE WEB ANALÝZÝ" & vbCrLf
Print #1, "Yazýlým Dili : " & FlatEdit4.Text
Print #1, "Script Adý : " & FlatEdit5.Text; vbCrLf
Print #1, "Bulunan Exploit Adresleri"
Print #1, Text57.Text; vbCrLf
Print #1, "Danger: Hedef Site Bilgi Toplama | _DangerOusMaN_"
Close #1
'Sitenin URL ADRESLERÝ
Open Kayýt & "\Sitenin URL Adresleri.txt" For Output As #1
Print #1, "SÝTENÝN URL ADRESLERÝ" & vbCrLf
Print #1, Text60.Text; vbCrLf
Print #1, "Danger: Hedef Site Bilgi Toplama | _DangerOusMaN_"
Close #1
End If
End Sub

Private Sub PushButton2_Click()
cýkýs = MsgBox("Programdan Çýkmak mý Ýstiyorsunuz ?", vbInformation, "Çýkýþ Yapýlýyor...")
If cýkýs = vbYes Then
End
End If
End Sub
Private Sub List1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then PopupMenu siteaç
End Sub

Private Sub SiteyiAc3_Click()
On Error Resume Next
Dim sayfa
sayfa = ListBox1.List(ListBox1.ListIndex)
ShellExecute GetActiveWindow(), "Open", sayfa, "", 0&, 1
End Sub
Private Sub SiteyiAc2_Click()
On Error Resume Next
Dim sayfa
sayfa = ListBox2.List(ListBox2.ListIndex)
ShellExecute GetActiveWindow(), "Open", sayfa, "", 0&, 1
End Sub
Private Sub SiteyiAc_Click()
On Error Resume Next
Dim sayfa
sayfa = List1.List(List1.ListIndex)
ShellExecute GetActiveWindow(), "Open", sayfa, "", 0&, 1
End Sub

Private Sub Text2_Change()
SorgulananAlanAdý
OlusturmaTarihi
BitisTarihi
SonGuncellemeTarihi
Durumu
ÝsimSunucularý
KimindirSunucusu
Whois
If InStr(1, Text2.Text, "alan adý kimse tarafýndan alýnmamýþ.") Then
'MsgBox "[ " & Text1.Text & " ] alan adý kimse tarafýndan alýnmamýþ.", vbExclamation, "Uyarý;"
End If
Timer3.Enabled = True
End Sub

Private Function Whois()
On Error Resume Next
Dim AranacakYer, aranan As String
Dim ilk, son
AranacakYer = Text2.Text 'Verinin Aranacaðý Yer
deg1 = deg1 & AranacakYer
aranan = Text25.Text  '1. Deger
ilk = InStr(1, deg1, aranan) + Len(aranan)
If ilk <> Len(aranan) Then
son = InStr(ilk, deg1, Text26.Text) - ilk  ' " "  Arasýndaki 2. Deger
Text10.Text = Mid(deg1, ilk, son) 'Degerlerin arasýndaki veri
Exit Function
End If
End Function

Private Function KimindirSunucusu()
On Error Resume Next
Dim AranacakYer, aranan As String
Dim ilk, son
AranacakYer = Text2.Text 'Verinin Aranacaðý Yer
deg1 = deg1 & AranacakYer
aranan = Text23.Text  '1. Deger
ilk = InStr(1, deg1, aranan) + Len(aranan)
If ilk <> Len(aranan) Then
son = InStr(ilk, deg1, Text24.Text) - ilk  ' " "  Arasýndaki 2. Deger
Text9.Text = Mid(deg1, ilk, son) 'Degerlerin arasýndaki veri
Exit Function
End If
End Function

Private Function ÝsimSunucularý()
On Error Resume Next
Dim AranacakYer, aranan As String
Dim ilk, son
AranacakYer = Text2.Text 'Verinin Aranacaðý Yer
deg1 = deg1 & AranacakYer
aranan = Text21.Text  '1. Deger
ilk = InStr(1, deg1, aranan) + Len(aranan)
If ilk <> Len(aranan) Then
son = InStr(ilk, deg1, Text22.Text) - ilk  ' " "  Arasýndaki 2. Deger
Text8.Text = Mid(deg1, ilk, son) 'Degerlerin arasýndaki veri
Exit Function
End If
End Function

Private Function Durumu()
On Error Resume Next
Dim AranacakYer, aranan As String
Dim ilk, son
AranacakYer = Text2.Text 'Verinin Aranacaðý Yer
deg1 = deg1 & AranacakYer
aranan = Text19.Text  '1. Deger
ilk = InStr(1, deg1, aranan) + Len(aranan)
If ilk <> Len(aranan) Then
son = InStr(ilk, deg1, Text20.Text) - ilk  ' " "  Arasýndaki 2. Deger
Text7.Text = Mid(deg1, ilk, son) 'Degerlerin arasýndaki veri
Exit Function
End If
End Function

Private Function SonGuncellemeTarihi()
On Error Resume Next
Dim AranacakYer, aranan As String
Dim ilk, son
AranacakYer = Text2.Text 'Verinin Aranacaðý Yer
deg1 = deg1 & AranacakYer
aranan = Text17.Text  '1. Deger
ilk = InStr(1, deg1, aranan) + Len(aranan)
If ilk <> Len(aranan) Then
son = InStr(ilk, deg1, Text18.Text) - ilk  ' " "  Arasýndaki 2. Deger
Text6.Text = Mid(deg1, ilk, son) 'Degerlerin arasýndaki veri
Exit Function
End If
End Function

Private Function BitisTarihi()
On Error Resume Next
Dim AranacakYer, aranan As String
Dim ilk, son
AranacakYer = Text2.Text 'Verinin Aranacaðý Yer
deg1 = deg1 & AranacakYer
aranan = Text15.Text  '1. Deger
ilk = InStr(1, deg1, aranan) + Len(aranan)
If ilk <> Len(aranan) Then
son = InStr(ilk, deg1, Text16.Text) - ilk  ' " "  Arasýndaki 2. Deger
Text5.Text = Mid(deg1, ilk, son) 'Degerlerin arasýndaki veri
Exit Function
End If
End Function

Private Function OlusturmaTarihi()
On Error Resume Next
Dim AranacakYer, aranan As String
Dim ilk, son
AranacakYer = Text2.Text 'Verinin Aranacaðý Yer
deg1 = deg1 & AranacakYer
aranan = Text13.Text  '1. Deger
ilk = InStr(1, deg1, aranan) + Len(aranan)
If ilk <> Len(aranan) Then
son = InStr(ilk, deg1, Text14.Text) - ilk  ' " "  Arasýndaki 2. Deger
Text4.Text = Mid(deg1, ilk, son) 'Degerlerin arasýndaki veri
Exit Function
End If
End Function

Private Function SorgulananAlanAdý()
On Error Resume Next
Dim AranacakYer, aranan As String
Dim ilk, son
AranacakYer = Text2.Text 'Verinin Aranacaðý Yer
deg1 = deg1 & AranacakYer
aranan = Text11.Text  '1. Deger
ilk = InStr(1, deg1, aranan) + Len(aranan)
If ilk <> Len(aranan) Then
son = InStr(ilk, deg1, Text12.Text) - ilk  ' " "  Arasýndaki 2. Deger
Text3.Text = Mid(deg1, ilk, son) 'Degerlerin arasýndaki veri
Exit Function
End If
End Function


Private Sub Text27_Change()
ReserveIP
End Sub
Private Function ReserveIP()
On Error Resume Next
Dim AranacakYer, aranan As String
Dim ilk, son
AranacakYer = Text27.Text 'Verinin Aranacaðý Yer
deg1 = deg1 & AranacakYer
aranan = Text28.Text  '1. Deger
ilk = InStr(1, deg1, aranan) + Len(aranan)
If ilk <> Len(aranan) Then
son = InStr(ilk, deg1, Text29.Text) - ilk  ' " "  Arasýndaki 2. Deger
Text30.Text = Mid(deg1, ilk, son) 'Degerlerin arasýndaki veri
Exit Function
End If
End Function

Private Sub Text30_Change()
Timer1.Enabled = True
'ReserveIPLinkleriKaydet
End Sub

Private Function ReserveIPoku()
List1.Clear
'* Ýf Döngüsü Baþlangýç
If Dir(App.Path & "\ReserveIP.txt") <> "" Then
If App.Path & "\ReserveIP.txt" = "" Then
'* Dosya Seçilmezse Bilgi Mesajý verir
MsgBox "Dosya Yok", vbInformation, "Uyarý ;"
'* Taramayý Baþlat Butonunu Aktifleþtirir
'* Deðilse
Else
'* Seçilen Dosyayý List1'e aktarýr
Open App.Path & "\ReserveIP.txt" For Input As #1
While Not EOF(1)
Input #1, a
List1.AddItem a
Wend
Close #1
'* Label'e Listbox'daki Toplam Deðeri Yansýtýr
Label1.Caption = "Sunucudaki Toplam Siteler: " & List1.ListCount
'* Taramayý Baþlat Butonunu Aktifleþtirir
End If
'* Ýf Döngüsü Bitiþ
End If
End Function

Private Function ReserveIPLinkleriKaydet()
Open App.Path & "\ReserveIP.txt" For Output As #1
Print #1, Text30.Text
Close #1
If Dir(App.Path & "\ReserveIP.txt") <> "" Then
ReserveIPoku
End If
End Function
Private Function URLRooToku()
ListBox2.Clear
'* Ýf Döngüsü Baþlangýç
If Dir(App.Path & "\URLRoot.txt") <> "" Then
If App.Path & "\URLRoot.txt" = "" Then
'* Dosya Seçilmezse Bilgi Mesajý verir
MsgBox "Dosya Yok", vbInformation, "Uyarý ;"
'* Taramayý Baþlat Butonunu Aktifleþtirir
'* Deðilse
Else
'* Seçilen Dosyayý List1'e aktarýr
Open App.Path & "\URLRoot.txt" For Input As #1
While Not EOF(1)
Input #1, a
ListBox2.AddItem a
Wend
Close #1
'* Label'e Listbox'daki Toplam Deðeri Yansýtýr
Label1.Caption = "Sunucudaki Toplam Siteler: " & ListBox2.ListCount
'* Taramayý Baþlat Butonunu Aktifleþtirir
End If
'* Ýf Döngüsü Bitiþ
End If
End Function
Private Function URLRooTKaydet()
Open App.Path & "\URLRoot.txt" For Output As #1
Print #1, Text60.Text
Close #1
If Dir(App.Path & "\URLRoot.txt") <> "" Then
URLRooToku
End If
End Function


Private Sub Text33_Change()
HostIP
PortNo
Durumm
HizmetÝsmi
DetaylýBilgi
End Sub

Private Function DetaylýBilgi()
On Error Resume Next
Dim AranacakYer, aranan As String
Dim ilk, son
AranacakYer = Text33.Text 'Verinin Aranacaðý Yer
deg1 = deg1 & AranacakYer
aranan = Text47.Text  '1. Deger
ilk = InStr(1, deg1, aranan) + Len(aranan)
If ilk <> Len(aranan) Then
son = InStr(ilk, deg1, Text48.Text) - ilk  ' " "  Arasýndaki 2. Deger
Text38.Text = Mid(deg1, ilk, son) 'Degerlerin arasýndaki veri
Exit Function
End If
End Function

Private Function HizmetÝsmi()
On Error Resume Next
Dim AranacakYer, aranan As String
Dim ilk, son
AranacakYer = Text33.Text 'Verinin Aranacaðý Yer
deg1 = deg1 & AranacakYer
aranan = Text45.Text  '1. Deger
ilk = InStr(1, deg1, aranan) + Len(aranan)
If ilk <> Len(aranan) Then
son = InStr(ilk, deg1, Text46.Text) - ilk  ' " "  Arasýndaki 2. Deger
Text37.Text = Mid(deg1, ilk, son) 'Degerlerin arasýndaki veri
Exit Function
End If
End Function


Private Function Durumm()
On Error Resume Next
Dim AranacakYer, aranan As String
Dim ilk, son
AranacakYer = Text33.Text 'Verinin Aranacaðý Yer
deg1 = deg1 & AranacakYer
aranan = Text43.Text  '1. Deger
ilk = InStr(1, deg1, aranan) + Len(aranan)
If ilk <> Len(aranan) Then
son = InStr(ilk, deg1, Text44.Text) - ilk  ' " "  Arasýndaki 2. Deger
Text36.Text = Mid(deg1, ilk, son) 'Degerlerin arasýndaki veri
Exit Function
End If
End Function


Private Function PortNo()
On Error Resume Next
Dim AranacakYer, aranan As String
Dim ilk, son
AranacakYer = Text33.Text 'Verinin Aranacaðý Yer
deg1 = deg1 & AranacakYer
aranan = Text41.Text  '1. Deger
ilk = InStr(1, deg1, aranan) + Len(aranan)
If ilk <> Len(aranan) Then
son = InStr(ilk, deg1, Text42.Text) - ilk  ' " "  Arasýndaki 2. Deger
Text35.Text = Mid(deg1, ilk, son) 'Degerlerin arasýndaki veri
Exit Function
End If
End Function

Private Function HostIP()
On Error Resume Next
Dim AranacakYer, aranan As String
Dim ilk, son
AranacakYer = Text33.Text 'Verinin Aranacaðý Yer
deg1 = deg1 & AranacakYer
aranan = Text39.Text  '1. Deger
ilk = InStr(1, deg1, aranan) + Len(aranan)
If ilk <> Len(aranan) Then
son = InStr(ilk, deg1, Text40.Text) - ilk  ' " "  Arasýndaki 2. Deger
Text34.Text = Mid(deg1, ilk, son) 'Degerlerin arasýndaki veri
Exit Function
End If
End Function

Private Sub Text49_Change()
On Error Resume Next
If InStr(1, Text49, ".asp") Then
FlatEdit4.Text = ".asp"
ElseIf InStr(1, Text49, ".php") Then
FlatEdit4.Text = ".php"
ElseIf InStr(1, Text49, ".html") Then
FlatEdit4.Text = ".html"
ElseIf InStr(1, Text49, ".htm") Then
FlatEdit4.Text = ".html"
End If
If InStr(1, Text49, "forum") Then
Text50.Text = Inet1.OpenURL(Text1.Text & "/forum" & FlatEdit4.Text)
Else
Text50.Text = Inet1.OpenURL(Text1.Text)
End If
If InStr(1, Text49, "index") Then
Text50.Text = Inet1.OpenURL(Text1.Text & "/index" & FlatEdit4.Text)
Else
Text50.Text = Inet1.OpenURL(Text1.Text)
End If
If InStr(1, Text49, "/forum/") Then
Text50.Text = Inet1.OpenURL(Text1.Text & "/forum/")
End If
End Sub

Private Function SMFBulucu()
On Error Resume Next
Dim AranacakYer, aranan As String
Dim ilk, son
AranacakYer = Text50.Text 'Verinin Aranacaðý Yer
deg1 = deg1 & AranacakYer
aranan = Text55.Text  '1. Deger
ilk = InStr(1, deg1, aranan) + Len(aranan)
If ilk <> Len(aranan) Then
son = InStr(ilk, deg1, Text56.Text) - ilk  ' " "  Arasýndaki 2. Deger
FlatEdit5.Text = "SMF" & Mid(deg1, ilk, son) 'Degerlerin arasýndaki veri
Exit Function
End If
End Function
Private Function JoomlaBulucu()
On Error Resume Next
Dim AranacakYer, aranan As String
Dim ilk, son
AranacakYer = Text50.Text 'Verinin Aranacaðý Yer
deg1 = deg1 & AranacakYer
aranan = Text53.Text  '1. Deger
ilk = InStr(1, deg1, aranan) + Len(aranan)
If ilk <> Len(aranan) Then
son = InStr(ilk, deg1, Text54.Text) - ilk  ' " "  Arasýndaki 2. Deger
FlatEdit5.Text = "Joomla" & Mid(deg1, ilk, son) 'Degerlerin arasýndaki veri
Exit Function
End If
End Function

Private Function vBulletinBulucu()
On Error Resume Next
Dim AranacakYer, aranan As String
Dim ilk, son
AranacakYer = Text50.Text 'Verinin Aranacaðý Yer
deg1 = deg1 & AranacakYer
aranan = Text51.Text  '1. Deger
ilk = InStr(1, deg1, aranan) + Len(aranan)
If ilk <> Len(aranan) Then
son = InStr(ilk, deg1, Text52.Text) - ilk  ' " "  Arasýndaki 2. Deger
FlatEdit5.Text = "vBulletin" & Mid(deg1, ilk, son) 'Degerlerin arasýndaki veri
Exit Function
End If
End Function




Private Sub Text50_Change()
Timer4.Enabled = True
End Sub

Private Sub Text57_Change()
Timer5.Enabled = True
End Sub



Private Sub Timer1_Timer()
Dim ArananKelime As String
Dim KelimeninYeri, AramayaBasla As Integer
On Error GoTo hata

ArananKelime = Text31 'text2 içindeki kelimeyi arayacaðýz
AramayaBasla = Text30.SelStart + Text30.SelLength 'arama yapýlacak metin uzunluðunda arama yapacaðýz
If AramayaBasla = 0 Or AramayaBasla = Len(Text30.Text) Then AramayaBasla = 1 'aranan kelime bulunmazsa baþa döneceðiz
KelimeninYeri = InStr(AramayaBasla, Text30.Text, ArananKelime, vbTextCompare)
Text30.SetFocus 'kelime bulunduðunda iþaretliyoruz
Text30.SelStart = KelimeninYeri - 1
Text30.SelLength = Len(ArananKelime)
Text30.SelText = ""
Exit Sub
hata: 'arama metni sonuna geldiðimizde baþtan bir daha baþlýyoruz
Text30.SelStart = 1
If Text30.SelStart = 1 Then
Timer1.Enabled = False
ReserveIPLinkleriKaydet
End If
End Sub

Private Function TTLtahmini()
On Error Resume Next
Dim AranacakYer, aranan As String
Dim ilk, son
AranacakYer = FlatEdit1.Text 'Verinin Aranacaðý Yer
deg1 = deg1 & AranacakYer
aranan = "TTL="  '1. Deger
ilk = InStr(1, deg1, aranan) + Len(aranan)
If ilk <> Len(aranan) Then
son = InStr(ilk, deg1, FlatEdit3.Text) - ilk
Label14.Caption = Mid(deg1, ilk, son)    ' " "  Arasýndaki 2. Deger
Label13.Caption = "TTL=" & Mid(deg1, ilk, son)

Exit Function
End If
End Function


Private Function SiteIPAdress()
On Error Resume Next
Dim AranacakYer, aranan As String
Dim ilk, son
AranacakYer = FlatEdit1.Text 'Verinin Aranacaðý Yer
deg1 = deg1 & AranacakYer
aranan = "["  '1. Deger
ilk = InStr(1, deg1, aranan) + Len(aranan)
If ilk <> Len(aranan) Then
son = InStr(ilk, deg1, "]") - ilk   ' " "  Arasýndaki 2. Deger
FlatEdit3.Text = Mid(deg1, ilk, son) 'Degerlerin arasýndaki veri
Exit Function
End If
End Function

Private Sub Timer2_Timer()
    Dim lngBytesread As Long
    Dim strBuff As String * 2048
    If ReadFile(hReadPipe, strBuff, 2048, lngBytesread, 0&) <> 0 Then
   FlatEdit1.Text = FlatEdit1.Text & Left(strBuff, lngBytesread)
    Else
    CloseHandle (proc.hProcess)
    CloseHandle (proc.hThread)
    CloseHandle (hReadPipe)
    Timer2.Enabled = False
    SiteIPAdress
    TTLtahmini
End If
End Sub

Private Sub Timer3_Timer()
Dim Bulunamýyor
Bulunamýyor = ""
Select Case Bulunamýyor
Case Is = Text3.Text
Text3.Text = "Bulunamýyor"
Case Is = Text4.Text
Text4.Text = "Bulunamýyor"
Case Is = Text5.Text
Text5.Text = "Bulunamýyor"
Case Is = Text6.Text
Text6.Text = "Bulunamýyor"
Case Is = Text7.Text
Text7.Text = "Bulunamýyor"
Case Is = Text8.Text
Text8.Text = "Bulunamýyor"
Case Is = Text9.Text
Text9.Text = "Bulunamýyor"
Case Is = Text10.Text
Text10.Text = "Bulunamýyor"
End Select
'Timer3.Enabled = False
End Sub

Private Sub Timer4_Timer()
On Error Resume Next
If InStr(1, Text50, "SMF") Then
ListBox1.Clear
SMFBulucu
WebBrowser1.navigate "http://www.exploit-db.com/search/?action=search&filter_page=1&filter_description=" & FlatEdit5.Text
Label18.Caption = "Bulunan Exploit : " & ListBox1.ListCount
Timer4.Enabled = False
ElseIf InStr(1, Text50, "Joomla") Then
ListBox1.Clear
JoomlaBulucu
WebBrowser1.navigate "http://www.exploit-db.com/search/?action=search&filter_page=1&filter_description=" & FlatEdit5.Text
Label18.Caption = "Bulunan Exploit : " & ListBox1.ListCount
Timer4.Enabled = False
ElseIf InStr(1, Text50, "WordPress") Then
ListBox1.Clear
FlatEdit5.Text = "WordPress"
WebBrowser1.navigate "http://www.exploit-db.com/search/?action=search&filter_page=1&filter_description=" & FlatEdit5.Text
Label18.Caption = "Bulunan Exploit : " & ListBox1.ListCount
Timer4.Enabled = False
ElseIf InStr(1, Text50, "vBulletin") Then
ListBox1.Clear
vBulletinBulucu
WebBrowser1.navigate "http://www.exploit-db.com/search/?action=search&filter_page=1&filter_description=" & FlatEdit5.Text
Label18.Caption = "Bulunan Exploit : " & ListBox1.ListCount
Timer4.Enabled = False
End If
If FlatEdit5.Text = "" Then
FlatEdit5.Text = "Bulunamýyor."
Timer4.Enabled = False
End If
'Timer4.Enabled = False
End Sub

Private Function ExploitLink()
On Error Resume Next
Dim AranacakYer, aranan As String
Dim ilk, son
AranacakYer = Text57.Text 'Verinin Aranacaðý Yer
deg1 = deg1 & AranacakYer
aranan = Text58.Text  '1. Deger
ilk = InStr(1, deg1, aranan) + Len(aranan)
If ilk <> Len(aranan) Then
son = InStr(ilk, deg1, Text59.Text) - ilk   ' " "  Arasýndaki 2. Deger
ListBox1.AddItem Text58.Text & Mid(deg1, ilk, son)  'Degerlerin arasýndaki veri
Text57.SelText = ""
Exit Function
End If
End Function
Private Function Seçici()
On Error Resume Next
Dim ArananKelime As String
Dim KelimeninYeri, AramayaBasla As Integer
ArananKelime = "http://www.exploit-db.com/exploits/" 'text2 içindeki kelimeyi arayacaðýz
AramayaBasla = Text57.SelStart + Text57.SelLength 'arama yapýlacak metin uzunluðunda arama yapacaðýz
If AramayaBasla = 0 Or AramayaBasla = Len(Text57.Text) Then AramayaBasla = 1 'aranan kelime bulunmazsa baþa döneceðiz
KelimeninYeri = InStr(AramayaBasla, Text57.Text, ArananKelime, vbTextCompare)
Text57.SetFocus 'kelime bulunduðunda iþaretliyoruz
Text57.SelStart = KelimeninYeri - 1
Text57.SelLength = Len(ArananKelime)
Exit Function
hata: 'arama metni sonuna geldiðimizde baþtan bir daha baþlýyoruz
Text57.SelStart = 1
End Function

Private Sub Timer5_Timer()
Seçici
ExploitLink
Label18.Caption = "Bulunan Exploit : " & ListBox1.ListCount
If Label18.Caption = "Bulunan Exploit : 10" Then
Timer5.Enabled = False
End If
End Sub

Private Sub Web1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
Text2.Text = Web1.document.body.innerText
End Sub

Private Sub Web2_DocumentComplete(ByVal pDisp As Object, URL As Variant)
Text27.Text = Web2.document.body.innerText
End Sub

Private Sub Web3_DocumentComplete(ByVal pDisp As Object, URL As Variant)
Text33.Text = Web3.document.body.innerHTML
End Sub

Private Sub Web4_DocumentComplete(ByVal pDisp As Object, URL As Variant)
'* Kodu KullanaBilmek için veya Hata Alýrsanýz
'* Project>References>Microsoft HTML Object Library Dll'sini Seçiniz
Dim HTMLdoc As HTMLDocument
Dim HTMLlinks As HTMLAnchorElement
Dim HTMLlink As HTMLAnchorElement
Dim STRtxt As String
'Dim STRtxt1 As String
On Error Resume Next
Set HTMLdoc = Web4.document
For Each HTMLlinks In HTMLdoc.links
STRtxt = STRtxt & HTMLlinks.href & vbCrLf
Next HTMLlinks
'* WebBrowser2'de Buldugu Tüm Linkleri Text2 Aktarýr
Text49.Text = STRtxt

End Sub
Private Sub Web5_DocumentComplete(ByVal pDisp As Object, URL As Variant)
'* Kodu KullanaBilmek için veya Hata Alýrsanýz
'* Project>References>Microsoft HTML Object Library Dll'sini Seçiniz
Dim HTMLdoc As HTMLDocument
Dim HTMLlinks As HTMLAnchorElement
Dim HTMLlink As HTMLAnchorElement
Dim STRtxt As String
'Dim STRtxt1 As String
On Error Resume Next
Set HTMLdoc = web5.document
For Each HTMLlinks In HTMLdoc.links
STRtxt = STRtxt & HTMLlinks.href & vbCrLf
Next HTMLlinks
'* WebBrowser2'de Buldugu Tüm Linkleri Text2 Aktarýr
Text60.Text = STRtxt
URLRooTKaydet
End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
'* Kodu KullanaBilmek için veya Hata Alýrsanýz
'* Project>References>Microsoft HTML Object Library Dll'sini Seçiniz
Dim HTMLdoc As HTMLDocument
Dim HTMLlinks As HTMLAnchorElement
Dim HTMLlink As HTMLAnchorElement
Dim STRtxt As String
'Dim STRtxt1 As String
On Error Resume Next
Set HTMLdoc = WebBrowser1.document
For Each HTMLlinks In HTMLdoc.links
STRtxt = STRtxt & HTMLlinks.href & vbCrLf
Next HTMLlinks
'* WebBrowser2'de Buldugu Tüm Linkleri Text2 Aktarýr
Text57.Text = STRtxt
End Sub
