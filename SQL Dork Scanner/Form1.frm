VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "ActiveX.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Danger SQL Ýnjection Dorks Scanner | _DangerOusMaN_ "
   ClientHeight    =   7455
   ClientLeft      =   2055
   ClientTop       =   2865
   ClientWidth     =   11715
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   12555
   ScaleWidth      =   17160
   Begin VB.TextBox Text 
      Height          =   285
      Left            =   7200
      TabIndex        =   41
      Top             =   7920
      Width           =   735
   End
   Begin VB.TextBox Text13 
      Height          =   285
      Left            =   120
      TabIndex        =   40
      Text            =   "Text13"
      Top             =   8400
      Width           =   3735
   End
   Begin XtremeSuiteControls.PushButton PushButton2 
      Height          =   375
      Left            =   8775
      TabIndex        =   39
      Top             =   5920
      Width           =   2655
      _Version        =   851968
      _ExtentX        =   4683
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Programý Kapat"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   1
   End
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   375
      Left            =   4560
      TabIndex        =   38
      Top             =   5920
      Width           =   2655
      _Version        =   851968
      _ExtentX        =   4683
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Bulunan Açýklý Siteleri Kaydet"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   1
   End
   Begin XtremeSuiteControls.PushButton Command5 
      Height          =   375
      Left            =   4440
      TabIndex        =   37
      Top             =   2300
      Width           =   2295
      _Version        =   851968
      _ExtentX        =   4048
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Bulunan Linkleri Kaydet"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      Appearance      =   1
   End
   Begin XtremeSuiteControls.ComboBox ComboBox1 
      Height          =   315
      Left            =   240
      TabIndex        =   36
      Top             =   5880
      Width           =   3855
      _Version        =   851968
      _ExtentX        =   6800
      _ExtentY        =   556
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
      Appearance      =   1
      UseVisualStyle  =   0   'False
      Text            =   "Ülke Uzantýlarý"
   End
   Begin XtremeSuiteControls.ListBox List3 
      Height          =   2600
      Left            =   4560
      TabIndex        =   32
      Top             =   3120
      Width           =   6900
      _Version        =   851968
      _ExtentX        =   12171
      _ExtentY        =   4586
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
      Appearance      =   1
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton Command4 
      Height          =   375
      Left            =   6675
      TabIndex        =   31
      Top             =   1365
      Width           =   1935
      _Version        =   851968
      _ExtentX        =   3413
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Taramayý Durdur"
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      Appearance      =   1
   End
   Begin XtremeSuiteControls.PushButton Command3 
      Height          =   375
      Left            =   4545
      TabIndex        =   30
      Top             =   1365
      Width           =   1935
      _Version        =   851968
      _ExtentX        =   3413
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Taramayý Baþlat"
      ForeColor       =   65280
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      Appearance      =   1
   End
   Begin XtremeSuiteControls.PushButton Command2 
      Height          =   255
      Left            =   2355
      TabIndex        =   29
      Top             =   5235
      Width           =   1815
      _Version        =   851968
      _ExtentX        =   3201
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Manuel Dork Eke"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   1
   End
   Begin XtremeSuiteControls.PushButton Command1 
      Height          =   255
      Left            =   260
      TabIndex        =   28
      Top             =   5235
      Width           =   1815
      _Version        =   851968
      _ExtentX        =   3201
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Dork List Ekle"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   1
   End
   Begin XtremeSuiteControls.PushButton Command6 
      Height          =   255
      Left            =   2700
      TabIndex        =   27
      Top             =   4920
      Width           =   1455
      _Version        =   851968
      _ExtentX        =   2566
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Listeyi Temizle"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   1
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3150
      Left            =   240
      TabIndex        =   25
      Top             =   1680
      Width           =   3920
   End
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   2160
      Top             =   7800
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1680
      Top             =   7800
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   3120
      Top             =   7800
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   2500
      Left            =   2640
      Top             =   7800
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1200
      Top             =   7800
   End
   Begin VB.Frame Frame3 
      Caption         =   "-Link Ayýrýcý / SQL Ýnjection Bulucu"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   8160
      TabIndex        =   6
      Top             =   7560
      Width           =   8175
      Begin VB.TextBox Text12 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "'a"
         Top             =   480
         Width           =   375
      End
      Begin VB.Frame Frame4 
         Caption         =   "Aranacak Hatalar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   4080
         TabIndex        =   18
         Top             =   3360
         Width           =   3975
         Begin VB.TextBox Text11 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   23
            Text            =   "Microsoft OLE DB"
            Top             =   600
            Width           =   2175
         End
         Begin VB.TextBox Text10 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   22
            Text            =   "Warning:"
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox Text9 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   21
            Text            =   "Microsoft JET Database"
            Top             =   240
            Width           =   2175
         End
         Begin VB.TextBox Text8 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   20
            Text            =   "Microsoft OLE DB Provider for SQL Server"
            Top             =   960
            Width           =   3735
         End
         Begin VB.TextBox Text7 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   19
            Text            =   "SQL syntax"
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   4080
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Text            =   "Form1.frx":8D25A
         Top             =   2040
         Width           =   3975
      End
      Begin XtremeSuiteControls.WebBrowser WebBrowser3 
         Height          =   975
         Left            =   4080
         TabIndex        =   15
         Top             =   840
         Width           =   3975
         _Version        =   851968
         _ExtentX        =   7011
         _ExtentY        =   1720
         _StockProps     =   173
         BackColor       =   -2147483643
         ScriptErrorsSuppressed=   -1  'True
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "Text5"
         Top             =   480
         Width           =   3495
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "Text4"
         Top             =   4440
         Width           =   3735
      End
      Begin VB.CommandButton Command55 
         Caption         =   "Kaydet"
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
         Left            =   2760
         TabIndex        =   11
         Top             =   4080
         Width           =   1095
      End
      Begin VB.ListBox List2 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2010
         Left            =   120
         TabIndex        =   9
         Top             =   2040
         Width           =   3735
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Text            =   "Form1.frx":8D260
         Top             =   1680
         Width           =   3735
      End
      Begin XtremeSuiteControls.WebBrowser WebBrowser2 
         Height          =   1455
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   3735
         _Version        =   851968
         _ExtentX        =   6588
         _ExtentY        =   2566
         _StockProps     =   173
         BackColor       =   -2147483643
         ScriptErrorsSuppressed=   -1  'True
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Site Kodlarý :"
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
         TabIndex        =   17
         Top             =   1800
         Width           =   1125
      End
      Begin VB.Line Line1 
         X1              =   3960
         X2              =   3960
         Y1              =   240
         Y2              =   6000
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Bulunan Site :"
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
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin XtremeSuiteControls.CommonDialog Cmd2 
         Left            =   2520
         Top             =   4200
         _Version        =   851968
         _ExtentX        =   423
         _ExtentY        =   423
         _StockProps     =   4
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Label4"
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
         Top             =   4080
         Width           =   555
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   720
      Top             =   7800
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   240
      Top             =   7800
   End
   Begin VB.Frame Frame2 
      Caption         =   "-Önizleme"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   4080
      TabIndex        =   0
      Top             =   7560
      Width           =   3975
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   2520
         Width           =   3735
      End
      Begin XtremeSuiteControls.WebBrowser WebBrowser1 
         Height          =   1575
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   3735
         _Version        =   851968
         _ExtentX        =   6588
         _ExtentY        =   2778
         _StockProps     =   173
         BackColor       =   -2147483643
         ScriptErrorsSuppressed=   -1  'True
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Açýlan sayfanýn kodlarý"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Dork :"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   435
      End
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit1 
      Height          =   255
      Left            =   6360
      TabIndex        =   34
      Top             =   1950
      Width           =   1575
      _Version        =   851968
      _ExtentX        =   2778
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
      Locked          =   -1  'True
      Appearance      =   1
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit2 
      Height          =   255
      Left            =   9840
      TabIndex        =   35
      Top             =   1950
      Width           =   1455
      _Version        =   851968
      _ExtentX        =   2566
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
      Locked          =   -1  'True
      Appearance      =   1
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.CommonDialog cmd1 
      Left            =   1680
      Top             =   9240
      _Version        =   851968
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   4
   End
   Begin XtremeSuiteControls.CommonDialog Cmd3 
      Left            =   12720
      Top             =   5520
      _Version        =   851968
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   4
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Taranan Site:"
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
      Left            =   240
      TabIndex        =   33
      Top             =   6510
      Width           =   11200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Toplam Dork:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   260
      TabIndex        =   26
      Top             =   4900
      Width           =   1305
   End
   Begin VB.Image Image1 
      Height          =   7500
      Left            =   0
      Picture         =   "Form1.frx":8D266
      Top             =   0
      Width           =   11700
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub ComboBox1_Click()
Select Case ComboBox1.Text
Case "Genel Arama"
Text.Text = ""
Case ".be  Belçika"
Text.Text = "site:be"
Case ".br  Brezilya"
Text.Text = "site:br"
Case ".by  Beyaz Rusya"
Text.Text = "site:by"
Case ".at  Avusturya"
Text.Text = "site:at"
Case ".ca  Kanada"
Text.Text = "site:ca"
Case ".am  Ermenistan"
Text.Text = "site:am"
Case ".gr  Yunanistan"
Text.Text = "site:gr"
Case ".il  Ýsrail"
Text.Text = "site:il"
Case ".ir  Ýran"
Text.Text = "site:ir"
Case ".it  Ýtalya"
Text.Text = "site:it"
Case ".jp  Japonya"
Text.Text = "site:jp"
Case ".nl  Hollanda"
Text.Text = "site:nl"
Case ".no  Norveç"
Text.Text = "site:no"
Case ".pt  Portekiz"
Text.Text = "site:pt"
Case ".ru  Rusya"
Text.Text = "site:ru"
Case ".se  Ýsveç"
Text.Text = "site:se"
Case ".sy  Suriye"
Text.Text = "site:sy"
Case ".us  ABD"
Text.Text = "site:us"
Case ".cn  Çin"
Text.Text = "site:cn"
Case ".de  Almanya"
Text.Text = "site:de"
Case ".fr  Fransa"
Text.Text = "site:fr"
End Select
End Sub


Private Sub Command1_Click()
'* Dosya Seçme Penceresi Açar
cmd1.ShowOpen
'* List1'i Tamamen Temizler
List1.Clear
'* Ýf Döngüsü Baþlangýç
If Dir(cmd1.FileName) <> "" Then
If cmd1.FileName = "" Then
'* Dosya Seçilmezse Bilgi Mesajý verir
MsgBox "Dosya Seçilmedi", vbInformation, "Uyarý ;"
'* Taramayý Baþlat Butonunu Aktifleþtirir
Command3.Enabled = False
'* Deðilse
Else
'* Seçilen Dosyayý List1'e aktarýr
Open cmd1.FileName For Input As #1
While Not EOF(1)
Input #1, a
List1.AddItem a
Wend
Close #1
'* Label'e Listbox'daki Toplam Deðeri Yansýtýr
Label1.Caption = "Toplam Dork: " & List1.ListCount
'* Taramayý Baþlat Butonunu Aktifleþtirir
Command3.Enabled = True
End If
'* Ýf Döngüsü Bitiþ
End If
End Sub

Private Sub Command2_Click()
'* Deðiþken
Dim Lstekle
'* Açýlan Pencereye Girilecek Deðer
Lstekle = InputBox("Eklemek Ýstediðiniz Dork", "Deðer ekleme ;")
'* Ýf Döngüsü Baþlangýç
If Lstekle = "" Then
'* Eðer Ýnputbox Boþ ise Bilgi Mesajý verir
MsgBox "Lütfen boþ deðer girmeyiniz...", vbInformation, "Bildiri ;"
'* Deðilse
Else
'* Listbox'a Manuel Deðer Ekler
List1.AddItem Lstekle
'* Label'e Listbox'daki Toplam Deðeri Yansýtýr
Label1.Caption = "Toplam Dork: " & List1.ListCount
'* Taramayý Baþlat Butonunu Aktifleþtirir
Command3.Enabled = True
'* Ýf Döngüsü Bitiþ
End If
End Sub

Private Sub Command3_Click()
On Error Resume Next
Command2.Enabled = False
Command1.Enabled = False
ComboBox1.Enabled = False
'* Taramayý Baþlat Butonunu Pasifleþtirir
Command3.Enabled = False
'* Taramayý Durur Butonunu Aktifleþtirir
Command4.Enabled = True
'List1.Clear
List2.Clear
List3.Clear
Kill App.Path & "\linkler.html"
Kill App.Path & "\linkler.txt"
'* Timer1'i Çalýþtýrýr
Timer1.Enabled = True
End Sub



Private Sub Command4_Click()
On Error Resume Next
Dim cevap
cevap = MsgBox("Ýþlemi Durdurulsun mu ?", vbYesNo + vbInformation, "Bildiri ;")
If cevap = vbYes Then
Timer1.Enabled = False
Timer2.Enabled = False
List1.ListIndex = -1
Command4.Enabled = False
Command3.Enabled = True
MsgBox "Dork Taramasý Ýþlemi Durduruldu. Fakat Taramada Bulunan Linkler Aktarýlýyor ve SQL injection Aranýyor..", vbExclamation, "Kullanýcý Ýstemi"
Timer6.Enabled = True
'List1.Clear
List2.Clear
List3.Clear
Command2.Enabled = True
Command1.Enabled = True
Command5.Enabled = False
ComboBox1.Enabled = True
Else
MsgBox "Tarama Ýþlemi Devam Ediyor...", vbInformation, "Bildiri ;"
End If
End Sub



Private Function Kaydet2()
On Error GoTo Son1
 Open App.Path & "\linkler" & ".txt" For Append As #1
Print #1, Text3.Text
Close #1
Son1:
If Dir(App.Path & "\linkler.txt") = "" Then
Else
Open App.Path & "\linkler.txt" For Input As #1
While Not EOF(1)
Input #1, a1
List2.AddItem a1
Label4.Caption = "Toplam Link : " & List2.ListCount
FlatEdit1.Text = List2.ListCount
Command5.Enabled = True
Timer3.Enabled = True
Wend
Close #1
End If
End Function


Sub LstARA(txt As String, list As ListBox)
On Error Resume Next
Dim ktxt As String, klst As String, i As Integer
If txt <> "" Then
For i = 0 To List2.ListCount - 1
ktxt = LCase(txt): klst = LCase(List2.list(i))
If ktxt = klst Then
    List2.ListIndex = i
    
    Exit For
        Else
    If InStr(1, klst, ktxt) Then
        List2.ListIndex = i
        Exit For
            Else: List2.ListIndex = -1
    End If
End If
Next
Else: List2.ListIndex = -1

End If
List2.RemoveItem (i)
End Sub



Private Sub Command5_Click()
On Error Resume Next
Timer3.Enabled = False
List2.RemoveItem ListIndex
List2.RemoveItem ListIndex
Dim i As Integer
Text4.Text = ""
For i = 0 To List2.ListCount - 1
Text4.Text = Text4.Text & List2.list(i) & vbNewLine
Next i
Cmd2.ShowSave
If Cmd2.FileName = "" Then
Else
Open Cmd2.FileName For Append As #1
Print #1, Text4.Text
Print #1, "Toplam Link : " & List2.ListCount
Close #1
End If
End Sub

Private Sub Command6_Click()
List1.Clear
Label1.Caption = "Toplam Dork: " & List1.ListCount
End Sub

Private Sub Form_Load()
On Error Resume Next
    If App.PrevInstance = True Then
        MsgBox "Programý iki kez açamazsýnýz kapatýlýyor...", vbCritical, "Hata;"
        End
    End If
Kill App.Path & "\linkler.html"
Kill App.Path & "\linkler.txt"
On Error Resume Next
'*Cdm1 Pencere Baþlýðý
cmd1.DialogTitle = "Dork Listenizi Seçiniz..."
'*Seçilecek Dosyanýn Uzantýsý
cmd1.Filter = ".txt Dosyasý|*.txt"
Cmd3.DialogTitle = "Kaydedilecek Yer"
Cmd3.Filter = ".txt Dosyasý|*.txt"

Cmd2.DialogTitle = "Kaydedilecek Yer"
Cmd2.Filter = ".txt Dosyasý|*.txt"

ComboBox1.AddItem "Genel Arama"
ComboBox1.AddItem ".be  Belçika"
ComboBox1.AddItem ".br  Brezilya"
ComboBox1.AddItem ".by  Beyaz Rusya"
ComboBox1.AddItem ".at  Avusturya"
ComboBox1.AddItem ".ca  Kanada"
ComboBox1.AddItem ".am  Ermenistan"
ComboBox1.AddItem ".gr  Yunanistan"
ComboBox1.AddItem ".il  Ýsrail"
ComboBox1.AddItem ".ir  Ýran"
ComboBox1.AddItem ".it  Ýtalya"
ComboBox1.AddItem ".jp  Japonya"
ComboBox1.AddItem ".nl  Hollanda"
ComboBox1.AddItem ".no  Norveç"
ComboBox1.AddItem ".pt  Portekiz"
ComboBox1.AddItem ".ru  Rusya"
ComboBox1.AddItem ".se  Ýsveç"
ComboBox1.AddItem ".sy  Suriye"
ComboBox1.AddItem ".us  ABD"
ComboBox1.AddItem ".cn  Çin"
ComboBox1.AddItem ".de  Almanya"
ComboBox1.AddItem ".fr  Fransa"

End Sub

Private Sub Form_Resize()
On Error Resume Next
Form1.Height = 7860
Form1.Width = 11835
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Kill App.Path & "\linkler.html"
Kill App.Path & "\linkler.txt"
End
End Sub

Private Sub List3_Click()
Dim sayfa
sayfa = List3.list(List3.ListIndex)
ShellExecute GetActiveWindow(), "Open", sayfa, "", 0&, 1

End Sub

Private Sub PushButton1_Click()
On Error Resume Next
Dim i As Integer
Text13.Text = ""
For i = 0 To List3.ListCount - 1
Text13.Text = Text13.Text & List3.list(i) & vbNewLine
Next i
Cmd3.ShowSave
If Cmd3.FileName = "" Then
Else
Open Cmd3.FileName For Append As #1
Print #1, Text13.Text
Print #1, "Toplam Link : " & List3.ListCount
Close #1
End If
End Sub

Private Sub PushButton2_Click()
cýkýs = MsgBox("Programdan Çýkmak mý Ýstiyorsunuz?", vbYesNo + vbInformation, "Çýkýþ;")
If cýkýs = vbYes Then
End
End If
End Sub

Private Sub Text2_Change()
On Error Resume Next
Kaydet
End Sub

Private Sub Text3_Change()
On Error Resume Next
Kaydet2
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
With List1
'*Ýndex sayýsýný 1 arttýrýr
.ListIndex = .ListIndex + 1
'*Seçili elemanýn deðeri
Text1.Text = .list(.ListIndex)
WebBrowser1.navigate ("http://www.bing.com/search?q=" & Text1.Text & " " & Text.Text)
'* Listenin son elemanýna gelince bilgi mesajý verir
If .ListIndex = .ListCount - 1 Then
.ListIndex = -1
'* Timer1'i Durdurur
Timer1.Enabled = False
'* Taramayý Baþlat Butonunu Aktifleþtirir
'Command3.Enabled = True
'* Taramayý Durdur Butonunu Aktifleþtirir
'Command4.Enabled = False
'bilgi = MsgBox("Dorklarýn Hepsi Tek Tek Tarandý ve Ýþlem Tamamlandý.", vbOKOnly + vbInformation, "ÝÞLEM YAPILDI")
Timer6.Enabled = True
End If
End With
End Sub
Private Function Ýþlemler()
If Dir(App.Path & "\linkler.html") = "" Then
MsgBox "'linkler.html' Dosyasý Yok", vbExclamation, Me.Caption
Else
WebBrowser2.navigate App.Path & "\linkler" & ".html"
End If
End Function

Private Sub Timer3_Timer()
On Error Resume Next
LstARA WebBrowser2.LocationURL, List2
Label4.Caption = "Toplam Link : " & List2.ListCount
FlatEdit1.Text = List2.ListCount
End Sub

Private Sub Timer4_Timer()
On Error Resume Next
With List2
'*Ýndex sayýsýný 1 arttýrýr
.ListIndex = .ListIndex + 1
'*Seçili elemanýn deðeri
Text5.Text = .list(.ListIndex) & Text12.Text
Label7.Caption = "Taranan Site: " & Text5.Text
WebBrowser3.navigate (Text5.Text)
FlatEdit2.Text = FlatEdit2.Text - 1
'* Listenin son elemanýna gelince bilgi mesajý verir
If .ListIndex = .ListCount - 1 Then
'MsgBox "Dorklarýn Hepsi Tek Tek Tarandý ve Ýþlem Tamamlandý.", vbInformation, "ÝÞLEM YAPILDI"
.ListIndex = -1
'* Timer4'i Durdurur
Timer4.Enabled = False
MsgBox "Tarama Ýþlemi Tamamlandý", vbInformation, "Bildiri;"
Command1.Enabled = True
Command2.Enabled = True
'command3.Enabled =True
End If
If FlatEdit2.Text = "0" Then
FlatEdit2.Text = "Bitti"
Kill App.Path & "\linkler.html"
Kill App.Path & "\linkler.txt"
End If
End With

End Sub

Private Sub Timer5_Timer()
On Error Resume Next
Text6.Text = WebBrowser3.document.body.innerText
If InStr(1, Text6, Text7) Then
List3.AddItem WebBrowser3.LocationURL
End If
If InStr(1, Text6, Text8) Then
List3.AddItem WebBrowser3.LocationURL
End If
If InStr(1, Text6, Text9) Then
List3.AddItem WebBrowser3.LocationURL
End If
If InStr(1, Text6, Text10) Then
List3.AddItem WebBrowser3.LocationURL
End If
If InStr(1, Text6, Text11) Then
List3.AddItem WebBrowser3.LocationURL
End If
Timer5.Enabled = False
End Sub

Private Sub Timer6_Timer()
On Error Resume Next
Ýþlemler
Timer6.Enabled = False
List2.RemoveItem ListIndex
List2.RemoveItem ListIndex
Timer7.Enabled = True

End Sub

Private Sub Timer7_Timer()
On Error Resume Next
Timer3.Enabled = False
Timer4.Enabled = True
Timer7.Enabled = False
FlatEdit2.Text = FlatEdit1.Text
End Sub

Private Sub WebBrowser1_DownloadComplete()
Timer2.Enabled = True
End Sub


Private Sub Timer2_Timer()
'* Timer2 Sürekli Çalýþtýðý için Programýn Hata Vermemesini Saðlar
On Error Resume Next
'* Text2'ye WebBrowser'deki Sitenin div Aralýðýný Alýr
'* Burdaki 'results' Sitedeki div Verilen Name Deðeri
Text2.Text = WebBrowser1.document.All("results").innerHTML
Timer2.Enabled = False
End Sub
Private Function Kaydet()
On Error GoTo Son
Dim Yol
Yol = App.Path & "\linkler" & ".html"
Open Yol For Append As #1
Print #1, Text2.Text
Close #1
Son:
End Function

Private Sub WebBrowser2_DownloadComplete()
'* Kodu KullanaBilmek için veya Hata Alýrsanýz
'* Project>References>Microsoft HTML Object Library Dll'sini Seçiniz
Dim HTMLdoc As HTMLDocument
Dim HTMLlinks As HTMLAnchorElement
Dim HTMLlink As HTMLAnchorElement
Dim STRtxt As String
'Dim STRtxt1 As String
On Error Resume Next
Set HTMLdoc = WebBrowser2.document
For Each HTMLlinks In HTMLdoc.links
STRtxt = STRtxt & HTMLlinks.href & vbCrLf
Next HTMLlinks
'* WebBrowser2'de Buldugu Tüm Linkleri Text2 Aktarýr
Text3.Text = STRtxt
End Sub

Private Sub WebBrowser3_DownloadComplete()
On Error Resume Next
Timer5.Enabled = True
PushButton1.Enabled = True
End Sub
