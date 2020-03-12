VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "ActiveX.ocx"
Begin VB.Form Panel 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   10935
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.WebBrowser WebBrowser1 
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   7080
      Visible         =   0   'False
      Width           =   1215
      _Version        =   851968
      _ExtentX        =   2143
      _ExtentY        =   1931
      _StockProps     =   173
      BackColor       =   -2147483643
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   300
      Left            =   4260
      TabIndex        =   1
      Top             =   1920
      Width           =   6355
      _Version        =   851968
      _ExtentX        =   11210
      _ExtentY        =   529
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
   End
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   4815
      Left            =   240
      TabIndex        =   0
      Top             =   1920
      Width           =   10380
      _Version        =   851968
      _ExtentX        =   18300
      _ExtentY        =   8493
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
      ItemCount       =   4
      SelectedItem    =   1
      Item(0).Caption =   "Lisans Paneli"
      Item(0).ControlCount=   3
      Item(0).Control(0)=   "ListView1"
      Item(0).Control(1)=   "GroupBox2"
      Item(0).Control(2)=   "GroupBox3"
      Item(1).Caption =   "Kontrol Paneli"
      Item(1).ControlCount=   7
      Item(1).Control(0)=   "GroupBox4"
      Item(1).Control(1)=   "PushButton1"
      Item(1).Control(2)=   "ProgressBar1"
      Item(1).Control(3)=   "PushButton2"
      Item(1).Control(4)=   "PushButton3"
      Item(1).Control(5)=   "PushButton4"
      Item(1).Control(6)=   "GroupBox5"
      Item(2).Caption =   "Ayarlar"
      Item(2).ControlCount=   0
      Item(3).Caption =   "Hakkýnda"
      Item(3).ControlCount=   0
      Begin XtremeSuiteControls.ListView ListView1 
         Height          =   3255
         Left            =   -69880
         TabIndex        =   2
         Top             =   1440
         Visible         =   0   'False
         Width           =   10170
         _Version        =   851968
         _ExtentX        =   17939
         _ExtentY        =   5741
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         View            =   3
      End
      Begin XtremeSuiteControls.GroupBox GroupBox5 
         Height          =   615
         Left            =   3000
         TabIndex        =   33
         Top             =   4120
         Width           =   4335
         _Version        =   851968
         _ExtentX        =   7646
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "{ Durum }"
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
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   240
            Width           =   4095
         End
      End
      Begin XtremeSuiteControls.PushButton PushButton4 
         Height          =   300
         Left            =   7440
         TabIndex        =   25
         Top             =   4440
         Width           =   2655
         _Version        =   851968
         _ExtentX        =   4683
         _ExtentY        =   529
         _StockProps     =   79
         Caption         =   "Taramayý Sonlandýr"
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
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton3 
         Height          =   300
         Left            =   7440
         TabIndex        =   24
         Top             =   4160
         Width           =   2655
         _Version        =   851968
         _ExtentX        =   4683
         _ExtentY        =   529
         _StockProps     =   79
         Caption         =   "Benzer Serial Taramasý"
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
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton2 
         Height          =   300
         Left            =   240
         TabIndex        =   23
         Top             =   4440
         Width           =   2655
         _Version        =   851968
         _ExtentX        =   4683
         _ExtentY        =   529
         _StockProps     =   79
         Caption         =   "Taramayý Sonlandýr"
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
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.GroupBox GroupBox4 
         Height          =   3460
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   10170
         _Version        =   851968
         _ExtentX        =   17939
         _ExtentY        =   6103
         _StockProps     =   79
         Caption         =   "[ Çýkan Sonuçlar ]"
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
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.ListView ListView2 
            Height          =   2175
            Left            =   120
            TabIndex        =   26
            Top             =   1200
            Width           =   9945
            _Version        =   851968
            _ExtentX        =   17542
            _ExtentY        =   3836
            _StockProps     =   77
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   162
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            View            =   3
         End
         Begin XtremeSuiteControls.PushButton PushButton7 
            Height          =   255
            Left            =   8040
            TabIndex        =   37
            Top             =   795
            Width           =   2055
            _Version        =   851968
            _ExtentX        =   3625
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Listeyi Temizle"
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
            UseVisualStyle  =   -1  'True
            PushButtonStyle =   3
         End
         Begin XtremeSuiteControls.PushButton PushButton6 
            Height          =   255
            Left            =   9360
            TabIndex        =   36
            Top             =   360
            Width           =   735
            _Version        =   851968
            _ExtentX        =   1296
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Siteyi Aç"
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
            Appearance      =   2
         End
         Begin XtremeSuiteControls.PushButton PushButton5 
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   795
            Width           =   2055
            _Version        =   851968
            _ExtentX        =   3625
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Detaylý Bilgi Göster"
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
            UseVisualStyle  =   -1  'True
            PushButtonStyle =   3
         End
         Begin XtremeSuiteControls.FlatEdit FlatEdit5 
            Height          =   255
            Left            =   6120
            TabIndex        =   32
            Top             =   360
            Width           =   3255
            _Version        =   851968
            _ExtentX        =   5741
            _ExtentY        =   450
            _StockProps     =   77
            ForeColor       =   8421504
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
         End
         Begin XtremeSuiteControls.FlatEdit FlatEdit3 
            Height          =   255
            Left            =   1320
            TabIndex        =   28
            Top             =   360
            Width           =   1095
            _Version        =   851968
            _ExtentX        =   1931
            _ExtentY        =   450
            _StockProps     =   77
            ForeColor       =   8421504
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
         Begin XtremeSuiteControls.FlatEdit FlatEdit4 
            Height          =   255
            Left            =   3720
            TabIndex        =   30
            Top             =   360
            Width           =   1095
            _Version        =   851968
            _ExtentX        =   1931
            _ExtentY        =   450
            _StockProps     =   77
            ForeColor       =   255
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
         Begin VB.Line Line2 
            BorderColor     =   &H00C0C0C0&
            X1              =   0
            X2              =   10200
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Label Label10 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Veri Kaynaðý  :"
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
            Left            =   5040
            TabIndex        =   31
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label9 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Çýkan Sonuçlar :"
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
            Left            =   2520
            TabIndex        =   29
            Top             =   360
            Width           =   1335
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00C0C0C0&
            Index           =   0
            X1              =   0
            X2              =   10200
            Y1              =   1080
            Y2              =   1080
         End
         Begin VB.Label Label8 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Toplam Kayýtlar :"
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
            Left            =   120
            TabIndex        =   27
            Top             =   360
            Width           =   1335
         End
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   300
         Left            =   240
         TabIndex        =   20
         Top             =   4160
         Width           =   2655
         _Version        =   851968
         _ExtentX        =   4683
         _ExtentY        =   529
         _StockProps     =   79
         Caption         =   "Benzer Kullanýcý Adý Taramasý"
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
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.GroupBox GroupBox3 
         Height          =   975
         Left            =   -62320
         TabIndex        =   5
         Top             =   360
         Visible         =   0   'False
         Width           =   2625
         _Version        =   851968
         _ExtentX        =   4630
         _ExtentY        =   1720
         _StockProps     =   79
         BackColor       =   16777215
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.FlatEdit FlatEdit1 
            Height          =   255
            Left            =   720
            TabIndex        =   8
            Top             =   240
            Width           =   1815
            _Version        =   851968
            _ExtentX        =   3201
            _ExtentY        =   450
            _StockProps     =   77
            ForeColor       =   8421504
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
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit FlatEdit2 
            Height          =   255
            Left            =   720
            TabIndex        =   9
            Top             =   600
            Width           =   1815
            _Version        =   851968
            _ExtentX        =   3201
            _ExtentY        =   450
            _StockProps     =   77
            ForeColor       =   8421504
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
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Tarih :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   162
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "Saat :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   162
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   615
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   975
         Left            =   -69880
         TabIndex        =   4
         Top             =   360
         Visible         =   0   'False
         Width           =   7455
         _Version        =   851968
         _ExtentX        =   13150
         _ExtentY        =   1720
         _StockProps     =   79
         Caption         =   "[ Özellikler ]"
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
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.FlatEdit Text2 
            Height          =   255
            Left            =   3120
            TabIndex        =   16
            Top             =   240
            Width           =   1575
            _Version        =   851968
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   77
            ForeColor       =   255
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
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit Text4 
            Height          =   255
            Left            =   840
            TabIndex        =   18
            Top             =   600
            Width           =   1575
            _Version        =   851968
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   77
            ForeColor       =   8421504
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
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit Text5 
            Height          =   255
            Left            =   3120
            TabIndex        =   19
            Top             =   600
            Width           =   4215
            _Version        =   851968
            _ExtentX        =   7435
            _ExtentY        =   450
            _StockProps     =   77
            ForeColor       =   8421504
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
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit Text1 
            Height          =   255
            Left            =   840
            TabIndex        =   15
            Top             =   240
            Width           =   735
            _Version        =   851968
            _ExtentX        =   1296
            _ExtentY        =   450
            _StockProps     =   77
            ForeColor       =   8421504
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
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit Text3 
            Height          =   255
            Left            =   5760
            TabIndex        =   17
            Top             =   240
            Width           =   1575
            _Version        =   851968
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   77
            ForeColor       =   8421504
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
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin VB.Label Label7 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Pc Adý :"
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
            Left            =   120
            TabIndex        =   14
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label6 
            BackColor       =   &H00FFFFFF&
            Caption         =   "IP Adresi :"
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
            Left            =   4800
            TabIndex        =   13
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label5 
            BackColor       =   &H00FFFFFF&
            Caption         =   "K. Adý :"
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
            Left            =   2520
            TabIndex        =   12
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Serial :"
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
            Left            =   2520
            TabIndex        =   11
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Sýra     :"
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
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   735
         End
      End
      Begin XtremeSuiteControls.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   3870
         Width           =   10170
         _Version        =   851968
         _ExtentX        =   17939
         _ExtentY        =   450
         _StockProps     =   93
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
   End
   Begin VB.Image Image1 
      Height          =   7470
      Left            =   0
      Picture         =   "Form1.frx":0000
      Top             =   0
      Width           =   10905
   End
End
Attribute VB_Name = "Panel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'===================={ WebBrowser1 Site Açma}=================
WebBrowser1.Navigate "http://dangerousman32.99k.org/veri/tarih.htm"
'===================={Module From1 Ayarlarý}====================
Panel_Tasarým
'===================={ListView1 Özellikleri}======================
ListView1.ColumnHeaders.Add , , "No:"
ListView1.ColumnHeaders.Add , , "Kullanýcý Adý"
ListView1.ColumnHeaders.Add , , "IP Adresi"
ListView1.ColumnHeaders.Add , , "Bilgisayar Adý"
ListView1.ColumnHeaders.Add , , "Serial No"
ListView1.ColumnHeaders(1).Width = 500
ListView1.ColumnHeaders(2).Width = 1700
ListView1.ColumnHeaders(3).Width = 1700
ListView1.ColumnHeaders(4).Width = 1700
ListView1.ColumnHeaders(5).Width = 3000
'===================={ListView2 Özellikleri}======================
ListView2.ColumnHeaders.Add , , "No:"
ListView2.ColumnHeaders.Add , , "Kullanýcý Adý"
ListView2.ColumnHeaders.Add , , "IP Adresi"
ListView2.ColumnHeaders.Add , , "Serial No"
ListView2.ColumnHeaders(1).Width = 500
ListView2.ColumnHeaders(2).Width = 1700
ListView2.ColumnHeaders(3).Width = 1700
ListView2.ColumnHeaders(4).Width = 3000
End Sub

Private Sub WebBrowser1_DownloadComplete()
On Error Resume Next
FlatEdit2.Text = WebBrowser1.Document.documentElement.Innertext
FlatEdit1.Text = Time
End Sub

