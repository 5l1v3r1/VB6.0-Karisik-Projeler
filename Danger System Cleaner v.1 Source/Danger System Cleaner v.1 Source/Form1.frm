VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "ActiveX.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Danger System Cleaner v.1 // _DangerOusMaN_"
   ClientHeight    =   11025
   ClientLeft      =   2775
   ClientTop       =   1935
   ClientWidth     =   14325
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   11025
   ScaleWidth      =   14325
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   13800
      Top             =   8400
   End
   Begin VB.TextBox sýra5 
      Height          =   285
      Left            =   11400
      TabIndex        =   70
      Top             =   8040
      Width           =   2895
   End
   Begin VB.TextBox sýra4 
      Height          =   285
      Left            =   11400
      TabIndex        =   69
      Top             =   7680
      Width           =   2895
   End
   Begin VB.TextBox sýra3 
      Height          =   285
      Left            =   11400
      TabIndex        =   68
      Top             =   7320
      Width           =   2895
   End
   Begin VB.TextBox sýra2 
      Height          =   285
      Left            =   11400
      TabIndex        =   67
      Top             =   6960
      Width           =   2895
   End
   Begin VB.TextBox sýra1 
      Height          =   285
      Left            =   11400
      TabIndex        =   66
      Top             =   6600
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   11400
      TabIndex        =   62
      Top             =   9000
      Width           =   2895
   End
   Begin VB.Timer Timer 
      Interval        =   150
      Left            =   13320
      Top             =   8400
   End
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   4335
      Left            =   8760
      TabIndex        =   49
      Top             =   2040
      Visible         =   0   'False
      Width           =   5535
      _Version        =   851968
      _ExtentX        =   9763
      _ExtentY        =   7646
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   8
      Color           =   8
      PaintManager.BoldSelected=   -1  'True
      PaintManager.HotTracking=   -1  'True
      ItemCount       =   3
      SelectedItem    =   1
      Item(0).Caption =   "Seçenekler"
      Item(0).ControlCount=   6
      Item(0).Control(0)=   "Label3"
      Item(0).Control(1)=   "CheckBox17"
      Item(0).Control(2)=   "GroupBox2"
      Item(0).Control(3)=   "CheckBox19"
      Item(0).Control(4)=   "CheckBox21"
      Item(0).Control(5)=   "Combo1"
      Item(1).Caption =   "Özel Ek Dosya ve Klasör Temizleme Aracý"
      Item(1).ControlCount=   4
      Item(1).Control(0)=   "ListView2"
      Item(1).Control(1)=   "PushButton6"
      Item(1).Control(2)=   "PushButton8"
      Item(1).Control(3)=   "GroupBox3"
      Item(2).Caption =   "Hakkýnda"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "Image7"
      Begin XtremeSuiteControls.GroupBox GroupBox3 
         Height          =   135
         Left            =   4560
         TabIndex        =   60
         Top             =   2040
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   238
         _StockProps     =   79
         BackColor       =   -2147483633
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton PushButton8 
         Height          =   735
         Left            =   4560
         TabIndex        =   59
         Top             =   2280
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   1296
         _StockProps     =   79
         Caption         =   "Seçiliyi Sil"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton6 
         Height          =   735
         Left            =   4560
         TabIndex        =   58
         Top             =   1320
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   1296
         _StockProps     =   79
         Caption         =   "Ekle"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   3375
         Left            =   120
         TabIndex        =   57
         Top             =   600
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   5953
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin XtremeSuiteControls.CheckBox CheckBox21 
         Height          =   255
         Left            =   -69880
         TabIndex        =   56
         Top             =   1920
         Visible         =   0   'False
         Width           =   5295
         _Version        =   851968
         _ExtentX        =   9340
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Temizlik Bittikten Sonra Uyar / Alarm Çalar"
         ForeColor       =   12632256
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
      End
      Begin XtremeSuiteControls.CheckBox CheckBox19 
         Height          =   255
         Left            =   -69880
         TabIndex        =   55
         Top             =   1440
         Visible         =   0   'False
         Width           =   5295
         _Version        =   851968
         _ExtentX        =   9340
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Temizlik Bittikten Sonra Programý Kapat"
         ForeColor       =   12632256
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
      End
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   615
         Left            =   -69880
         TabIndex        =   53
         Top             =   2400
         Visible         =   0   'False
         Width           =   5295
         _Version        =   851968
         _ExtentX        =   9340
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "[ Saydamlýk ]"
         ForeColor       =   -2147483635
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   2
         Begin XtremeSuiteControls.ScrollBar hs 
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   240
            Width           =   5055
            _Version        =   851968
            _ExtentX        =   8916
            _ExtentY        =   0
            _StockProps     =   64
            Min             =   50
            Max             =   255
            Value           =   255
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
      End
      Begin XtremeSuiteControls.CheckBox CheckBox17 
         Height          =   255
         Left            =   -69880
         TabIndex        =   51
         Top             =   960
         Visible         =   0   'False
         Width           =   5295
         _Version        =   851968
         _ExtentX        =   9340
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Bilgisayar Açýldýðýnda Programda Açýlsýn"
         ForeColor       =   12632256
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
      End
      Begin XtremeSuiteControls.ComboBox Combo1 
         Height          =   315
         Left            =   -68680
         TabIndex        =   52
         Top             =   480
         Visible         =   0   'False
         Width           =   3255
         _Version        =   851968
         _ExtentX        =   5741
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Text            =   "Dil Seç"
      End
      Begin VB.Image Image7 
         Height          =   3240
         Left            =   -69880
         Picture         =   "Form1.frx":2EA5A
         Top             =   960
         Visible         =   0   'False
         Width           =   5280
      End
      Begin VB.Label Label3 
         Caption         =   "Dil Seçeneði :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   -69880
         TabIndex        =   50
         Top             =   480
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Left            =   11880
      Top             =   8400
   End
   Begin XtremeSuiteControls.TabControl Tab2 
      Height          =   4335
      Left            =   120
      TabIndex        =   28
      Top             =   6600
      Visible         =   0   'False
      Width           =   5535
      _Version        =   851968
      _ExtentX        =   9763
      _ExtentY        =   7646
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   8
      Color           =   8
      PaintManager.BoldSelected=   -1  'True
      PaintManager.HotTracking=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      ItemCount       =   2
      SelectedItem    =   1
      Item(0).Caption =   "Baþlangýca Program Ekleme Sihirbazý"
      Item(0).ControlCount=   4
      Item(0).Control(0)=   "PushButton1"
      Item(0).Control(1)=   "PushButton2"
      Item(0).Control(2)=   "ListView1"
      Item(0).Control(3)=   "GroupBox4"
      Item(1).Caption =   "Sistem Ýyileþtirme / Hýzlandýrma [ Uzman Kullanýcýlar için ]"
      Item(1).ControlCount=   17
      Item(1).Control(0)=   "CheckBox7"
      Item(1).Control(1)=   "CheckBox8"
      Item(1).Control(2)=   "CheckBox9"
      Item(1).Control(3)=   "CheckBox10"
      Item(1).Control(4)=   "CheckBox11"
      Item(1).Control(5)=   "CheckBox12"
      Item(1).Control(6)=   "CheckBox13"
      Item(1).Control(7)=   "PushButton4"
      Item(1).Control(8)=   "ProgressBar2"
      Item(1).Control(9)=   "PushButton5"
      Item(1).Control(10)=   "CheckBox14"
      Item(1).Control(11)=   "CheckBox15"
      Item(1).Control(12)=   "CheckBox16"
      Item(1).Control(13)=   "CheckBox20"
      Item(1).Control(14)=   "GroupBox1"
      Item(1).Control(15)=   "Label2"
      Item(1).Control(16)=   "PushButton11"
      Begin XtremeSuiteControls.GroupBox GroupBox4 
         Height          =   30
         Left            =   -65320
         TabIndex        =   63
         Top             =   2040
         Visible         =   0   'False
         Width           =   735
         _Version        =   851968
         _ExtentX        =   1296
         _ExtentY        =   53
         _StockProps     =   79
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton PushButton11 
         Height          =   300
         Left            =   4680
         TabIndex        =   61
         Top             =   3600
         Width           =   735
         _Version        =   851968
         _ExtentX        =   1296
         _ExtentY        =   529
         _StockProps     =   79
         Caption         =   "Uyarý.!"
         BackColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   1695
         Left            =   3120
         TabIndex        =   46
         Top             =   1800
         Width           =   2295
         _Version        =   851968
         _ExtentX        =   4048
         _ExtentY        =   2990
         _StockProps     =   79
         Caption         =   "[ Pc Özellikleri ]"
         ForeColor       =   -2147483635
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   3
         Begin VB.Label Label1 
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   162
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   1335
            Left            =   120
            TabIndex        =   47
            Top             =   240
            Width           =   2055
         End
      End
      Begin XtremeSuiteControls.CheckBox CheckBox20 
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   3240
         Width           =   2895
         _Version        =   851968
         _ExtentX        =   5106
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "8. Sistem Bip'lerini Kaldýrma"
         ForeColor       =   12632256
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
      End
      Begin XtremeSuiteControls.CheckBox CheckBox16 
         Height          =   255
         Left            =   2880
         TabIndex        =   44
         Top             =   1440
         Width           =   2655
         _Version        =   851968
         _ExtentX        =   4683
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "11. Yanýt Verme. Uygulamalar"
         ForeColor       =   12632256
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
      End
      Begin XtremeSuiteControls.CheckBox CheckBox15 
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   1800
         Width           =   2895
         _Version        =   851968
         _ExtentX        =   5106
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "4. Windows Daha da Hýzlandýrma"
         ForeColor       =   12632256
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
      End
      Begin XtremeSuiteControls.CheckBox CheckBox14 
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   2160
         Width           =   2895
         _Version        =   851968
         _ExtentX        =   5106
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "5. Bellek Performansý Arttýrma"
         ForeColor       =   12632256
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton PushButton5 
         Height          =   300
         Left            =   2400
         TabIndex        =   41
         Top             =   3600
         Width           =   2175
         _Version        =   851968
         _ExtentX        =   3836
         _ExtentY        =   529
         _StockProps     =   79
         Caption         =   "Ýyileþtirmeyi Ýptal Et"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ProgressBar ProgressBar2 
         Height          =   300
         Left            =   120
         TabIndex        =   40
         Top             =   3960
         Width           =   4815
         _Version        =   851968
         _ExtentX        =   8493
         _ExtentY        =   529
         _StockProps     =   93
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton PushButton4 
         Height          =   300
         Left            =   120
         TabIndex        =   39
         Top             =   3600
         Width           =   2175
         _Version        =   851968
         _ExtentX        =   3836
         _ExtentY        =   529
         _StockProps     =   79
         Caption         =   "Ýyileþtirmeyi Uygula"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox CheckBox13 
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   2880
         Width           =   2895
         _Version        =   851968
         _ExtentX        =   5106
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "7. Görsel Efektlerin Azaltýlmasý"
         ForeColor       =   12632256
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
      End
      Begin XtremeSuiteControls.CheckBox CheckBox12 
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   2520
         Width           =   2895
         _Version        =   851968
         _ExtentX        =   5106
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "6. Menülerin Hýzlandýrýlmasý"
         ForeColor       =   12632256
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
      End
      Begin XtremeSuiteControls.CheckBox CheckBox11 
         Height          =   255
         Left            =   2880
         TabIndex        =   36
         Top             =   720
         Width           =   2655
         _Version        =   851968
         _ExtentX        =   4683
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "9. %20 internet Hýzý Artýþý"
         ForeColor       =   12632256
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
      End
      Begin XtremeSuiteControls.CheckBox CheckBox10 
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   1080
         Width           =   2655
         _Version        =   851968
         _ExtentX        =   4683
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "2. Daha Fazla Ýþlemci Gücü"
         ForeColor       =   12632256
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
      End
      Begin XtremeSuiteControls.CheckBox CheckBox9 
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   1440
         Width           =   2655
         _Version        =   851968
         _ExtentX        =   4683
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "3. Windows XP'yi Hýzlý Açma"
         ForeColor       =   12632256
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
      End
      Begin XtremeSuiteControls.CheckBox CheckBox8 
         Height          =   255
         Left            =   2880
         TabIndex        =   33
         Top             =   1080
         Width           =   2655
         _Version        =   851968
         _ExtentX        =   4683
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "10. Konuþma Balonlarý"
         ForeColor       =   12632256
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
      End
      Begin XtremeSuiteControls.CheckBox CheckBox7 
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   720
         Width           =   2775
         _Version        =   851968
         _ExtentX        =   4895
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "1.Windows XP’yi Hýzlý Kapatma"
         ForeColor       =   12632256
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3615
         Left            =   -69880
         TabIndex        =   31
         Top             =   360
         Visible         =   0   'False
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   6376
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   12632256
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin XtremeSuiteControls.PushButton PushButton2 
         Height          =   735
         Left            =   -65320
         TabIndex        =   30
         Top             =   2160
         Visible         =   0   'False
         Width           =   735
         _Version        =   851968
         _ExtentX        =   1296
         _ExtentY        =   1296
         _StockProps     =   79
         Caption         =   "Girdiyi Sil"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         Height          =   735
         Left            =   -65320
         TabIndex        =   29
         Top             =   1200
         Visible         =   0   'False
         Width           =   735
         _Version        =   851968
         _ExtentX        =   1296
         _ExtentY        =   1296
         _StockProps     =   79
         Caption         =   "Girdi Ekle"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CommonDialog cmd1 
         Left            =   4800
         Top             =   4080
         _Version        =   851968
         _ExtentX        =   423
         _ExtentY        =   423
         _StockProps     =   4
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Left            =   4920
         TabIndex        =   48
         Top             =   3960
         Width           =   495
         _Version        =   851968
         _ExtentX        =   873
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "&"
         ForeColor       =   8421504
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
      End
   End
   Begin XtremeSuiteControls.FlatEdit Text2 
      Height          =   255
      Left            =   11400
      TabIndex        =   27
      Top             =   9360
      Width           =   2895
      _Version        =   851968
      _ExtentX        =   5106
      _ExtentY        =   450
      _StockProps     =   77
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.ListBox List1 
      Height          =   3615
      Left            =   8520
      TabIndex        =   26
      Top             =   6600
      Width           =   2775
      _Version        =   851968
      _ExtentX        =   4895
      _ExtentY        =   6376
      _StockProps     =   77
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   11400
      Top             =   8400
   End
   Begin XtremeSuiteControls.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   5760
      TabIndex        =   25
      Top             =   10800
      Width           =   5535
      _Version        =   851968
      _ExtentX        =   9763
      _ExtentY        =   238
      _StockProps     =   93
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton Command6 
      Height          =   375
      Left            =   10080
      TabIndex        =   24
      Top             =   10320
      Width           =   1215
      _Version        =   851968
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Ýptal Et"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton Command5 
      Height          =   375
      Left            =   8520
      TabIndex        =   17
      Top             =   10320
      Width           =   1455
      _Version        =   851968
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Temizliðe Baþla"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.TabControl Tab1 
      CausesValidation=   0   'False
      Height          =   4095
      Left            =   5760
      TabIndex        =   4
      Top             =   6600
      Width           =   2655
      _Version        =   851968
      _ExtentX        =   4683
      _ExtentY        =   7223
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   8
      Color           =   8
      PaintManager.BoldSelected=   -1  'True
      PaintManager.HotTracking=   -1  'True
      ItemCount       =   2
      Item(0).Caption =   "Windows"
      Item(0).ControlCount=   9
      Item(0).Control(0)=   "Image6"
      Item(0).Control(1)=   "CheckBox1"
      Item(0).Control(2)=   "CheckBox2"
      Item(0).Control(3)=   "CheckBox3"
      Item(0).Control(4)=   "CheckBox4"
      Item(0).Control(5)=   "CheckBox5"
      Item(0).Control(6)=   "CheckBox6"
      Item(0).Control(7)=   "GroupBox5"
      Item(0).Control(8)=   "CheckBox18"
      Item(1).Caption =   "Tarayýcýlar"
      Item(1).ControlCount=   15
      Item(1).Control(0)=   "Image3"
      Item(1).Control(1)=   "Image4"
      Item(1).Control(2)=   "Image5"
      Item(1).Control(3)=   "Check12"
      Item(1).Control(4)=   "Check11"
      Item(1).Control(5)=   "Check10"
      Item(1).Control(6)=   "Check9"
      Item(1).Control(7)=   "Check8"
      Item(1).Control(8)=   "Check7"
      Item(1).Control(9)=   "Check6"
      Item(1).Control(10)=   "Check5"
      Item(1).Control(11)=   "Check4"
      Item(1).Control(12)=   "Check3"
      Item(1).Control(13)=   "Check2"
      Item(1).Control(14)=   "Check1"
      Begin XtremeSuiteControls.CheckBox CheckBox18 
         Height          =   615
         Left            =   120
         TabIndex        =   65
         Top             =   2160
         Width           =   2415
         _Version        =   851968
         _ExtentX        =   4260
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Özel Ek Dosya ve Klasördeki Seçili Ögeleri Temizleme"
         ForeColor       =   12632256
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextAlignment   =   4
         Appearance      =   6
      End
      Begin XtremeSuiteControls.GroupBox GroupBox5 
         Height          =   30
         Left            =   120
         TabIndex        =   64
         Top             =   2040
         Width           =   2415
         _Version        =   851968
         _ExtentX        =   4260
         _ExtentY        =   53
         _StockProps     =   79
         Appearance      =   6
      End
      Begin XtremeSuiteControls.CheckBox CheckBox6 
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1680
         Width           =   2775
         _Version        =   851968
         _ExtentX        =   4895
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Gereksiz Temp Dosyalarý"
         ForeColor       =   12632256
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
      End
      Begin XtremeSuiteControls.CheckBox CheckBox5 
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1440
         Width           =   2775
         _Version        =   851968
         _ExtentX        =   4895
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Son Açýlan Dosyalar"
         ForeColor       =   12632256
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
      End
      Begin XtremeSuiteControls.CheckBox CheckBox4 
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1200
         Width           =   2775
         _Version        =   851968
         _ExtentX        =   4895
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Gereksiz Dosyalar"
         ForeColor       =   12632256
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
      End
      Begin XtremeSuiteControls.CheckBox CheckBox3 
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   960
         Width           =   2775
         _Version        =   851968
         _ExtentX        =   4895
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Baþlat Menüsü Kýsa Yollarý"
         ForeColor       =   12632256
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
      End
      Begin XtremeSuiteControls.CheckBox CheckBox2 
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   720
         Width           =   2775
         _Version        =   851968
         _ExtentX        =   4895
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Çöp Kutusunu Bosalt"
         ForeColor       =   12632256
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
      End
      Begin XtremeSuiteControls.CheckBox CheckBox1 
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   360
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Sistem"
         ForeColor       =   -2147483635
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
      End
      Begin XtremeSuiteControls.CheckBox Check12 
         Height          =   255
         Left            =   -69880
         TabIndex        =   16
         Top             =   3840
         Visible         =   0   'False
         Width           =   2775
         _Version        =   851968
         _ExtentX        =   4895
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Ýndirme Geçmiþi "
         ForeColor       =   12632256
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
      End
      Begin XtremeSuiteControls.CheckBox Check11 
         Height          =   255
         Left            =   -69880
         TabIndex        =   15
         Top             =   3600
         Visible         =   0   'False
         Width           =   2535
         _Version        =   851968
         _ExtentX        =   4471
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Kayýtlý Þifreler"
         ForeColor       =   12632256
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
      End
      Begin XtremeSuiteControls.CheckBox Check10 
         Height          =   255
         Left            =   -69880
         TabIndex        =   14
         Top             =   3360
         Visible         =   0   'False
         Width           =   2415
         _Version        =   851968
         _ExtentX        =   4260
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Ýnternet Geçmiþi"
         ForeColor       =   12632256
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
      End
      Begin XtremeSuiteControls.CheckBox Check9 
         Height          =   255
         Left            =   -69880
         TabIndex        =   13
         Top             =   2640
         Visible         =   0   'False
         Width           =   2775
         _Version        =   851968
         _ExtentX        =   4895
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Ýndirme Geçmiþi "
         ForeColor       =   12632256
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
      End
      Begin XtremeSuiteControls.CheckBox Check8 
         Height          =   255
         Left            =   -69880
         TabIndex        =   12
         Top             =   2400
         Visible         =   0   'False
         Width           =   2415
         _Version        =   851968
         _ExtentX        =   4260
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Kayýtlý Þifreler"
         ForeColor       =   12632256
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
      End
      Begin XtremeSuiteControls.CheckBox Check7 
         Height          =   255
         Left            =   -69880
         TabIndex        =   11
         Top             =   2160
         Visible         =   0   'False
         Width           =   2415
         _Version        =   851968
         _ExtentX        =   4260
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Ýnternet Geçmiþi"
         ForeColor       =   12632256
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
      End
      Begin XtremeSuiteControls.CheckBox Check6 
         Height          =   255
         Left            =   -69880
         TabIndex        =   10
         Top             =   1440
         Visible         =   0   'False
         Width           =   2775
         _Version        =   851968
         _ExtentX        =   4895
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Ýndirme Geçmiþi"
         ForeColor       =   12632256
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
      End
      Begin XtremeSuiteControls.CheckBox Check5 
         Height          =   255
         Left            =   -69880
         TabIndex        =   9
         Top             =   1200
         Visible         =   0   'False
         Width           =   2415
         _Version        =   851968
         _ExtentX        =   4260
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Kayýtlý Þifreler"
         ForeColor       =   12632256
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
      End
      Begin XtremeSuiteControls.CheckBox Check4 
         Height          =   255
         Left            =   -69880
         TabIndex        =   8
         Top             =   960
         Visible         =   0   'False
         Width           =   2415
         _Version        =   851968
         _ExtentX        =   4260
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Ýnternet Geçmiþi"
         ForeColor       =   12632256
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
      End
      Begin XtremeSuiteControls.CheckBox Check3 
         Height          =   255
         Left            =   -69640
         TabIndex        =   7
         Top             =   3000
         Visible         =   0   'False
         Width           =   1575
         _Version        =   851968
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Google Chrome"
         ForeColor       =   -2147483635
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
      End
      Begin XtremeSuiteControls.CheckBox Check2 
         Height          =   255
         Left            =   -69640
         TabIndex        =   6
         Top             =   1800
         Visible         =   0   'False
         Width           =   2415
         _Version        =   851968
         _ExtentX        =   4260
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Firefox"
         ForeColor       =   -2147483635
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
      End
      Begin XtremeSuiteControls.CheckBox Check1 
         Height          =   255
         Left            =   -69640
         TabIndex        =   5
         Top             =   600
         Visible         =   0   'False
         Width           =   1695
         _Version        =   851968
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Ýnternet Explorer"
         ForeColor       =   -2147483635
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
      End
      Begin VB.Image Image6 
         Height          =   300
         Left            =   0
         Picture         =   "Form1.frx":42F1E
         Top             =   360
         Width           =   300
      End
      Begin VB.Image Image4 
         Height          =   300
         Left            =   -70000
         Picture         =   "Form1.frx":45476
         Top             =   600
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Image Image3 
         Height          =   300
         Left            =   -70000
         Picture         =   "Form1.frx":478CB
         Top             =   1800
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Image Image5 
         Height          =   300
         Left            =   -70000
         Picture         =   "Form1.frx":47D75
         Top             =   3000
         Visible         =   0   'False
         Width           =   300
      End
   End
   Begin XtremeSuiteControls.PushButton Command4 
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   5160
      Width           =   2655
      _Version        =   851968
      _ExtentX        =   4683
      _ExtentY        =   1931
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Picture         =   "Form1.frx":4A265
   End
   Begin XtremeSuiteControls.PushButton Command3 
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   4080
      Width           =   2655
      _Version        =   851968
      _ExtentX        =   4683
      _ExtentY        =   1931
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Picture         =   "Form1.frx":50159
   End
   Begin XtremeSuiteControls.PushButton Command2 
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Width           =   2655
      _Version        =   851968
      _ExtentX        =   4683
      _ExtentY        =   1931
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Picture         =   "Form1.frx":54C93
   End
   Begin XtremeSuiteControls.PushButton Command1 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   2655
      _Version        =   851968
      _ExtentX        =   4683
      _ExtentY        =   1931
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Picture         =   "Form1.frx":5A2FF
   End
   Begin VB.Image Image 
      Height          =   495
      Left            =   2880
      Top             =   1320
      Width           =   5295
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   8640
      X2              =   0
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   0
      Y1              =   6360
      Y2              =   0
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   2880
      X2              =   2880
      Y1              =   6360
      Y2              =   1800
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FFFFFF&
      X1              =   8640
      X2              =   8640
      Y1              =   6360
      Y2              =   0
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      X1              =   -360
      X2              =   8640
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Image Image2 
      Height          =   4650
      Left            =   0
      Picture         =   "Form1.frx":5FE96
      Top             =   1800
      Width           =   8655
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   8640
      X2              =   -600
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Image Image1 
      Height          =   1950
      Left            =   0
      Picture         =   "Form1.frx":69BC0
      Top             =   -120
      Width           =   8715
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function sndPlaySound Lib "WINMM.DLL" Alias "sndPlaySoundA" (ByVal lpszSoundName As Any, ByVal uFlags As Long) As Long
Dim SoundBuffer As Long
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Private Sub CheckBox18_Click()
If CheckBox18.Value = 1 Then
Timer1.Enabled = True
Else
Timer1.Enabled = False
End If
End Sub

Private Sub Form_Initialize()
InitCommonControls
End Sub
Private Sub Check1_Click()
If Check1.Value = 1 Then
Check4.Value = 1
Check5.Value = 1
Check6.Value = 1
Else
Check4.Value = 0
Check5.Value = 0
Check6.Value = 0
End If
End Sub
Private Sub Check10_Click()
If Check10.Value = 1 Then
Timer1.Enabled = True
Else
Timer1.Enabled = False
End If
End Sub
Private Sub Check11_Click()
If Check11.Value = 1 Then
Timer1.Enabled = True
Else
Timer1.Enabled = False
End If
End Sub
Private Sub Check12_Click()
If Check12.Value = 1 Then
Timer1.Enabled = True
Else
Timer1.Enabled = False
End If
End Sub
Private Sub Check2_Click()
If Check2.Value = 1 Then
Check7.Value = 1
Check8.Value = 1
Check9.Value = 1
Else
Check7.Value = 0
Check8.Value = 0
Check9.Value = 0
End If
End Sub
Private Sub Check3_Click()
If Check3.Value = 1 Then
Check10.Value = 1
Check11.Value = 1
Check12.Value = 1
Else
Check10.Value = 0
Check11.Value = 0
Check12.Value = 0
End If
End Sub
Private Sub Check4_Click()
If Check4.Value = 1 Then
Timer1.Enabled = True
Else
Timer1.Enabled = False
End If
End Sub
Private Sub Check5_Click()
If Check5.Value = 1 Then
Timer1.Enabled = True
Else
Timer1.Enabled = False
End If
End Sub
Private Sub Check6_Click()
If Check6.Value = 1 Then
Timer1.Enabled = True
Else
Timer1.Enabled = False
End If
End Sub
Private Sub Check7_Click()
If Check7.Value = 1 Then
Timer1.Enabled = True
Else
Timer1.Enabled = False
End If
End Sub
Private Sub Check8_Click()
If Check8.Value = 1 Then
Timer1.Enabled = True
Else
Timer1.Enabled = False
End If
End Sub
Private Sub Check9_Click()
If Check9.Value = 1 Then
Timer1.Enabled = True
Else
Timer1.Enabled = False
End If
End Sub
Private Sub CheckBox1_Click()
If CheckBox1.Value = 1 Then
CheckBox2.Value = 1
CheckBox3.Value = 1
CheckBox4.Value = 1
CheckBox5.Value = 1
CheckBox6.Value = 1
Else
CheckBox2.Value = 0
CheckBox3.Value = 0
CheckBox4.Value = 0
CheckBox5.Value = 0
CheckBox6.Value = 0
End If
End Sub
Private Sub CheckBox10_Click()
If CheckBox10.Value = 1 Then
Timer2.Enabled = True
Else
Timer2.Enabled = False
End If
End Sub
Private Sub CheckBox11_Click()
If CheckBox11.Value = 1 Then
Timer2.Enabled = True
Else
Timer2.Enabled = False
End If
End Sub
Private Sub CheckBox12_Click()
If CheckBox12.Value = 1 Then
Timer2.Enabled = True
Else
Timer2.Enabled = False
End If
End Sub
Private Sub CheckBox13_Click()
If CheckBox13.Value = 1 Then
Timer2.Enabled = True
Else
Timer2.Enabled = False
End If
End Sub

Private Sub CheckBox14_Click()
If CheckBox14.Value = 1 Then
Timer2.Enabled = True
Else
Timer2.Enabled = False
End If
End Sub

Private Sub CheckBox15_Click()
If CheckBox15.Value = 1 Then
Timer2.Enabled = True
Else
Timer2.Enabled = False
End If
End Sub

Private Sub CheckBox16_Click()
If CheckBox16.Value = 1 Then
Timer2.Enabled = True
Else
Timer2.Enabled = False
End If
End Sub

Private Sub CheckBox17_Click()
If CheckBox17.Value = 1 Then
Dim KayitDefteri As Object
Dim reg As Object
Set KayitDefteri = CreateObject("wscript.shell")
KayitDefteri.regwrite "HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN\" & App.EXEName, App.Path & "\" & App.EXEName & ".exe"
Else
Set reg = CreateObject("wscript.shell")
reg.regdelete "HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN\" & App.EXEName
End If
End Sub

Private Sub CheckBox19_Click()
If CheckBox19.Value = 1 Then
Timer4.Interval = 1000
Else
Timer4.Enabled = False
Timer4.Interval = 0
End If
End Sub

Private Sub CheckBox2_Click()
If CheckBox2.Value = 1 Then
Timer1.Enabled = True
Else
Timer1.Enabled = False
End If
End Sub

Private Sub CheckBox20_Click()
If CheckBox20Value = 1 Then
Timer2.Enabled = True
Else
Timer2.Enabled = False
End If
End Sub

Private Sub CheckBox21_Click()
On Error Resume Next
MsgBox "Ýþlem Bittiðinde Bilgilendirmek için Alarm Çalacaktýr Bilginize .!", 16, "UYARI"
If CheckBox21.Value = 1 Then
Timer3.Interval = 1000
Else
Timer3.Enabled = False
Timer3.Interval = 0
End If
End Sub

Private Sub CheckBox3_Click()
If CheckBox3.Value = 1 Then
Timer1.Enabled = True
Else
Timer1.Enabled = False
End If
End Sub

Private Sub CheckBox4_Click()
If CheckBox4.Value = 1 Then
Timer1.Enabled = True
Else
Timer1.Enabled = False
End If
End Sub

Private Sub CheckBox5_Click()
If CheckBox5.Value = 1 Then
Timer1.Enabled = True
Else
Timer1.Enabled = False
End If
End Sub

Private Sub CheckBox6_Click()
If CheckBox6.Value = 1 Then
Timer1.Enabled = True
Else
Timer1.Enabled = False
End If
End Sub

Private Sub CheckBox7_Click()
If CheckBox7.Value = 1 Then
Timer2.Enabled = True
Else
Timer2.Enabled = False
End If
End Sub

Private Sub CheckBox8_Click()
If CheckBox8.Value = 1 Then
Timer2.Enabled = True
Else
Timer2.Enabled = False
End If
End Sub

Private Sub CheckBox9_Click()
If CheckBox9.Value = 1 Then
Timer2.Enabled = True
Else
Timer2.Enabled = False
End If
End Sub

Private Sub Combo1_Click()
If Combo1.Text = "Türkçe" Then
'Turkce
Else
End If
If Combo1.Text = "English" Then
English
End If
End Sub

Private Sub Command1_Click()
'Gösterme
Tab1.Visible = True
ProgressBar1.Visible = True
List1.Visible = True
Command5.Visible = True
Command6.Visible = True
'Gizleme
Tab2.Visible = False
TabControl1.Visible = False

Timer3.Enabled = True
End Sub

Private Sub Command2_Click()
'Gizleme
Tab1.Visible = False
ProgressBar1.Visible = False
List1.Visible = False
Command5.Visible = False
Command6.Visible = False
TabControl1.Visible = False
'Gösterme
Tab2.Visible = True

End Sub

Private Sub Command3_Click()
Tab1.Visible = False
Tab2.Visible = False
ProgressBar1.Visible = False
List1.Visible = False
Command5.Visible = False
Command6.Visible = False
TabControl1.Visible = True
End Sub

Private Sub Command4_Click()
cýkýs = MsgBox("Programdan Çýkmak mý Ýstiyorsunuz ?", vbQuestion + vbYesNo, " Çýkýþ ;")
If cýkýs = vbYes Then
End
End If
End Sub

Private Sub Command5_Click()
Timer1.Enabled = True
Timer1.Interval = 110
Command5.Enabled = False
Tab1.Enabled = False
Command6.Enabled = True
List1.AddItem "Temizleme Baþladý"
End Sub

Private Sub Command6_Click()
Timer1.Enabled = False
ProgressBar1.Value = 0
List1.Clear
Command5.Enabled = True
Command6.Enabled = False
Tab1.Enabled = True
List1.AddItem "Temizleme Ýptal Edildi"
End Sub

Private Sub Form_Load()

check_gizle
'Listview1 Özellikleri
ListView1.View = lvwReport
ListView1.ColumnHeaders.Add , , "No:"
ListView1.ColumnHeaders.Add , , "Program"
ListView1.ColumnHeaders.Add , , "Dosya Yolu"
ListView1.ColumnHeaders(1).Width = 500
ListView1.ColumnHeaders(2).Width = 1700
ListView1.ColumnHeaders(3).Width = 5000
'Listview2 Özellikleri
ListView2.View = lvwReport
ListView2.ColumnHeaders.Add , , "No:"
ListView2.ColumnHeaders.Add , , "Temizlenecek Dosya ve Klasörler"
ListView2.ColumnHeaders(1).Width = 500
ListView2.ColumnHeaders(2).Width = 7000
'------------------------------------------------------------------------
CheckBox1.Value = 1
Check1.Value = 1
Check2.Value = 1
Check3.Value = 1
Menu
Text2.Text = UserName
Dim bellek As MEMORYSTATUS
GlobalMemoryStatus bellek
Label1.Caption = "" & ReadKey("HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System\CentralProcessor\0\ProcessorNameString") & "                                " & Int(bellek.dwTotalPhys / 1024 / 1024) & " " & "MB RAM" & "   " & ReadKey("HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System\CentralProcessor\0\VendorIdentifier") & "                    Ýþlemci Hýzý : " & ReadKey("HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System\CentralProcessor\0\~MHz")

cmd1.DialogTitle = "Eklemek Ýstediðiniz Programý Seçiniz."
cmd1.Filter = "Exe|*.exe"

Combo1.AddItem "Türkçe"
Combo1.AddItem "English"
End Sub


Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub hs_Change()
MakeTransparent Form1.hwnd, hs.Value
End Sub

Private Sub Image_Click()
On Error Resume Next
Siteler
End Sub



Private Sub ListView1_Click()
If ListView1.ListItems.Count = 0 Then Exit Sub
X = ListView1.SelectedItem.Index
Text1.Text = ListView1.ListItems(X).ListSubItems(1).Text
End Sub

Private Sub PushButton1_Click()
On Error Resume Next
cmd1.ShowOpen
girdi = MsgBox("Seçtiðiniz Program Bilgisayar Baþlangýcýna Eklensin mi ?", vbInformation + vbYesNo, "Bildiri ;")
If girdi = vbYes Then
'----------------------------------
ProgramÝsmi = cmd1.FileTitle
yol = cmd1.FileName
Aktif = "Aktif"
Dim Baþlangýç As ListItem
Set Baþlangýç = ListView1.ListItems.Add
Baþlangýç.SubItems(0) = Aktif
Baþlangýç.SubItems(1) = ProgramÝsmi
Baþlangýç.SubItems(2) = yol
'----------------------------------
Dim KayitDefteri As Object
Dim reg As Object
Set KayitDefteri = CreateObject("wscript.shell")
KayitDefteri.regwrite "HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN\" & ProgramÝsmi, yol & "\" & ProgramÝsmi & ".exe"
End If
End Sub


Private Sub PushButton11_Click()
MsgBox "Sistem Ýyileþtirmede Kullanýlan Regedit Ayarlarý ve Girilen Deðerlerini Ayrýntýlý Olarak Görebilirsiniz.", vbInformation, "Bildiri ;"
Form3.Show
End Sub

Private Sub PushButton2_Click()
On Error Resume Next
Dim KayitDefteri As Object
Dim reg As Object
ListView1.ListItems.Remove ListView1.SelectedItem.Index
Set reg = CreateObject("wscript.shell")
reg.regdelete "HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN\" & Text1.Text
End Sub

Private Sub PushButton3_Click()
SendKeys "{F5}"
End Sub

Private Sub PushButton4_Click()
Timer2.Interval = 100
Timer2.Enabled = True
Label2.Caption = ""
PushButton4.Enabled = False
PushButton5.Enabled = True
check_gizle
End Sub

Private Sub PushButton5_Click()
Timer2.Enabled = False
Label2.Caption = ""
ProgressBar2.Value = 0
PushButton4.Enabled = False
PushButton4.Enabled = True
check_göster
End Sub



Private Sub PushButton6_Click()
Form2.Show
Form2.Text1.Text = ""
Form2.Text2.Text = ""
End Sub

Private Sub PushButton8_Click()
ListView2.ListItems.Remove ListView2.SelectedItem.Index
End Sub

Private Sub PushButton9_Click()

End Sub

Private Sub Timer_Timer()
On Error Resume Next
Form1.Width = 8805
Form1.Height = 6900
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
ProgressBar1.Value = ProgressBar1.Value + 1
'========================
If ProgressBar1 = 25 Then
If Check4.Value = 1 Then ' Ýnternet Explorer
List1.AddItem "              [ Ýnternet Explorer ]             "
If Check4.Value = 1 Then
        Kill "C:\Documents and Settings\" & Text2.Text & "\Cookies\*.*"
        Kill "C:\Documents and Settings\" & Text2.Text & "\Local Settings\Temporary Internet Files\Content.IE5\GAYG1ADO\search[1]"
        Kill "C:\Documents and Settings\" & Text2.Text & "\Local Settings\Temporary Internet Files\Content.IE5\R0CAIXNO\syntax[1]"
        Kill "C:\Documents and Settings\" & Text2.Text & "\Local Settings\Temporary Internet Files\Content.IE5\ZIZ301QG\pagerror[1]"
        List1.AddItem "Ýnternet Geçmiþi Temizlendi."
End If

If Check5.Value = 1 Then

        List1.AddItem "Kayýtlý Þifreler Temizlendi."
End If
If Check6.Value = 1 Then
        Kill "C:\Documents and Settings\" & Text2.Text & "\Local Settings\Temporary Internet Files\Content.IE5\R6PSY2HZ"
        Kill "C:\Documents and Settings\" & Text2.Text & "\Local Settings\Temporary Internet Files\Content.IE5\9IPRQBY4"
        List1.AddItem "Ýndirme Geçmiþi Temizlendi."
End If
End If
End If
'================================================================================================
If ProgressBar1 = 40 Then
If Check2.Value = 1 Then ' Firefox
List1.AddItem "============================="
List1.AddItem "                    [ Firefox ]                       "
 If Check7.Value = 1 Then
             Kill "C:\Documents and Settings\" & Text2.Text & "\Local Settings\Application Data\Mozilla\Firefox\Profiles\7ptg83hm.default\Cache\*.*"
             Kill "C:\Documents and Settings\" & Text2.Text & "\Application Data\Mozilla\Firefox\Profiles\7ptg83hm.default\downloads.sqlite"
             List1.AddItem "Ýnternet Geçmiþi Temizlendi."
End If
If Check8.Value = 1 Then
              Kill "C:\Documents and Settings\" & Text2.Text & "\Application Data\Mozilla\Firefox\Profiles\7ptg83hm.default\key3.db"
              Kill "C:\Documents and Settings\" & Text2.Text & "\Application Data\Mozilla\Firefox\Profiles\7ptg83hm.default\signons.sqlite"
              Kill "C:\Documents and Settings\" & Text2.Text & "\Application Data\Mozilla\Firefox\Profiles\7ptg83hm.default\formhistory.sqlite"
              List1.AddItem "Kayýtlý Þifreler Temizlendi."
End If
If Check9.Value = 1 Then
             Kill "C:\Documents and Settings\" & Text2.Text & "\Belgelerim\Downloads\*.*"
             List1.AddItem "Ýndirme Geçmiþi Temizlendi."
End If
End If
End If
'================================================================================================
If ProgressBar1 = 75 Then
If Check3.Value = 1 Then ' Google Chrome
List1.AddItem "============================="
List1.AddItem "            [ Google Chrome ]              "
 If Check10.Value = 1 Then
            List1.AddItem "Ýnternet Geçmiþi Temizlendi."
           Kill "C:\Documents and Settings\" & Text2.Text & "\Local Settings\Application Data\Google\Chrome\User Data\Default\Cache\*.*"
           Kill "C:\Documents and Settings\" & Text2.Text & "\Local Settings\Application Data\Google\Chrome\User Data\Default\History Index 2011-03"
           Kill "C:\Documents and Settings\" & Text2.Text & "\Local Settings\Application Data\Google\Chrome\User Data\Default\Archived History"
           Kill "C:\Documents and Settings\" & Text2.Text & "\Local Settings\Application Data\Google\Chrome\User Data\Default\Visited Links"
           Kill "C:\Documents and Settings\" & Text2.Text & "\Local Settings\Application Data\Google\Chrome\User Data\Default\Current Tabs"
End If
If Check11.Value = 1 Then
           Kill "C:\Documents and Settings\" & Text2.Text & "\Local Settings\Application Data\Google\Chrome\User Data\Default\Login Data"
           Kill "C:\Documents and Settings\" & Text2.Text & "\Local Settings\Application Data\Google\Chrome\User Data\Default\Current Session"
           Kill "C:\Documents and Settings\" & Text2.Text & "\Local Settings\Application Data\Google\Chrome\User Data\Default\Last Session"
           List1.AddItem "Kayýtlý Þifreler Temizlendi."
End If
If Check12.Value = 1 Then
           Kill "C:\Documents and Settings\" & Text2.Text & "\Belgelerim\Karþýdan Yüklenenler\*.*"
           List1.AddItem "Ýndirme Geçmiþi Temizlendi."
End If
End If
End If
'================================================================================================
If ProgressBar1 = 89 Then
If CheckBox1.Value = 1 Then ' sistem
List1.AddItem "============================="
List1.AddItem "                 [ Sistem ]              "
If CheckBox2.Value = 1 Then
         
        List1.AddItem "Çöp Kutusu Temizlendi."
End If
If CheckBox3.Value = 1 Then
         
        List1.AddItem "Baþlat Menüsü Temizlendi."
End If
If CheckBox4.Value = 1 Then
        Kill "C:\WINDOWS\system32\wbem\Logs\*.*"
        List1.AddItem "Gereksiz Dosyalar Temizlendi."
End If
If CheckBox5.Value = 1 Then
        Kill "C:\Documents and Settings\" & Text2.Text & "\Recent\*.*"
        List1.AddItem "Son Açýlan Dosyalar Temizlendi."
End If
If CheckBox6.Value = 1 Then
Kill "C:\WINDOWS\TEMP\*.*"
Kill "C:\WINDOWS\TEMP\HTT389.tmp"
Kill "C:\WINDOWS\TEMP\HTT38A.tmp"
Kill "C:\WINDOWS\TEMP\HTT38B.tmp"
Kill "C:\WINDOWS\TEMP\HWIDS.txt"
Kill "C:\WINDOWS\TEMP\NOD38D.tmp"
Kill "C:\Documents and Settings\" & Text2.Text & "\Local Settings\Temp\2yr5592.tmp"
Kill "C:\Documents and Settings\" & Text2.Text & "\Local Settings\Temp\31r5596.tmp  "
Kill "C:\Documents and Settings\" & Text2.Text & "\Local Settings\Temp\45r5598.tmp"
Kill "C:\Documents and Settings\" & Text2.Text & "\Local Settings\Temp\68r559A.tmp"
Kill "C:\Documents and Settings\" & Text2.Text & "\Local Settings\Temp\7cs559C.tmp"
Kill "C:\Documents and Settings\" & Text2.Text & "\Local Settings\Temp\amt.log"
Kill "C:\Documents and Settings\" & Text2.Text & "\Local Settings\Temp\boost_interprocess\DDM0serviceCmdLock"
Kill "C:\Documents and Settings\" & Text2.Text & "\Local Settings\Temp\boost_interprocess\DDM0serviceCmdSerializeLock"
Kill "C:\Documents and Settings\" & Text2.Text & "\Local Settings\Temp\boost_interprocess\DDM0serviceCmdShared"
Kill "C:\Documents and Settings\" & Text2.Text & "\Local Settings\Temp\boost_interprocess\DDM0serviceLock"
Kill "C:\Documents and Settings\" & Text2.Text & "\Local Settings\Temp\DDMCache\149465.avi.ddp"
Kill "C:\Documents and Settings\" & Text2.Text & "\Local Settings\Temp\DDMCache\149465.avi.ddr"
Kill "C:\Documents and Settings\" & Text2.Text & "\Local Settings\Temp\dzr5595.tmp  "
Kill "C:\Documents and Settings\" & Text2.Text & "\Local Settings\Temp\etilqs_QBgCdhvjhqAR4VEB9Wqt  "
Kill "C:\Documents and Settings\" & Text2.Text & "\Local Settings\Temp\etilqs_V2SdpMgfd3jUe9ImhQON"
Kill "C:\Documents and Settings\" & Text2.Text & "\Local Settings\Temp\etilqs_X59EMPjDIZZG6F3EjscR"
Kill "C:\Documents and Settings\" & Text2.Text & "\Local Settings\Temp\flaE71.tmp"
Kill "C:\Documents and Settings\" & Text2.Text & "\Local Settings\Temp\ids559F.tmp"
Kill "C:\Documents and Settings\" & Text2.Text & "\Local Settings\Temp\iu65426.tmp  "
Kill "C:\Documents and Settings\" & Text2.Text & "\Local Settings\Temp\MessengerCache\92FBGvcEQbVhIcl0O2YURovqSl4E="
Kill "C:\Documents and Settings\" & Text2.Text & "\Local Settings\Temp\MessengerCache\g2FbcTf1VmNTJyKJTG2pjR3V2FrA4="
Kill "C:\Documents and Settings\" & Text2.Text & "\Local Settings\Temp\MessengerCache\OAIyPMS4IzDz+uW5YvlfQXJ2FdaE="
Kill "C:\Documents and Settings\" & Text2.Text & "\Local Settings\Temp\MessengerCache\SQ12FsDiU5FqY5a0q80MBMi0E3vE="
Kill "C:\Documents and Settings\" & Text2.Text & "\Local Settings\Temp\MessengerCache\uNDmL7p3U1us22eWEb9BsE0mjsc="
Kill "C:\Documents and Settings\" & Text2.Text & "\Local Settings\Temp\MessengerCache\X42Fs233V2IJ9TjW8Ub12FOWozTY8="
Kill "C:\Documents and Settings\" & Text2.Text & "\Local Settings\Temp\pyl6696.tmp.exe"
Kill "C:\Documents and Settings\" & Text2.Text & "\Local Settings\Temp\u3r5597.tmp"
Kill "C:\Documents and Settings\" & Text2.Text & "\Local Settings\Temp\v7r5599.tmp"
Kill "C:\Documents and Settings\" & Text2.Text & "\Local Settings\Temp\was559B.tmp"
Kill "C:\Documents and Settings\" & Text2.Text & "\Local Settings\Temp\wubi-10.04.1-rev190.log"
Kill "C:\Documents and Settings\" & Text2.Text & "\Local Settings\Temp\~DFEA22.tmp"
Kill "C:\WINDOWS\system32\wbem\Logs\wbemcore.log"
Kill "C:\WINDOWS\system32\wbem\Logs\wbemess.log"
Kill "C:\WINDOWS\system32\wbem\Logs\wmiprov.log"
Kill "C:\WINDOWS\system32\wbem\Logs\wbemess.lo_"
Kill "C:\WINDOWS\0.log"
Kill "C:\WINDOWS\setupact.log"
Kill "C:\WINDOWS\setupapi.log"
Kill "C:\WINDOWS\setuperr.log"
Kill "C:\WINDOWS\Debug\UserMode\userenv.log"

       List1.AddItem "Temp Dosyalarý Temizlendi."
End If
End If
End If
'Özel Ek Dosya ve Klasör Temizleme
If ProgressBar1 = 95 Then
If CheckBox18.Value = 1 Then
Timer3.Enabled = False
        On Error Resume Next
        Kill sýra1.Text
        Kill sýra2.Text
        Kill sýra3.Text
        Kill sýra4.Text
        Kill sýra5.Text
        List1.AddItem "Özel Ek Dosya ve Klasör Temizleme."
End If
End If
'================================================================================================

If ProgressBar1 = 100 Then
ProgressBar1.Value = 0
Timer1.Enabled = False
Timer1.Interval = 0
If CheckBox19.Value = 1 Then
Timer3.Interval = 1000
Else
Timer3.Interval = 0
End If
If CheckBox21.Value = 1 Then
Timer4.Enabled = True
Else
Timer4.Enabled = False
Timer4.Interval = 0
End If
MsgBox "Ýþlem Tamamlandý Þeçilen Ögeler Silinmiþtir.!", 48, "Bitti ;"
Command5.Enabled = True
Command6.Enabled = False
End If
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
ProgressBar2.Value = ProgressBar2.Value + 1
Label2.Caption = "%" & ProgressBar2.Value + 0
If ProgressBar2 = 97 Then
'--------------------------------------------
If CheckBox7.Value = 1 Then '1.
'Pc Hýzlý Kapatma
Set reg = CreateObject("wscript.shell")
reg.regwrite "HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Control\WaitToKillServiceTimeout", "1000"
End If
'--------------------------------------------
If CheckBox10.Value = 1 Then '2.
'Daha Fazla Ýþlemci Gücü
Shell "Rundll32.exe advapi32.dllProcessIdleTasks"
End If
'--------------------------------------------
If CheckBox9.Value = 1 Then '3.
'3. Pc Hýzlý Açma
Set reg = CreateObject("wscript.shell")
reg.regwrite "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management\PrefetchParameters\EnablePrefetcher", "6"
End If
'--------------------------------------------
If CheckBox15.Value = 1 Then '4.
'Windows Dahada Hýzlansýn
Set reg = CreateObject("wscript.shell")
reg.regwrite "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\PriorityControl\IRQ8Priority", "1", "REG_DWORD"
Set reg = CreateObject("wscript.shell")
reg.regwrite "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\FileSystem\NtfsDisableLastAccessUpdate", "1"
End If
'--------------------------------------------
If CheckBox14.Value = 1 Then '5.
'Bellek Performansý Arttýrma
Set reg = CreateObject("wscript.shell")
reg.regwrite "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management\LargeSystemCache", "1"
End If
'--------------------------------------------
If CheckBox12.Value = 1 Then '6.
'Menülerin Hýzlandýrýlmasý
Set reg = CreateObject("wscript.shell")
reg.regwrite "HKEY_CURRENT_USER\Control Panel\Desktop\MenuShowDelay", "0"
End If
'--------------------------------------------
If CheckBox13.Value = 1 Then '7.
'Görsel Efektlerin Azaltýlmasý
Set reg = CreateObject("wscript.shell")
reg.regwrite "HKEY_CURRENT_USER\Control Panel\Desktop\WindowMetrics\MinAnimate", "0"
End If
'--------------------------------------------
If CheckBox20.Value = 1 Then '8.
'Bip'leri Kaldýrma
Set reg = CreateObject("wscript.shell")
reg.regwrite "HKEY_CURRENT_USER\Control Panel\Sound\Beep", "no"
End If
'--------------------------------------------
If CheckBox11.Value = 1 Then '9.
'%20 Ýnternet Artýþý

End If
'--------------------------------------------
If CheckBox8.Value = 1 Then '10.
'Konuþma Balonlarý
Set reg = CreateObject("wscript.shell")
reg.regwrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advance\ShowInfoTip", "0"
Set reg = CreateObject("wscript.shell")
reg.regwrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advance\EnableBallonTips", "0", "REG_DWORD"
End If
'--------------------------------------------
If CheckBox16.Value = 1 Then '11.
'Yanýt Vermeyen Uygulamalar
Set reg = CreateObject("wscript.shell")
reg.regwrite "HKEY_CURRENT_USER\Control Panel\Desktop\HungAppTimeout", "500"
End If
End If
'-----------------------------------------
If ProgressBar2 = 100 Then
ProgressBar2.Value = 0
Timer2.Enabled = False
Timer2.Interval = 0
Label2.Caption = ""
MsgBox "Sistem Ýyileþtirme Tamamlanmýþtýr ve Seçilen Ögeler Uygulanmýþtýr.", 48, "Bitti ;"
Kapat = MsgBox("Sistem Ýyileþtirme Tamamlandý.Uygulanmasý için Bilgisayar Yeniden Baþlatýlmasý Gerekmektedir.! ", vbInformation + vbYesNo, "Kapatma")
If Kapat = vbYes Then
Shell ("shutdown -r -t 5")
End If
End If
End Sub

Private Sub Timer3_Timer()
On Error Resume Next
'sýra1
If ListView2.ListItems.Count = 0 Then Exit Sub
X = ListView2.SelectedItem.Index
sýra1.Text = ListView2.ListItems(X).ListSubItems(0).Text

'sýra2
If ListView2.ListItems.Count = 1 Then Exit Sub
X = ListView2.SelectedItem.Index
sýra2.Text = ListView2.ListItems(X).ListSubItems(1).Text

'sýra3
If ListView2.ListItems.Count = 2 Then Exit Sub
X = ListView2.SelectedItem.Index
sýra3.Text = ListView2.ListItems(X).ListSubItems(2).Text

'sýra4
If ListView2.ListItems.Count = 3 Then Exit Sub
X = ListView2.SelectedItem.Index
sýra4.Text = ListView2.ListItems(X).ListSubItems(1).Text

'sýra5
If ListView2.ListItems.Count = 4 Then Exit Sub
X = ListView2.SelectedItem.Index
sýra5.Text = ListView2.ListItems(X).ListSubItems(1).Text

Timer3.Enabled = False
End Sub
