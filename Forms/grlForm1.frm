VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{04759352-AB08-11D5-863E-0010B563AF37}#1.0#0"; "Cgridv11.ocx"
Begin VB.Form Form1 
   Caption         =   "Hedhja e notave"
   ClientHeight    =   9840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9840
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin MSComCtl2.FlatScrollBar FlatScrollBar2 
      Height          =   4695
      Left            =   15000
      TabIndex        =   44
      Top             =   2760
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   8281
      _Version        =   393216
      Appearance      =   0
      Orientation     =   1245184
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   370
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "NR  EMRI   MBIEMRI"
      Top             =   2760
      Visible         =   0   'False
      Width           =   2415
   End
   Begin Cgridv11.Cgrid Cgrid2 
      Height          =   375
      Left            =   2400
      TabIndex        =   13
      Top             =   2760
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      GridMode        =   0
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridBackColor   =   -2147483644
      GridForeColor   =   16777215
      GridForeColor   =   16777215
      GridBackColor   =   -2147483644
      RowHeight       =   16
      FixedColumnVisible=   -1  'True
      FixedRowVisible =   -1  'True
      ScrollBarV      =   0   'False
      ScrollBarh      =   0   'False
      Appearance3D    =   0   'False
      CellEditColor   =   0
   End
   Begin VB.Frame fraInfo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   12720
      TabIndex        =   20
      Top             =   0
      Width           =   2655
      Begin VB.CommandButton cmdWebsiste 
         BackColor       =   &H80000009&
         Caption         =   "Website"
         Height          =   375
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton cmdEmail 
         BackColor       =   &H80000009&
         Caption         =   "E-mail"
         Height          =   375
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   960
         Width           =   1095
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   120
         Stretch         =   -1  'True
         Top             =   600
         Width           =   960
      End
      Begin VB.Label lblWebsite 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   24
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label lblShkolla 
         Appearance      =   0  'Flat
         Caption         =   "Shkolla jo publike Nr 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   120
         TabIndex        =   23
         Top             =   120
         Width           =   2415
      End
   End
   Begin Cgridv11.Cgrid provimet 
      Height          =   1575
      Left            =   0
      TabIndex        =   16
      Top             =   3240
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   2778
      GridMode        =   0
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RowHeight       =   16
      FixedColumnVisible=   -1  'True
      FixedRowVisible =   -1  'True
      ScrollBarV      =   -1  'True
      ScrollBarh      =   0   'False
      Appearance3D    =   0   'False
      CellEditColor   =   16777088
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   -360
      TabIndex        =   5
      Top             =   0
      Width           =   12855
      Begin VB.Frame fraProvimet 
         Caption         =   "Provimet :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   1575
         Left            =   10080
         TabIndex        =   32
         Top             =   240
         Width           =   2655
         Begin VB.OptionButton optMature 
            Caption         =   "Provimi i matures"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   495
            Left            =   120
            TabIndex        =   40
            Top             =   840
            Width           =   2055
         End
         Begin VB.OptionButton optLirimi 
            Caption         =   "Provimi i lirimit"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   495
            Left            =   120
            TabIndex        =   39
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Frame fraTipiNota 
         Caption         =   "Lloji i notes :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   1575
         Left            =   5880
         TabIndex        =   31
         Top             =   240
         Width           =   4095
         Begin VB.OptionButton optSemestrale 
            Caption         =   "Semestrale"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   375
            Left            =   120
            TabIndex        =   41
            Top             =   1080
            Width           =   1695
         End
         Begin VB.OptionButton optMungesePa 
            Caption         =   "Mungese pa arsye"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   1920
            TabIndex        =   38
            Top             =   640
            Width           =   2055
         End
         Begin VB.OptionButton optDetyreKontrolli 
            Caption         =   "Detyre kontrolli"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   615
            Left            =   120
            TabIndex        =   37
            Top             =   540
            Width           =   1695
         End
         Begin VB.OptionButton optMungeseMe 
            Caption         =   "Mungese me arsye"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1920
            TabIndex        =   36
            Top             =   240
            Width           =   2055
         End
         Begin VB.OptionButton optMomentale 
            Caption         =   "Momentale"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   35
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame fraSemestri 
         Caption         =   "Semestri :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   1575
         Left            =   3480
         TabIndex        =   30
         Top             =   240
         Width           =   2295
         Begin VB.OptionButton optRiprovim 
            Caption         =   "Riprovim"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   1200
            Width           =   1095
         End
         Begin VB.OptionButton optVjetore 
            Caption         =   "Vjetore"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   375
            Left            =   120
            TabIndex        =   42
            Top             =   840
            Width           =   2055
         End
         Begin VB.OptionButton optSemestri2 
            Caption         =   "Semestri II"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   600
            Width           =   1455
         End
         Begin VB.OptionButton optSemestri1 
            Caption         =   "Semestri I"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   375
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.TextBox txtMuaji 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   29
         Top             =   170
         Width           =   600
      End
      Begin VB.TextBox txtDitet 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   450
         Locked          =   -1  'True
         TabIndex        =   28
         Text            =   "Hen   Mar  Mer   Enjt   Pre  Shtu  Diel"
         Top             =   500
         Width           =   2745
      End
      Begin VB.ComboBox cboMuaji 
         Height          =   315
         Left            =   900
         TabIndex        =   26
         Top             =   180
         Width           =   1455
      End
      Begin VB.CommandButton cmdDataSot 
         BackColor       =   &H80000009&
         Caption         =   "Data sot"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   3120
         Width           =   1935
      End
      Begin VB.ComboBox cboVitiShkollor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         ForeColor       =   &H80000004&
         Height          =   360
         Left            =   7920
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2040
         Width           =   1455
      End
      Begin VB.ComboBox cboIndeksi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         ForeColor       =   &H80000004&
         Height          =   360
         Left            =   5880
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2040
         Width           =   735
      End
      Begin VB.ComboBox cboKlasa 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         ForeColor       =   &H80000004&
         Height          =   360
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2040
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         DisabledPicture =   "grlForm1.frx":0000
         DownPicture     =   "grlForm1.frx":6442
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9600
         Picture         =   "grlForm1.frx":C884
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2040
         Width           =   1695
      End
      Begin MSACAL.Calendar Calendar1 
         CausesValidation=   0   'False
         Height          =   2415
         Left            =   360
         TabIndex        =   27
         Top             =   120
         Width           =   2895
         _Version        =   524288
         _ExtentX        =   5106
         _ExtentY        =   4260
         _StockProps     =   1
         BackColor       =   16777152
         Year            =   2005
         Month           =   3
         Day             =   1
         DayLength       =   1
         MonthLength     =   1
         DayFontColor    =   0
         FirstDay        =   1
         GridCellEffect  =   1
         GridFontColor   =   10485760
         GridLinesColor  =   -2147483632
         ShowDateSelectors=   -1  'True
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   -1  'True
         ShowTitle       =   -1  'True
         ShowVerticalGrid=   -1  'True
         TitleFontColor  =   10485760
         ValueIsNull     =   0   'False
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblVitiShkollor 
         Caption         =   "Viti Shkollor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6960
         TabIndex        =   12
         Top             =   2070
         Width           =   855
      End
      Begin VB.Label lblIndeksi 
         Caption         =   "Indeksi"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         TabIndex        =   11
         Top             =   2070
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Klasa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   10
         Top             =   2070
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdDalje 
      BackColor       =   &H00FFFFFF&
      DisabledPicture =   "grlForm1.frx":12CC6
      DownPicture     =   "grlForm1.frx":19108
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      Picture         =   "grlForm1.frx":1F54A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9120
      Width           =   2535
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00FFFFFF&
      DisabledPicture =   "grlForm1.frx":2598C
      DownPicture     =   "grlForm1.frx":2BDCE
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      Picture         =   "grlForm1.frx":32210
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9120
      Width           =   2535
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   7440
      Visible         =   0   'False
      Width           =   15250
      _ExtentX        =   26908
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Arrows          =   65536
      Orientation     =   1245185
      Value           =   1
   End
   Begin Cgridv11.Cgrid Amza 
      Height          =   945
      Left            =   240
      TabIndex        =   15
      Top             =   5040
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1667
      GridMode        =   0
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RowHeight       =   16
      FixedColumnVisible=   -1  'True
      FixedRowVisible =   -1  'True
      ScrollBarV      =   0   'False
      ScrollBarh      =   0   'False
      Appearance3D    =   0   'False
      CellEditColor   =   0
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   0
      TabIndex        =   43
      Top             =   7800
      Width           =   17650
      Begin VB.CommandButton cmdHelp 
         BackColor       =   &H00FFFFFF&
         DisabledPicture =   "grlForm1.frx":38652
         DownPicture     =   "grlForm1.frx":3BB8C
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11400
         Picture         =   "grlForm1.frx":3F0C6
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   1320
         Width           =   2535
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   2775
      Left            =   0
      TabIndex        =   45
      Top             =   0
      Width           =   15555
   End
   Begin Cgridv11.Cgrid Emrat 
      Height          =   945
      Left            =   0
      TabIndex        =   1
      Top             =   3120
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1667
      GridMode        =   0
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RowHeight       =   16
      FixedColumnVisible=   -1  'True
      FixedRowVisible =   -1  'True
      ScrollBarV      =   0   'False
      ScrollBarh      =   0   'False
      Appearance3D    =   0   'False
      CellEditColor   =   0
   End
   Begin Cgridv11.Cgrid Cgrid1 
      Height          =   945
      Left            =   2400
      TabIndex        =   14
      Top             =   3120
      Visible         =   0   'False
      Width           =   19095
      _ExtentX        =   33681
      _ExtentY        =   1667
      GridMode        =   0
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RowHeight       =   16
      FixedColumnVisible=   -1  'True
      FixedRowVisible =   -1  'True
      ScrollBarV      =   0   'False
      ScrollBarh      =   0   'False
      Appearance3D    =   0   'False
      CellEditColor   =   16777088
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Provimet e"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   19
      Top             =   2760
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "lirimit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   18
      Top             =   2760
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "matures"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   17
      Top             =   2760
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public col As Integer
Dim objektGabimi As New clsErrorHandler
Public w As Integer
Dim tipiNota As Integer
Dim rowCount As Integer
Dim colCount As Integer
Dim dataVjeter As Date
Dim klasaVjeter As String
Dim indeksiVjeter As String
Dim llojNoteVjeter As String
Dim semesterVjeter As String
Dim vitiShkollorVjeter As String
Dim oldRow As Integer
Dim oldRowProvimet As Integer


Private Sub Calendar1_BeforeUpdate(Cancel As Integer)
    Dim dt As Date
    dt = Calendar1.Value
    HidhNeTabele
    dataVjeter = DataTekst(Calendar1.Value)
End Sub

Private Sub cboIndeksi_Change()
    HidhNeTabele
End Sub

Private Sub cboIndeksi_Click()
    Pastro
    HidhNeTabele
    indeksiVjeter = cboIndeksi.Text
End Sub

Private Sub cboKlasa_Click()
    Pastro
    HidhNeTabele
    klasaVjeter = cboKlasa.Text
End Sub



Private Sub cboVitiShkollor_Click()
    Pastro
    If Me.optLirimi.Value Then
        Dim vitFillimi As Integer
        vitFillimi = CInt(Mid(Me.cboVitiShkollor.Text, 1, 4))
        If (vitFillimi <= 2006) Then
            cboKlasa.Text = "8"
        Else
            cboKlasa.Text = "9"
        End If
    End If
End Sub

Private Sub Cgrid1_CellClick(row As Long, col As Long)
    If oldRow Mod 2 = 0 Then
        Cgrid1.RowBackColor(oldRow) = vbWhite
        emrat.RowBackColor(oldRow) = vbWhite
    Else
        Cgrid1.RowBackColor(oldRow) = &HE0E0E0
        emrat.RowBackColor(oldRow) = &HE0E0E0
    End If
    Cgrid1.RowBackColor(row) = &H80000013
    emrat.RowBackColor(row) = &H80000003
    oldRow = row
End Sub

Private Sub Cgrid1_CellGotFocus(row As Long, col As Long)
    'Cgrid1.CellFontBold(row, col) = True
    Cgrid1.CellEditColor = &HFFFFFF
    Select Case tipiNota
        Case 1
           Cgrid1.CellForeColor(row, col) = &H80000008
        Case 2
            Cgrid1.CellForeColor(row, col) = &HC00000
        Case 3
            Cgrid1.CellForeColor(row, col) = &HFF&
        Case 4
            Cgrid1.CellForeColor(row, col) = &H8000&
        Case 5
             Cgrid1.CellForeColor(row, col) = &HC000C0
    End Select
End Sub


Private Sub cmdEmail_Click()
    If Not OpenEmailProgram(email) Then
        MsgBox "Ju nuk keni asnje program te instaluar per te derguar email", vbExclamation, "Gabim ne dergimin e emailit"
    End If
End Sub

Private Sub cmdHelp_Click()
CallHelp indeksHelp
End Sub

Private Sub cmdOK_Click()
   Dim data As String
   data = date
   objektGabimi.kapGabimin
   objektGabimi.menazhim_gabimi
   HidhNeTabele
   If objektGabimi.mvarGabimi = 0 Then
      If Cgrid1.Visible = True Or provimet.Visible = True Then
         If optMature.Value Or optLirimi.Value Then

            objectInitialization HEDHJA_PROVIME

         ElseIf optMomentale.Value Or optDetyreKontrolli.Value Then

            objectInitialization HEDHJA_NOTAVE_MOMENTALE

         ElseIf optSemestrale.Value Or optVjetore.Value Or optRiprovim.Value Then

            objectInitialization HEDHJA_NOTAVE_MOMENTALE 'HEDHJA_NOTAVE_PERFUNDIMTARE

         ElseIf optMungesePa.Value Or optMungeseMe.Value Then

            objectInitialization HEDHJA_NOTAVE_MOMENTALE 'HEDHJA_E_MUNGESAVE
         Else
         End If
         If data_modifikimi = "" Or data <> data_modifikimi Then
            objectInitialization MODIFIKIMI_I_DATES
         End If
         optMomentale.Enabled = True
         optLirimi.Enabled = True
         optMature.Enabled = True
         optDetyreKontrolli.Enabled = True
         optSemestrale.Enabled = True
         optMungeseMe.Enabled = True
         optMungesePa.Enabled = True
         optSemestri1.Enabled = True
         optSemestri2.Enabled = True
         optVjetore.Enabled = True
         optRiprovim.Enabled = True
         'cboKlasa.Enabled = True
         'cboIndeksi.Enabled = True
         'cboVitiShkollor.Enabled = True

      End If
   End If
   PastroTabele
End Sub

Private Sub Calendar1_NewMonth()
     perkthe_muajin
End Sub

Private Sub cboMuaji_Change()
    cboMuaji.Locked = True
    sinkronizo_muajt
End Sub

Private Sub cboMuaji_Click()
    Calendar1.Month = cboMuaji.ListIndex + 1
End Sub

Private Sub cboMuaji_DropDown()
    cboMuaji.Locked = False
End Sub

Private Sub cmdDataSot_Click()
    Calendar1.Today
End Sub
Private Sub cmdDalje_Click()
Unload Me
  Set active_form = Nothing
End Sub

Private Sub cmdWebsiste_Click()
    If website <> "" Then
        GoToWeb website
    Else
        MsgBox "Ju nuk e keni dhene adresen e faqes tuaj te web-it.", vbInformation
    End If
End Sub

Private Sub Command1_Click()
    Command1.Enabled = False
   Pastro
   objektGabimi.kapGabimin
   objektGabimi.menazhim_gabimi
   If objektGabimi.mvarGabimi = 0 Then
      If Not (optLirimi.Value) Or Not (optMature.Value) Then
        optLirimi.Enabled = False
        optMature.Enabled = False
      Else
          optMomentale.Enabled = False
          optLirimi.Enabled = False
          optMature.Enabled = False
          optDetyreKontrolli.Enabled = False
          optSemestrale.Enabled = False
          optMungeseMe.Enabled = False
          optMungesePa.Enabled = False
          optSemestri2.Enabled = False
          optSemestri1.Enabled = False
          optVjetore.Enabled = False
          optRiprovim.Enabled = False
      End If
    
     ' cboKlasa.Enabled = False
     ' cboIndeksi.Enabled = False
     ' cboVitiShkollor.Enabled = False
      
      If optLirimi.Value Or optMature.Value Then
           objectInitialization VISUALIZO_KLASA_PROVIMET
      Else
         objectInitialization HEDHJE_NOTASH_INFO_KLASA
      End If
      If scrbar = True Then
         FlatScrollBar1.Visible = True
      Else
         FlatScrollBar1.Visible = False
      End If
      If scrbarver = True Then
         FlatScrollBar2.Visible = True
      Else
         FlatScrollBar2.Visible = False
      End If
   Else
      FlatScrollBar1.Visible = False
      FlatScrollBar2.Visible = False
   End If
   colCount = Cgrid1.Width / 1200
   rowCount = Cgrid1.Height / 255
   ReDim vektor(rowCount, colCount)
   For i = 1 To rowCount
    For j = 1 To colCount
        'vektor(I, j).date = ""
        vektor(i, j).emer = ""
        vektor(i, j).indeksi = ""
        'vektor(I, j).Klasa = ""
        vektor(i, j).lende = ""
        vektor(i, j).llojNote = ""
        vektor(i, j).note = ""
        vektor(i, j).semester = ""
    Next
   Next
   
   Command1.Enabled = True
End Sub

Private Sub Emrat_CellClick(row As Long, col As Long)
    If oldRow Mod 2 = 0 Then
        Cgrid1.RowBackColor(oldRow) = vbWhite
        emrat.RowBackColor(oldRow) = vbWhite
    Else
        Cgrid1.RowBackColor(oldRow) = &HE0E0E0
        emrat.RowBackColor(oldRow) = &HE0E0E0
    End If
    Cgrid1.RowBackColor(row) = &H80000013
    emrat.RowBackColor(row) = &H80000003
    oldRow = row
    Cgrid1.RowBackColor(row) = &H80000013
    emrat.RowBackColor(row) = &H80000003
End Sub


Private Sub FlatScrollBar1_Change()

   Dim a                As Integer
   l = 1200

   Cgrid1.left = 2 * l - FlatScrollBar1.Value
   Cgrid2.left = 2 * l - FlatScrollBar1.Value '0



End Sub

Private Sub FlatScrollBar1_Scroll()
   Dim l                As Integer
   l = 1200
   'Emri.Left = -FlatScrollBar1.Value
   'Emrat.Left = -FlatScrollBar1.Value

   Cgrid1.left = 2 * l - FlatScrollBar1.Value
   Cgrid2.left = 2 * l - FlatScrollBar1.Value



End Sub

Private Sub FlatScrollBar2_Change()
    l = 3120
    Cgrid1.top = l - FlatScrollBar2.Value
    emrat.top = l - FlatScrollBar2.Value

End Sub

Private Sub Form_Load()
   mbushKomboBox
   viti
   Image1.Picture = LoadPicture(adresaLogo)
   lblShkolla.Caption = emerShkolla
   Me.FlatScrollBar1.Visible = False
   Me.FlatScrollBar2.Visible = False
   
   'Call shfaq_grile(11, 10)
   Calendar1.Today
   perkthe_muajin
   mbush_kombo_muaji
   cboMuaji.Locked = True
   If scrbar And Me.Cgrid1.Visible = True Then
      FlatScrollBar1.Visible = True
   End If
   percaktoTeDrejtat
   cmdOK.Enabled = False
   colCount = 0
   rowCount = 0
   oldRow = 0
   oldRowProvimet = 0
End Sub




Private Sub optLirimi_Click()
      
   'Cgrid1.CellEditColor = &HFF00&
   
   optMomentale.Value = False
   optDetyreKontrolli.Value = False
   optMungeseMe.Value = False
   optMungesePa.Value = False
   optSemestrale.Value = False
   optSemestri1.Value = False
   optSemestri2.Value = False
   optVjetore.Value = False
   optRiprovim.Value = False
   If optLirimi.Value = True Then
        optDetyreKontrolli.Enabled = False
        optMomentale.Enabled = False
        optSemestrale.Enabled = False
        optMungeseMe.Enabled = False
        optMungesePa.Enabled = False
   End If
   Dim vitFillimi As Integer
   vitFillimi = CInt(Mid(Me.cboVitiShkollor.Text, 1, 4))
   If (vitFillimi <= 2006) Then
        cboKlasa.Text = "8"
    Else
        cboKlasa.Text = "9"
    End If
   tipiNota = 5
End Sub

Private Sub optMature_Click()

   'Cgrid1.CellEditColor = &HFF00&
   optMomentale.Value = False
   optDetyreKontrolli.Value = False
   optMungeseMe.Value = False
   optMungesePa.Value = False
   optSemestrale.Value = False
   optRiprovim.Value = False
   optSemestri1.Value = False
   optSemestri2.Value = False
   optVjetore.Value = False
   If optMature.Value = True Then
        optDetyreKontrolli.Enabled = False
        optMomentale.Enabled = False
        optSemestrale.Enabled = False
        optMungeseMe.Enabled = False
        optMungesePa.Enabled = False
   End If
   cboKlasa.Text = "12"
   tipiNota = 5
   
End Sub

Private Sub optDetyreKontrolli_Click()

    optLirimi.Value = False
    optMature.Value = False
    optVjetore.Value = False
    If optSemestri1.Value = False And optSemestri2.Value = False Then
        MsgBox "Ju duhet te percaktoni semestrin para se te zgjidhni llojin e notes.", vbExclamation, "Hedhja e notave."
        optDetyreKontrolli.Value = False
        Exit Sub
    End If
    llojNoteVjeter = "D"
   tipiNota = 4
End Sub

Private Sub optMomentaleSI_Click()
   'Cgrid1.CellEditColor = &HFFFF80
   optLirimi.Value = False
   optMature.Value = False
   tipiNota = 1
End Sub

Private Sub optMomentaleSII_Click()
   'Cgrid1.CellEditColor = &HFFFF80
   optLirimi.Value = False
   optMature.Value = False
   tipiNota = 1
End Sub


Private Sub optMungese_me_Click()
   optLirimi.Value = False
   optMature.Value = False
End Sub

Private Sub optMungese_pa_Click()
    optLirimi.Value = False
    optMature.Value = False
End Sub

Private Sub optSemestraleII_Click()
   'Cgrid1.CellEditColor = vbBlue
   optLirimi.Value = False
   optMature.Value = False
   tipiNota = 2
End Sub

Private Sub optSemestriI_Click()
   'Cgrid1.CellEditColor = vbBlue
   optLirimi.Value = False
   optMature.Value = False
   tipiNota = 2
End Sub


Private Sub optMomentale_Click()
    optLirimi.Value = False
    optMature.Value = False
    optVjetore.Value = False
    HidhNeTabele
    llojNoteVjeter = "M"
    If optSemestri1.Value = False And optSemestri2.Value = False Then
        MsgBox "Ju duhet te percaktoni semestrin para se te zgjidhni llojin e notes.", vbExclamation, "Hedhja e notave."
        optMomentale.Value = False
        Exit Sub
    End If
    tipiNota = 1
End Sub

Private Sub optMungeseMe_Click()
    optLirimi.Value = False
    optMature.Value = False
    optVjetore.Value = False
    HidhNeTabele
    llojNoteVjeter = "Mm"
    If optSemestri1.Value = False And optSemestri2.Value = False Then
        MsgBox "Ju duhet te percaktoni semestrin para se te zgjidhni llojin e notes.", vbExclamation, "Hedhja e notave."
        optMungeseMe.Value = False
    End If
    tipiNota = 1
End Sub

Private Sub optMungesePa_Click()
    optLirimi.Value = False
    optMature.Value = False
    optVjetore.Value = False
    HidhNeTabele
    llojNoteVjeter = "Mp"
    If optSemestri1.Value = False And optSemestri2.Value = False Then
        MsgBox "Ju duhet te percaktoni semestrin para se te zgjidhni llojin e notes.", vbExclamation, "Hedhja e notave."
        optMungesePa.Value = False
    End If
    tipiNota = 1
End Sub

Private Sub optRiprovim_Click()
    
   optLirimi.Value = False
   optMature.Value = False
   optMomentale.Value = False
   optDetyreKontrolli.Value = False
   optMungeseMe.Value = False
   optMungesePa.Value = False
   optSemestrale.Value = False
   If optRiprovim.Value = True Then
        optDetyreKontrolli.Enabled = False
        optMomentale.Enabled = False
        optSemestrale.Enabled = False
        optMungeseMe.Enabled = False
        optMungesePa.Enabled = False
   End If
   tipiNota = 5
   HidhNeTabele
   semesterVjeter = "R"
   
End Sub

Private Sub optSemestrale_Click()
    optLirimi.Value = False
    optMature.Value = False
    optVjetore.Value = False
    If optSemestri1.Value = False And optSemestri2.Value = False Then
        MsgBox "Ju duhet te percaktoni semestrin para se te zgjidhni llojin e notes.", vbExclamation, "Hedhja e notave."
        optSemestrale.Value = False
        Exit Sub
    End If
    HidhNeTabele
    llojNoteVjeter = "S"
    tipiNota = 2
End Sub

Private Sub optSemestri1_Click()
    optLirimi.Value = False
    optMature.Value = False
    If optSemestri1.Value = True Then
        optDetyreKontrolli.Enabled = True
        optMomentale.Enabled = True
        optSemestrale.Enabled = True
        optMungeseMe.Enabled = True
        optMungesePa.Enabled = True
    End If
    HidhNeTabele
    semesterVjeter = "S1"
End Sub

Private Sub optSemestri2_Click()
    optLirimi.Value = False
    optMature.Value = False
    If optSemestri2.Value = True Then
        optDetyreKontrolli.Enabled = True
        optMomentale.Enabled = True
        optSemestrale.Enabled = True
        optMungeseMe.Enabled = True
        optMungesePa.Enabled = True
    End If
    HidhNeTabele
    semesterVjeter = "S2"
End Sub

Private Sub optVjetore_Click()
   'Cgrid1.CellEditColor = &HFF&
   optLirimi.Value = False
   optMature.Value = False
   optMomentale.Value = False
   optDetyreKontrolli.Value = False
   optMungeseMe.Value = False
   optMungesePa.Value = False
   optSemestrale.Value = False
   If optVjetore.Value = True Then
        optDetyreKontrolli.Enabled = False
        optMomentale.Enabled = False
        optSemestrale.Enabled = False
        optMungeseMe.Enabled = False
        optMungesePa.Enabled = False
   End If
   HidhNeTabele
   semesterVjeter = "V"
   tipiNota = 3
End Sub

Private Sub objectInitialization(actionName As GUI_ACTION__ENUM)
   Set objUIController = New clsUIController

   objUIController.actionName = actionName
   objUIController.ExecuteActions

   Set objUIController = Nothing
End Sub

Private Sub mbushKomboBox()
    
    cboVitiShkollor.AddItem "1990-1991"
    cboVitiShkollor.AddItem "1991-1992"
    cboVitiShkollor.AddItem "1992-1993"
    cboVitiShkollor.AddItem "1993-1994"
    cboVitiShkollor.AddItem "1994-1995"
    cboVitiShkollor.AddItem "1995-1996"
    cboVitiShkollor.AddItem "1996-1997"
    cboVitiShkollor.AddItem "1997-1998"
    cboVitiShkollor.AddItem "1998-1999"
    cboVitiShkollor.AddItem "1999-2000"
    cboVitiShkollor.AddItem "2000-2001"
    cboVitiShkollor.AddItem "2001-2002"
    cboVitiShkollor.AddItem "2002-2003"
    cboVitiShkollor.AddItem "2003-2004"
    cboVitiShkollor.AddItem "2004-2005"
    cboVitiShkollor.AddItem "2005-2006"
    cboVitiShkollor.AddItem "2006-2007"
    cboVitiShkollor.AddItem "2007-2008"
    cboVitiShkollor.AddItem "2008-2009"
    cboVitiShkollor.AddItem "2009-2010"
    cboVitiShkollor.AddItem "2010-2011"
    cboVitiShkollor.AddItem "2011-2012"
    cboVitiShkollor.AddItem "2012-2012"
    cboVitiShkollor.AddItem "2013-2014"
    cboVitiShkollor.AddItem "2014-2015"
    cboVitiShkollor.AddItem "2015-2016"
    cboVitiShkollor.AddItem "2016-2017"
    cboVitiShkollor.AddItem "2017-2018"
    cboVitiShkollor.AddItem "2018-2019"
    cboVitiShkollor.AddItem "2019-2020"
    
    
   
    
    cboKlasa.AddItem "1"
    cboKlasa.AddItem "2"
    cboKlasa.AddItem "3"
    cboKlasa.AddItem "4"
    cboKlasa.AddItem "5"
    cboKlasa.AddItem "6"
    cboKlasa.AddItem "7"
    cboKlasa.AddItem "8"
    cboKlasa.AddItem "9"
    cboKlasa.AddItem "10"
    cboKlasa.AddItem "11"
    cboKlasa.AddItem "12"
    
    
    
    cboIndeksi.AddItem "A"
    cboIndeksi.AddItem "B"
    cboIndeksi.AddItem "C"
    cboIndeksi.AddItem "D"
    cboIndeksi.AddItem "E"
    cboIndeksi.AddItem "F"
    
    
    
    
    
End Sub
Private Sub viti()
   Dim d, muai, viti    As String
   d = Now
   Dim M, v             As Integer
   muai = DateTime.Month(DateTime.Now)
   viti = DateTime.Year(DateTime.date)
   M = Val(muai)
   v = Val(viti)
   cboVitiShkollor.Text = gjej_vitin(M, v)

End Sub

Private Sub perkthe_muajin()
    
    Dim muaji As Integer
    muaji = Calendar1.Month
    Select Case muaji
        Case 1
            txtMuaji.Text = "Janar"
            cboMuaji.Text = "Janar"
        Case 2
            txtMuaji.Text = "Shkurt"
            cboMuaji.Text = "Shkurt"
        Case 3
            txtMuaji.Text = "Mars"
            cboMuaji.Text = "Mars"
        Case 4
            txtMuaji.Text = "Prill"
            cboMuaji.Text = "Prill"
        Case 5
            txtMuaji.Text = "Maj"
            cboMuaji.Text = "Maj"
        Case 6
            txtMuaji.Text = "Qershor"
            cboMuaji.Text = "Qershor"
        Case 7
            txtMuaji.Text = "Korrik"
            cboMuaji.Text = "Korrik"
        Case 8
            txtMuaji.Text = "Gusht"
            cboMuaji.Text = "Gusht"
        Case 9
            txtMuaji.Text = "Shtator"
            cboMuaji.Text = "Shtator"
        Case 10
            txtMuaji.Text = "Tetor"
            cboMuaji.Text = "Tetor"
        Case 11
            txtMuaji.Text = "Nentor"
            cboMuaji.Text = "Nentor"
        Case 12
            txtMuaji.Text = "Dhjetor"
            cboMuaji.Text = "Dhjetor"
    End Select
    txtMuaji.Text = ""
End Sub

Private Sub mbush_kombo_muaji()
    cboMuaji.AddItem "Janar"
    cboMuaji.AddItem "Shkurt"
    cboMuaji.AddItem "Mars"
    cboMuaji.AddItem "Prill"
    cboMuaji.AddItem "Maj"
    cboMuaji.AddItem "Qershor"
    cboMuaji.AddItem "Korrik"
    cboMuaji.AddItem "Gusht"
    cboMuaji.AddItem "Shtator"
    cboMuaji.AddItem "Tetor"
    cboMuaji.AddItem "Nentor"
    cboMuaji.AddItem "Dhjetor"
End Sub

Private Sub sinkronizo_muajt()
    Dim muaji As String
    muaji = cboMuaji.Text
    Select Case muaji
        Case "Janar"
            txtMuaji.Text = "Janar"
        Case "Shkurt"
            txtMuaji.Text = "Shkurt"
        Case "Mars"
            txtMuaji.Text = "Mars"
        Case "Prill"
            txtMuaji.Text = "Prill"
        Case "Maj"
            txtMuaji.Text = "Maj"
        Case "Qershor"
            txtMuaji.Text = "Qershor"
        Case "Gusht"
            txtMuaji.Text = "Gusht"
        Case "Shtator"
            txtMuaji.Text = "Shtator"
        Case "Tetor"
            txtMuaji.Text = "Tetor"
        Case "Nentor"
            txtMuaji.Text = "Nentor"
        Case "Dhjetor"
            txtMuaji.Text = "Dhjetor"
    End Select
    
End Sub

Private Sub provimet_CellClick(row As Long, col As Long)
   If row > 1 Then
    If oldRowProvimet Mod 2 = 0 Then
        provimet.RowBackColor(oldRowProvimet) = vbWhite
    Else
        provimet.RowBackColor(oldRowProvimet) = &HE0E0E0
    End If
    provimet.RowBackColor(row) = &H80000013
    oldRowProvimet = row
   End If

End Sub

Private Sub provimet_CellGotFocus(row As Long, col As Long)
   If row > 1 And col > 1 Then
      'provimet.CellFontBold(row, col) = True
      If tipiNota = 5 Then
         provimet.CellForeColor(row, col) = &H4080&
      End If
    If oldRowProvimet Mod 2 = 0 Then
        emrat.RowBackColor(oldRow) = vbWhite
        provimet.RowBackColor(oldRow) = vbWhite
    Else
        emrat.RowBackColor(oldRow) = &HE0E0E0
        provimet.RowBackColor(oldRow) = &HE0E0E0
    End If
   End If
End Sub

Private Sub percaktoTeDrejtat()
    
    Select Case statusi
        
        Case "SupervizorEmesme"
            mbushKomboMesme
            optLirimi.Visible = False
        Case "SupervizorTetevjecare"
            mbushKomboTetevjecare
            optMature.Visible = False
        Case Else
            
    End Select
    
            
            
End Sub

Private Sub mbushKomboMesme()
    cboKlasa.Clear
    cboKlasa.AddItem "9"
    cboKlasa.AddItem "10"
    cboKlasa.AddItem "11"
    cboKlasa.AddItem "12"
End Sub

Private Sub mbushKomboTetevjecare()
   cboKlasa.Clear
   cboKlasa.AddItem "1"
   cboKlasa.AddItem "2"
   cboKlasa.AddItem "3"
   cboKlasa.AddItem "4"
   cboKlasa.AddItem "5"
   cboKlasa.AddItem "6"
   cboKlasa.AddItem "7"
   cboKlasa.AddItem "8"
End Sub

Private Sub Pastro()
    Cgrid1.Visible = False
    emrat.Visible = False
    provimet.Visible = False
    Text1.Visible = False
    'Text2.Visible = False
    Cgrid2.Visible = False
    Amza.Visible = False
    FlatScrollBar1.Visible = False
    FlatScrollBar2.Visible = False
End Sub

Private Sub HidhNeTabele()
    If Not Cgrid1.Visible Then
        Exit Sub
    End If
    If rowCount = 0 Or colCount = 0 Then
        Exit Sub
    End If
    For i = 1 To rowCount
        For j = 1 To colCount
            If Cgrid1.Text(i, j) <> "" Then
                If vektor(i, j).emer = "" Then 'Or vektor(I, j) <> Null Then
                    vektor(i, j).note = Cgrid1.Text(i, j)
                    vektor(i, j).emer = Me.emrat.Text(i, 1)
                    vektor(i, j).lende = Cgrid2.Text(1, j)
                    vektor(i, j).date = DataTekst(dataVjeter)
                    vektor(i, j).indeksi = cboIndeksi.Text
                    vektor(i, j).Klasa = cboKlasa.Text
                    vektor(i, j).llojNote = llojNoteVjeter
                    vektor(i, j).semester = semesterVjeter
                End If
            End If
        Next
    Next
End Sub

Private Sub PastroTabele()
    For i = 1 To rowCount
        For j = 1 To colCount
            vektor(i, j).note = ""
            vektor(i, j).emer = ""
            vektor(i, j).lende = ""
            vektor(i, j).date = Empty
            vektor(i, j).indeksi = ""
            vektor(i, j).Klasa = Empty
            vektor(i, j).llojNote = ""
            vektor(i, j).semester = ""
            vektor(i, j).llojNote = ""
            
        Next
    Next
End Sub

Private Function DataTekst(dataDt As Date) As String
    
    Dim dataMbrapsht As String
    Dim i As String
    dataMbrapsht = ""
    Dim dita As String
    Dim muaji As String
    Dim viti As String
    dita = DateTime.Day(dataDt)
    muaji = DateTime.Month(dataDt)
    viti = DateTime.Year(dataDt)
    'viti = Mid(data, 7, 4)
    If (Len(dita) = 1) Then
        dita = "0" + dita
    End If
    
    If (Len(muaji) = 1) Then
        muaji = "0" + muaji
    End If
    dataMbrapsht = dita & "/" & muaji & "/" & viti
    DataTekst = dataMbrapsht
End Function


