VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmKonsultimeAmza 
   Caption         =   "Konsultime - Amza"
   ClientHeight    =   9840
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14280
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9840
   ScaleWidth      =   14280
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cboVitiShkollor 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4320
      Style           =   2  'Dropdown List
      TabIndex        =   41
      Top             =   6240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H80000009&
      Caption         =   "Printo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   9000
      Width           =   2055
   End
   Begin VB.Frame fraInfo 
      Height          =   1575
      Left            =   12480
      TabIndex        =   35
      Top             =   0
      Width           =   2655
      Begin VB.CommandButton cmdEmail 
         BackColor       =   &H80000009&
         Caption         =   "E-mail"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton cmdWebsiste 
         BackColor       =   &H80000009&
         Caption         =   "Website"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblShkolla 
         Appearance      =   0  'Flat
         Caption         =   "Shkolla jo publike Nr 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   120
         TabIndex        =   39
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label lblWebsite 
         Height          =   375
         Left            =   1560
         TabIndex        =   38
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   120
         Picture         =   "frmKonsultimeAmza.frx":0000
         Stretch         =   -1  'True
         Top             =   600
         Width           =   960
      End
   End
   Begin VB.OptionButton optUlet 
      Caption         =   "Shkolla tetevjecare"
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   240
      TabIndex        =   34
      Top             =   1080
      Width           =   1095
   End
   Begin VB.OptionButton optMesme 
      Caption         =   "Shkolla e mesme"
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   240
      TabIndex        =   33
      Top             =   240
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cikli"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   2175
      Left            =   120
      TabIndex        =   32
      Top             =   0
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid gridaPjekurie 
      Height          =   975
      Left            =   6840
      TabIndex        =   28
      Top             =   6600
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   1720
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      ScrollBars      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid gridaAmza 
      Height          =   2175
      Left            =   120
      TabIndex        =   26
      Top             =   2760
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   3836
      _Version        =   393216
      Rows            =   6
      Cols            =   6
      FixedCols       =   0
      FocusRect       =   0
      HighLight       =   2
      SelectionMode   =   1
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000009&
      Caption         =   "Ndihme"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   9000
      Width           =   2055
   End
   Begin VB.TextBox txtKlasa 
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
      Height          =   375
      Left            =   8640
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H80000009&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   9000
      Width           =   2055
   End
   Begin VB.Frame fraKerkim 
      Caption         =   "Te dhenat e amzes ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   2175
      Left            =   1560
      TabIndex        =   10
      Top             =   0
      Width           =   9615
      Begin VB.TextBox txtDatelindja 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox txtEmri 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   1
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtMbiemri 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   2
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtIndeksi 
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
         Height          =   375
         Left            =   8040
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtAmzaNo 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtAtesia 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   4
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtMemesia 
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
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox txtSeksi 
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
         Height          =   375
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox txtVendlindja 
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
         Height          =   375
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Indeksi"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8040
         TabIndex        =   24
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblEmri 
         Caption         =   "Emri"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   19
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblMbiemri 
         Caption         =   "Mbiemri"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   18
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblKlasa 
         Caption         =   "Klasa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7080
         TabIndex        =   17
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblAmzaNo 
         Caption         =   "Numri Amzes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblAtesia 
         Caption         =   "Atesia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         TabIndex        =   15
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblMemesia 
         Caption         =   "Memesia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lblSeksi 
         Caption         =   "Seksi"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   13
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label lblVendlindja 
         Caption         =   "Vendlindja"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4200
         TabIndex        =   12
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label lblDatelindja 
         Caption         =   "Datelindja"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6240
         TabIndex        =   11
         Top             =   1200
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdDil 
      BackColor       =   &H80000009&
      Caption         =   "Dil"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9000
      Width           =   2055
   End
   Begin RichTextLib.RichTextBox rtbShenime 
      Height          =   2055
      Left            =   120
      TabIndex        =   8
      Top             =   6600
      Width           =   5820
      _ExtentX        =   10266
      _ExtentY        =   3625
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"frmKonsultimeAmza.frx":6F75
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblVitiShkollor 
      Caption         =   "Viti Shkollor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   42
      Top             =   5880
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "matures"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9360
      TabIndex        =   31
      Top             =   6120
      Width           =   975
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "lirimit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9360
      TabIndex        =   30
      Top             =   6120
      Width           =   735
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Notat ne provimet e"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   29
      Top             =   6120
      Width           =   2535
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Perparimi sipas lendeve"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   27
      Top             =   2280
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "Indeksi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7440
      TabIndex        =   23
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Shenime per sjelljen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   6120
      Width           =   2895
   End
End
Attribute VB_Name = "frmKonsultimeAmza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim objektGabimi As New clsErrorHandler
Dim objUIController As clsUIController
Dim objBusManager As clsBusManager



Private Sub cboVitiShkollor_Click()
    rtbShenime.Text = ""
    objectInitialization AMZA_SJELLJE
End Sub

Private Sub cmdDil_Click()
  Unload Me
  Set active_form = Nothing
End Sub

Private Sub cmdOK_Click()

   objektGabimi.kapGabimin
   objektGabimi.menazhim_gabimi
   If objektGabimi.mvarGabimi = 0 Then
      pastroGriden
      objectInitialization KONSULTIME_AMZA
      'rtbShenime.Text = ""
      
      'objectInitialization AMZA_SJELLJE
      
      
   End If
End Sub

Private Sub Command1_Click()
    CallHelp indeksHelp
End Sub

Private Sub Form_Load()
   loadForm Me
   mbush_combo
   'viti
   Image1.Picture = LoadPicture(adresaLogo)
   lblShkolla.Caption = emerShkolla
   ' initializeObjects
   ' inicializoGridaAmza
   'inicializoGridaLendet
   gridaAmza.Visible = False
   gridaPjekurie.Visible = False
   Label4.Visible = False
   Label5.Visible = False
   Label1.Visible = False
   Label7.Visible = False
   Label6.Visible = False
   lblVitiShkollor.Visible = False
   cboVitiShkollor.Visible = False
   rtbShenime.Visible = False
   
End Sub

Private Sub initializeObjects()
  Set objUIController = New clsUIController
  Set objBusManager = New clsBusManager
End Sub

Private Sub inicializoGridaAmzaEmesme()
    
    
    
    gridaAmza.Height = 1850
    gridaAmza.Width = 14575
    gridaAmza.Rows = 5
    gridaAmza.Cols = 18
    gridaAmza.FixedRows = 1
    gridaAmza.FixedCols = 0
    gridaAmza.row = 0
    gridaAmza.col = 0
    gridaAmza.ColWidth(0) = 1500
    gridaAmza.Text = "Klasa"
    gridaAmza.col = 1
    gridaAmza.ColWidth(1) = 1500
    gridaAmza.Text = "Viti shkollor"
    gridaAmza.col = 2
    gridaAmza.ColWidth(2) = 3000
    gridaAmza.Text = "Letersi dhe Gjuhe Shqipe"
    gridaAmza.col = 3
    gridaAmza.ColWidth(3) = 2500
    gridaAmza.Text = "Gjuhe e huaj"
    gridaAmza.col = 4
    gridaAmza.ColWidth(4) = 2500
    gridaAmza.Text = "Histori"
    gridaAmza.col = 5
    gridaAmza.ColWidth(5) = 2500
    gridaAmza.Text = "Edukim Artistik"
    gridaAmza.col = 6
    gridaAmza.ColWidth(6) = 2500
    gridaAmza.Text = "Gjeografi"
    gridaAmza.col = 7
    gridaAmza.ColWidth(7) = 2500
    gridaAmza.Text = "Njohuri Per Shoqerine"
    gridaAmza.col = 8
    gridaAmza.ColWidth(8) = 2500
    gridaAmza.Text = "Njohuri Per Ekonomine"
    gridaAmza.col = 9
    gridaAmza.ColWidth(9) = 2500
    gridaAmza.Text = "Psikologji"
    gridaAmza.col = 10
    gridaAmza.ColWidth(10) = 2500
    gridaAmza.Text = "Filozofi"
    gridaAmza.col = 11
    gridaAmza.ColWidth(11) = 2500
    gridaAmza.Text = "Matematike"
    gridaAmza.col = 12
    gridaAmza.ColWidth(12) = 2500
    gridaAmza.Text = "Fizike"
    gridaAmza.col = 13
    gridaAmza.ColWidth(13) = 2500
    gridaAmza.Text = "Kimi"
    gridaAmza.col = 14
    gridaAmza.ColWidth(14) = 2500
    gridaAmza.Text = "Biologji"
    gridaAmza.col = 15
    gridaAmza.ColWidth(15) = 2500
    gridaAmza.Text = "Teknologji"
    gridaAmza.col = 16
    gridaAmza.ColWidth(16) = 2500
    gridaAmza.Text = "Informatike"
    gridaAmza.col = 17
    gridaAmza.ColWidth(17) = 2500
    gridaAmza.Text = "Edukim Fizik"
    
    gridaAmza.col = 0
    gridaAmza.row = 1
    gridaAmza.Text = "9"
    gridaAmza.row = 2
    gridaAmza.Text = "10"
    gridaAmza.row = 3
    gridaAmza.Text = "11"
    gridaAmza.row = 4
    gridaAmza.Text = "12"
    
    gridaAmza.RowHeight(0) = 300
    gridaAmza.RowHeight(1) = 300
    gridaAmza.RowHeight(2) = 300
    gridaAmza.RowHeight(3) = 300
    gridaAmza.RowHeight(4) = 300
    
End Sub

Private Sub inicializoGridaTetevjecare()
    gridaAmza.Height = 2940
    gridaAmza.Width = 14575
    gridaAmza.Rows = 9
    gridaAmza.Cols = 18
    gridaAmza.FixedRows = 1
    gridaAmza.FixedCols = 0
    gridaAmza.row = 0
    gridaAmza.col = 0
    gridaAmza.ColWidth(0) = 1500
    gridaAmza.Text = "Klasa"
    gridaAmza.col = 1
    gridaAmza.ColWidth(1) = 1500
    gridaAmza.Text = "Viti shkollor"
    gridaAmza.col = 2
    gridaAmza.ColWidth(2) = 2000
    gridaAmza.Text = "Abetare"
    gridaAmza.col = 3
    gridaAmza.ColWidth(3) = 2000
    gridaAmza.Text = "Gjuhe shqipe"
    gridaAmza.col = 4
    gridaAmza.ColWidth(4) = 2000
    gridaAmza.Text = "Lexim letrar"
    gridaAmza.col = 5
    gridaAmza.ColWidth(5) = 2000
    gridaAmza.Text = "Gjuhe e huaj"
    gridaAmza.col = 6
    gridaAmza.ColWidth(6) = 2000
    gridaAmza.Text = "Histori"
    gridaAmza.col = 7
    gridaAmza.ColWidth(7) = 2000
    gridaAmza.Text = "Dituri natyre"
    gridaAmza.col = 8
    gridaAmza.ColWidth(8) = 2000
    gridaAmza.Text = "Gjeografi"
    gridaAmza.col = 9
    gridaAmza.ColWidth(9) = 2000
    gridaAmza.Text = "Matematike"
    gridaAmza.col = 10
    gridaAmza.ColWidth(10) = 2000
    gridaAmza.Text = "Fizike"
    gridaAmza.col = 11
    gridaAmza.ColWidth(11) = 2000
    gridaAmza.Text = "Kimi"
    gridaAmza.col = 12
    gridaAmza.ColWidth(12) = 2000
    gridaAmza.Text = "Biologji"
    gridaAmza.col = 13
    gridaAmza.ColWidth(13) = 2200
    gridaAmza.Text = "Edukate shoqerore"
    gridaAmza.col = 14
    gridaAmza.ColWidth(14) = 2000
    gridaAmza.Text = "Edukim figurativ"
    gridaAmza.col = 15
    gridaAmza.ColWidth(15) = 2000
    gridaAmza.Text = "Edukim muzikor"
    gridaAmza.col = 16
    gridaAmza.ColWidth(16) = 2000
    gridaAmza.Text = "Mesim pune"
    gridaAmza.col = 17
    gridaAmza.ColWidth(17) = 2000
    gridaAmza.Text = "Edukim Fizik"
    
    gridaAmza.col = 0
    gridaAmza.row = 1
    gridaAmza.Text = "1"
    gridaAmza.row = 2
    gridaAmza.Text = "2"
    gridaAmza.row = 3
    gridaAmza.Text = "3"
    gridaAmza.row = 4
    gridaAmza.Text = "4"
    gridaAmza.row = 5
    gridaAmza.Text = "5"
    gridaAmza.row = 6
    gridaAmza.Text = "6"
    gridaAmza.row = 7
    gridaAmza.Text = "7"
    gridaAmza.row = 8
    gridaAmza.Text = "8"
End Sub

Private Sub inicializoGridaLendet()
    
    gridaLendet.Width = 14575
    gridaLendet.Rows = 2
    gridaLendet.Cols = 3
    gridaLendet.FixedRows = 1
    gridaLendet.FixedCols = 0
    gridaLendet.row = 0
    gridaLendet.col = 0
    gridaLendet.ColWidth(0) = 3500
    gridaLendet.Text = "Matematike"
    gridaLendet.col = 1
    gridaLendet.ColWidth(1) = 3500
    gridaLendet.Text = "Fizike"
    gridaLendet.col = 2
    gridaLendet.ColWidth(2) = 7500
    gridaLendet.Text = "Letersi"
    
    
End Sub
Private Sub inicializoGridaPjekurieMesme()

    gridaPjekurie.Height = 600
    gridaPjekurie.Width = 14575
    gridaPjekurie.Rows = 2
    gridaPjekurie.Cols = 6
    gridaPjekurie.row = 0
    gridaPjekurie.col = 0
    gridaPjekurie.ColWidth(0) = 2000
    gridaPjekurie.Text = "Letersi"
    gridaPjekurie.col = 1
    gridaPjekurie.ColWidth(1) = 2000
    gridaPjekurie.Text = "Matematike"
    gridaPjekurie.col = 2
    gridaPjekurie.ColWidth(2) = 2000
    gridaPjekurie.Text = "Fizike"
    gridaPjekurie.col = 3
    gridaPjekurie.ColWidth(3) = 2000
    gridaPjekurie.Text = "Kimi-Biologji"
    
    gridaPjekurie.col = 4
    gridaPjekurie.ColWidth(4) = 2000
    gridaPjekurie.Text = "Histori-Gjeografi"
    gridaPjekurie.col = 5
    gridaPjekurie.ColWidth(5) = 4500
    gridaPjekurie.Text = "Njohuri per shoqerine, ekonomine dhe filozofine"
    
    
    
End Sub
Private Sub inicializoGridaPjekurieTetevjecare()

    gridaPjekurie.Height = 600
    gridaPjekurie.Width = 5100
    gridaPjekurie.Rows = 2
    gridaPjekurie.Cols = 2
    gridaPjekurie.row = 0
    gridaPjekurie.col = 0
    gridaPjekurie.ColWidth(0) = 3000
    gridaPjekurie.Text = "Gjuhe shqipe dhe lexim letrar"
    gridaPjekurie.row = 0
    gridaPjekurie.col = 1
    gridaPjekurie.ColWidth(1) = 2000
    gridaPjekurie.Text = "Matematike"
    
End Sub

Private Sub objectInitialization(actionName As GUI_ACTION__ENUM)
   Set objUIController = New clsUIController

   objUIController.actionName = actionName
   objUIController.ExecuteActions

   Set objUIController = Nothing
End Sub





Private Sub optMesme_Click()
   ' inicializoGridaAmzaEmesme
   gridaAmza.Clear
   objectInitialization KONSULTIME_AMZA_E_MESME
   formatoGrida gridaAmza, 10575
   gridaAmza.Visible = False
   Label4.Visible = False
   Label5.Visible = False
   Label6.Visible = False
   Label7.Visible = False
   lblVitiShkollor.Visible = False
   cboVitiShkollor.Visible = False

   'inicializoGridaPjekurieMesme
   gridaPjekurie.Visible = False
   Label1.Visible = False
   rtbShenime.Visible = False
   txtAmzaNo.Text = ""
   txtEmri.Text = ""
   txtMbiemri.Text = ""
   txtKlasa.Text = ""
   txtIndeksi.Text = ""
   txtAtesia.Text = ""
   txtMemesia.Text = ""
   txtSeksi.Text = ""
   txtVendlindja.Text = ""
   txtDatelindja.Text = ""
End Sub

Private Sub optUlet_Click()
   
   gridaAmza.Clear
   objectInitialization KONSULTIME_AMZA_TETEVJECARE
   formatoGrida gridaAmza, 10575
   gridaAmza.Visible = False
   Label4.Visible = False
   Label5.Visible = False
   Label6.Visible = False
   Label7.Visible = False
   'inicializoGridaPjekurieMesme
   gridaPjekurie.Visible = False
   Label1.Visible = False
   rtbShenime.Visible = False
   'gridaAmza.Visible = True
   'inicializoGridaPjekurieTetevjecare
   'gridaPjekurie.Visible = True
   'Label4.Visible = True
   'Label5.Visible = True
   'Label1.Visible = True
   'rtbShenime.Visible = True
   lblVitiShkollor.Visible = False
   cboVitiShkollor.Visible = False
   txtAmzaNo.Text = ""
   txtEmri.Text = ""
   txtMbiemri.Text = ""
   txtKlasa.Text = ""
   txtIndeksi.Text = ""
   txtAtesia.Text = ""
   txtMemesia.Text = ""
   txtSeksi.Text = ""
   txtVendlindja.Text = ""
   txtDatelindja.Text = ""
End Sub
Private Sub viti()
   Dim d, muai, viti    As String
   d = Now
   Dim m, v             As Integer
   muai = Mid(d, 4, 2)
   viti = Mid(d, 7, 4)
   m = Val(muai)
   v = Val(viti)
   cboVitiShkollor.Text = gjej_vitin(m, v)

End Sub
Private Sub mbush_combo()
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
End Sub

Private Sub pastroGriden()
    
    Dim i As Integer
    Dim j As Integer
    Dim nr_rreshta As Integer
    Dim nr_shtylla As Integer
    nr_rreshta = gridaAmza.Rows
    nr_shtylla = gridaAmza.Cols
    For i = 1 To nr_rreshta - 1
        For j = 1 To nr_shtylla - 1
            gridaAmza.row = i
            gridaAmza.col = j
            gridaAmza.Text = ""
        Next j
    Next i
    
End Sub

Private Sub percakto_Nxenesit_Ngeles()
    
    Dim i As Integer
    Dim j As Integer
    Dim ugjet As Boolean
    Dim nota As Integer
    Dim nr_rreshta As Integer
    Dim nr_shtylla As Integer
    nr_rreshta = gridaAmza.Rows
    nr_shtylla = gridaAmza.Cols
    For i = 1 To nr_rreshta - 1
        ugjet = False
        j = 1
        Do While Not ugjet And j < nr_shtylla
            gridaAmza.row = i
            gridaAmza.col = j
            nota = gridaAmza.Text
            If nota = "4" Then
                ugjet = True
            End If
            j = j + 1
        Loop
        If ugjet Then
            gridaAmza.row = i
            gridaAmza.col = nr_shtylla - 1
            If gridaAmza.Text <> "" Then
                gridaAmza.Text = "Ka ngelur"
            End If
            
        End If
        
    Next i
End Sub

Private Sub formatoGrida(grida As MSFlexGrid, lartesia As Long)
    
    Dim l As Long
    Dim i As Integer
    i = grida.Cols
    l = i * 1500 + 100
    If l < lartesia Then
        grida.Width = l
    End If
    
    
End Sub

