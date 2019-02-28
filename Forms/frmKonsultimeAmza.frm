VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmKonsultimeAmza 
   Caption         =   "Konsultime - Amza"
   ClientHeight    =   9840
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9840
   ScaleWidth      =   14910
   WindowState     =   2  'Maximized
   Begin RichTextLib.RichTextBox rtbShenime 
      Height          =   1935
      Left            =   120
      TabIndex        =   49
      Top             =   6600
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   3413
      _Version        =   393217
      TextRTF         =   $"frmKonsultimeAmza.frx":0000
   End
   Begin MSComCtl2.FlatScrollBar fsbHorizontali 
      Height          =   255
      Left            =   2120
      TabIndex        =   47
      Top             =   5445
      Width           =   12250
      _ExtentX        =   21616
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Arrows          =   65536
      LargeChange     =   3000
      Max             =   12000
      Orientation     =   1245185
      SmallChange     =   1500
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridaLabel 
      Height          =   300
      Left            =   120
      TabIndex        =   44
      Top             =   2700
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComCtl2.FlatScrollBar fsbVertikali 
      Height          =   2750
      Left            =   14115
      TabIndex        =   43
      Top             =   2700
      Width           =   270
      _ExtentX        =   476
      _ExtentY        =   4842
      _Version        =   393216
      Appearance      =   0
      LargeChange     =   300
      Max             =   3000
      Orientation     =   1245184
      SmallChange     =   300
   End
   Begin VB.ComboBox cboVitiShkollor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   360
      Left            =   4320
      Style           =   2  'Dropdown List
      TabIndex        =   37
      Top             =   6240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame fraInfo 
      Height          =   1575
      Left            =   12480
      TabIndex        =   32
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
         TabIndex        =   34
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
         TabIndex        =   33
         Top             =   600
         Width           =   1095
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
         TabIndex        =   36
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label lblWebsite 
         Height          =   375
         Left            =   1560
         TabIndex        =   35
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   120
         Picture         =   "frmKonsultimeAmza.frx":0082
         Stretch         =   -1  'True
         Top             =   600
         Width           =   960
      End
   End
   Begin VB.OptionButton optUlet 
      Caption         =   "Shkolla nëntëvjeçare"
      ForeColor       =   &H000000C0&
      Height          =   735
      Left            =   240
      TabIndex        =   31
      Top             =   1080
      Width           =   1335
   End
   Begin VB.OptionButton optMesme 
      Caption         =   "Shkolla e mesme"
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   240
      TabIndex        =   30
      Top             =   240
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cikli"
      ForeColor       =   &H00008000&
      Height          =   1935
      Left            =   120
      TabIndex        =   29
      Top             =   0
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid gridaPjekurie 
      Height          =   975
      Left            =   6840
      TabIndex        =   25
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
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000009&
      DisabledPicture =   "frmKonsultimeAmza.frx":6FF7
      DownPicture     =   "frmKonsultimeAmza.frx":A531
      Height          =   375
      Left            =   11280
      Picture         =   "frmKonsultimeAmza.frx":DA6B
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   9000
      Width           =   2175
   End
   Begin VB.TextBox txtKlasa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   1320
      Width           =   855
   End
   Begin VB.Frame fraKerkim 
      Caption         =   "Te dhenat e amzes ..."
      ForeColor       =   &H00008000&
      Height          =   1935
      Left            =   1800
      TabIndex        =   9
      Top             =   0
      Width           =   9615
      Begin VB.CommandButton cmdKerko 
         BackColor       =   &H00FFFFFF&
         DisabledPicture =   "frmKonsultimeAmza.frx":10FA5
         DownPicture     =   "frmKonsultimeAmza.frx":173E7
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   6960
         Picture         =   "frmKonsultimeAmza.frx":1D829
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox txtDatelindja 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txtEmri 
         BackColor       =   &H80000003&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000004&
         Height          =   375
         Left            =   1800
         TabIndex        =   1
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtMbiemri 
         BackColor       =   &H80000003&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000004&
         Height          =   375
         Left            =   3480
         TabIndex        =   2
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtIndeksi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txtAmzaNo 
         BackColor       =   &H80000003&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000004&
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtAtesia 
         BackColor       =   &H80000003&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000004&
         Height          =   375
         Left            =   5160
         TabIndex        =   4
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtMemesia 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtSeksi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txtVendlindja 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Indeksi"
         Height          =   255
         Left            =   6120
         TabIndex        =   22
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblEmri 
         Caption         =   "Emri"
         Height          =   255
         Left            =   1800
         TabIndex        =   18
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblMbiemri 
         Caption         =   "Mbiemri"
         Height          =   255
         Left            =   3480
         TabIndex        =   17
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblKlasa 
         Caption         =   "Klasa"
         Height          =   255
         Left            =   5160
         TabIndex        =   16
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblAmzaNo 
         Caption         =   "Numri Amzes"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblAtesia 
         Caption         =   "Atesia"
         Height          =   255
         Left            =   5160
         TabIndex        =   14
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblMemesia 
         Caption         =   "Memesia"
         Height          =   255
         Left            =   6960
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblSeksi 
         Caption         =   "Seksi"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lblVendlindja 
         Caption         =   "Vendlindja"
         Height          =   255
         Left            =   1800
         TabIndex        =   11
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label lblDatelindja 
         Caption         =   "Datelindja"
         Height          =   255
         Left            =   3480
         TabIndex        =   10
         Top             =   1080
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdDil 
      BackColor       =   &H80000009&
      DisabledPicture =   "frmKonsultimeAmza.frx":23C6B
      DownPicture     =   "frmKonsultimeAmza.frx":2A0AD
      Height          =   375
      Left            =   6480
      Picture         =   "frmKonsultimeAmza.frx":304EF
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9000
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   -480
      TabIndex        =   48
      Top             =   2400
      Width           =   600
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridaLendet 
      Height          =   250
      Left            =   2120
      TabIndex        =   42
      Top             =   2700
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   0
      TabIndex        =   45
      Top             =   0
      Width           =   15255
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridaKlasat 
      Height          =   2475
      Left            =   120
      TabIndex        =   40
      Top             =   3000
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   4366
      _Version        =   393216
      BackColorBkg    =   -2147483634
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridaAmza 
      Height          =   2475
      Left            =   2120
      TabIndex        =   41
      Top             =   3000
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   4366
      _Version        =   393216
      BackColorBkg    =   -2147483634
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
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
      TabIndex        =   38
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
      TabIndex        =   28
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
      TabIndex        =   27
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
      TabIndex        =   26
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
      TabIndex        =   24
      Top             =   2280
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "Indeksi"
      Height          =   255
      Left            =   7440
      TabIndex        =   21
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
      TabIndex        =   19
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

Private Sub cmdEmail_Click()
    If Not OpenEmailProgram(email) Then
        MsgBox "Ju nuk keni asnje program te instaluar per te derguar email", vbExclamation, "Gabim ne dergimin e emailit"
    End If
End Sub

Private Sub cmdKerko_Click()
    Me.gridaAmza.Visible = False
    Me.gridaKlasat.Visible = False
    Me.gridaLabel.Visible = False
    Me.gridaLendet.Visible = False
   cmdKerko.Enabled = False
   objektGabimi.kapGabimin
   objektGabimi.menazhim_gabimi
   If objektGabimi.mvarGabimi = 0 Then
      pastroGriden
      If Me.optMesme.Value Then
        objectInitialization KONSULTIME_AMZA_E_MESME
      Else
        objectInitialization KONSULTIME_AMZA_TETEVJECARE
      End If
      objectInitialization KONSULTIME_AMZA
      'rtbShenime.Text = ""
       
      'objectInitialization AMZA_SJELLJE
      formatoGrida gridaAmza, 2400
      formatoGrida gridaKlasat, 2700
      formatoGridaGjeresi gridaLendet, 12000
      formatoGridaGjeresi gridaAmza, 12000
      
      'Me.gridaAmza.Visible = True
      'Me.gridaKlasat.Visible = True
      'Me.gridaLabel.Visible = True
      'Me.gridaLendet.Visible = True
      Me.fsbHorizontali.Value = 0
      Me.fsbVertikali.Value = 0
      
   End If
   cmdKerko.Enabled = True
End Sub



Private Sub cmdPrint_Click()

End Sub

Private Sub cmdWebsiste_Click()
    If website <> "" Then
        GoToWeb website
    Else
        MsgBox "Ju nuk e keni dhene adresen e faqes tuaj te web-it.", vbInformation
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
   gridaKlasat.Visible = False
   gridaLendet.Visible = False
   gridaPjekurie.Visible = False
   gridaLabel.Visible = False
   fsbVertikali.Visible = False
   fsbHorizontali.Visible = False
   Label4.Visible = False
   Label5.Visible = False
   Label1.Visible = False
   Label7.Visible = False
   Label6.Visible = False
   lblVitiShkollor.Visible = False
   cboVitiShkollor.Visible = False
   rtbShenime.Visible = False
   percaktoRendinSipasTabit
   
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






Private Sub Frame2_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub fsbHorizontali_Change()
     Dim t As Integer
     t = 2120
     Dim nr As Integer
     nr = gridaAmza.Cols + 1 - 8
     Dim X As Integer
     X = fsbHorizontali.SmallChange
     Dim distanca As Integer
     distanca = CInt(12000 / nr)
     Dim vlera As Long
     vlera = CLng((fsbHorizontali.Value / 8) * nr)
    
     gridaAmza.Left = 2120 - vlera
     gridaAmza.Width = 12000 + vlera
     gridaLendet.Left = 2120 - vlera
     gridaLendet.Width = 12000 + vlera
End Sub

Private Sub fsbHorizontali_Scroll()
     Dim t As Integer
     t = 2120
     Dim gjeresia As Long
     gjeresia = gridaAmza.Cols * 1500 + 1500
     Dim distanca As Long
     distanca = gjeresia - 12000
     Dim nrLarge As Integer
     Dim nrSmall As Integer
     gridaAmza.Left = t - (fsbHorizontali.Value * distanca) / 12000
     gridaAmza.Width = 12000 + (fsbHorizontali.Value * distanca) / 12000
     gridaLendet.Left = t - (fsbHorizontali.Value * distanca) / 12000
     gridaLendet.Width = 12000 + (fsbHorizontali.Value * distanca) / 12000
End Sub



Private Sub fsbVertikali_Change()
    Dim t As Integer
    t = 3000
    
    gridaAmza.Top = t - fsbVertikali.Value
    gridaKlasat.Top = t - fsbVertikali.Value
    gridaKlasat.Height = 2400 + fsbVertikali.Value
    gridaAmza.Height = 2400 + fsbVertikali.Value
    
End Sub



Private Sub fsbVertikali_Scroll()
    Dim t As Integer
    t = 3000
    
    gridaAmza.Top = t - fsbVertikali.Value
    gridaKlasat.Top = t - fsbVertikali.Value
    gridaKlasat.Height = 2400 + fsbVertikali.Value
    gridaAmza.Height = 2400 + fsbVertikali.Value
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
   fsbVertikali.Visible = False
   fsbHorizontali.Visible = False
   'inicializoGridaPjekurieMesme
   gridaPjekurie.Visible = False
   gridaAmza.Visible = False
   gridaKlasat.Visible = False
   gridaLendet.Visible = False
   gridaLabel.Visible = False
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
   gridaAmza.Visible = False
   gridaKlasat.Visible = False
   gridaLendet.Visible = False
   fsbVertikali.Visible = False
   fsbHorizontali.Visible = False
   Label1.Visible = False
   rtbShenime.Visible = False
   gridaPjekurie.Visible = False
   gridaLabel.Visible = False
   
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
   Dim M, v             As Integer
   muai = DateTime.Month(DateTime.Now)
   viti = DateTime.Year(DateTime.date)
   M = Val(muai)
   v = Val(viti)
   cboVitiShkollor.Text = gjej_vitin(M, v)

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
End Sub

Private Sub pastroGriden()
    
    Dim I As Integer
    Dim j As Integer
    Dim nr_rreshta As Integer
    Dim nr_shtylla As Integer
    nr_rreshta = gridaAmza.Rows
    nr_shtylla = gridaAmza.Cols
    For I = 0 To nr_rreshta - 1
        For j = 0 To nr_shtylla - 1
            gridaAmza.row = I
            gridaAmza.col = j
            gridaAmza.Text = ""
        Next j
        For j = 0 To gridaKlasat.Cols - 1
            gridaKlasat.row = I
            gridaKlasat.col = j
            gridaKlasat.Text = ""
        Next
    Next I
    
End Sub

Private Sub percakto_Nxenesit_Ngeles()
    
    Dim I As Integer
    Dim j As Integer
    Dim ugjet As Boolean
    Dim nota As Integer
    Dim nr_rreshta As Integer
    Dim nr_shtylla As Integer
    nr_rreshta = gridaAmza.Rows
    nr_shtylla = gridaAmza.Cols
    For I = 1 To nr_rreshta - 1
        ugjet = False
        j = 1
        Do While Not ugjet And j < nr_shtylla
            gridaAmza.row = I
            gridaAmza.col = j
            nota = gridaAmza.Text
            If nota = "4" Then
                ugjet = True
            End If
            j = j + 1
        Loop
        If ugjet Then
            gridaAmza.row = I
            gridaAmza.col = nr_shtylla - 1
            If gridaAmza.Text <> "" Then
                gridaAmza.Text = "Ka ngelur"
            End If
            
        End If
        
    Next I
End Sub

Private Sub formatoGrida(grida As MSHFlexGrid, lartesia As Long)
    
    Dim l As Long
    Dim I As Integer
    I = grida.Cols
    l = I * CLng(1500) + 100
    If l < lartesia Then
        grida.Width = l
    End If
    
    
End Sub
Private Sub percaktoRendinSipasTabit()
    txtAmzaNo.TabIndex = 0
    txtEmri.TabIndex = 1
    txtMbiemri.TabIndex = 2
    txtAtesia.TabIndex = 3
   
End Sub

Private Sub formatoGridaGjeresi(grida As MSHFlexGrid, gjeresia As Long)

   Dim l                As Long
   Dim I                As Integer
   I = grida.Cols
   l = I * CLng(grida.ColWidth(0)) + 50
   If l < gjeresia Then
      grida.Width = l
   Else
      grida.Width = gjeresia
   End If

End Sub

