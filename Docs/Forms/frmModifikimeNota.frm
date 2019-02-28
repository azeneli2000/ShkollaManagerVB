VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{04759352-AB08-11D5-863E-0010B563AF37}#1.0#0"; "Cgridv11.ocx"
Begin VB.Form frmModifikimeNota 
   Caption         =   "Modifikime Nota"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7935
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame fraInfo 
      Height          =   1575
      Left            =   12360
      TabIndex        =   26
      Top             =   0
      Width           =   2655
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
         TabIndex        =   28
         Top             =   600
         Width           =   1095
      End
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
         TabIndex        =   27
         Top             =   960
         Width           =   1095
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   120
         Picture         =   "frmModifikimeNota.frx":0000
         Stretch         =   -1  'True
         Top             =   600
         Width           =   960
      End
      Begin VB.Label lblWebsite 
         Height          =   375
         Left            =   1560
         TabIndex        =   30
         Top             =   2040
         Width           =   1815
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
         TabIndex        =   29
         Top             =   120
         Width           =   2415
      End
   End
   Begin VB.TextBox txtData 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   6480
      TabIndex        =   23
      Top             =   2040
      Width           =   1455
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
      Left            =   9480
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   840
      Width           =   735
   End
   Begin VB.Frame fraKerkim 
      Caption         =   "Kerkim sipas ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   1680
      TabIndex        =   5
      Top             =   120
      Width           =   10455
      Begin VB.TextBox txtAtesia 
         BackColor       =   &H00C0C0FF&
         Height          =   375
         Left            =   5160
         TabIndex        =   31
         Top             =   720
         Width           =   1575
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
         TabIndex        =   10
         Top             =   720
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
         TabIndex        =   9
         Top             =   720
         Width           =   1575
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
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   720
         Width           =   615
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
         Left            =   480
         TabIndex        =   7
         Top             =   720
         Width           =   1215
      End
      Begin VB.ComboBox cboVitiShkollor 
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
         ItemData        =   "frmModifikimeNota.frx":6F75
         Left            =   8760
         List            =   "frmModifikimeNota.frx":6F88
         TabIndex        =   6
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label4 
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
         TabIndex        =   32
         Top             =   480
         Width           =   1335
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
         TabIndex        =   16
         Top             =   480
         Width           =   495
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
         TabIndex        =   15
         Top             =   480
         Width           =   735
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
         Left            =   6840
         TabIndex        =   14
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblAmza 
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
         Height          =   375
         Left            =   480
         TabIndex        =   13
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lblVitiShkollor 
         Caption         =   "Viti Shkollor"
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
         Left            =   8760
         TabIndex        =   12
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label1 
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
         Left            =   7800
         TabIndex        =   11
         Top             =   480
         Width           =   735
      End
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
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7440
      Width           =   2295
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
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7440
      Width           =   2295
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
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7440
      Width           =   2055
   End
   Begin VB.OptionButton optMesme 
      Caption         =   "Shkolla e mesme"
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   975
   End
   Begin VB.OptionButton optUlet 
      Caption         =   "Shkolla tetevjecare"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   1095
   End
   Begin Cgridv11.Cgrid Lendet 
      Height          =   945
      Left            =   0
      TabIndex        =   17
      Top             =   2760
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
      TabIndex        =   18
      Top             =   2760
      Visible         =   0   'False
      Width           =   19185
      _ExtentX        =   33840
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
   Begin MSComCtl2.FlatScrollBar FlatScrollBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   6840
      Visible         =   0   'False
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Arrows          =   65536
      Orientation     =   1245185
      Value           =   1
   End
   Begin VB.Frame fraCikli 
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
      Height          =   1575
      Left            =   120
      TabIndex        =   25
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   6480
      TabIndex        =   24
      Top             =   1680
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Lendet"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      TabIndex        =   22
      Top             =   2400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblNotatDheMungesat 
      Alignment       =   2  'Center
      Caption         =   "Notat  Dhe  Mungesat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2400
      TabIndex        =   20
      Top             =   2400
      Visible         =   0   'False
      Width           =   2775
   End
End
Attribute VB_Name = "frmModifikimeNota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objUIController As clsUIController
Dim objektGabimi As New clsErrorHandler

Private Sub Cgrid1_CellClick(row As Long, col As Long)
   rreshtiModifiko = row
   shtyllaModifiko = col
   txtData.Text = ""
   If Cgrid1.Text(row, col) <> "" Then
      If Not IsNull(matricaNotat(row, col)) Then
         txtData.Text = matricaNotat(row, col)

      End If

      frmModifikoOseFshi.Show

   End If

End Sub

Private Sub Cgrid1_CellLostFocus(row As Long, col As Long)
   ' txtData.Text = ""
End Sub

Private Sub cmdDil_Click()
   Unload Me
   Set active_form = Nothing
End Sub


Private Sub cmdOK_Click()
   objektGabimi.kapGabimin
   objektGabimi.menazhim_gabimi
   If objektGabimi.mvarGabimi = 0 Then
      txtData.Visible = True
      objectInitialization KONSULTIME_NOTA_MOMENTALEI
   End If
End Sub

Private Sub Command1_Click()
    CallHelp indeksHelp
End Sub

Private Sub Form_Activate()


  'SGGrid1.Width = Me.Width - 300
  'SGGrid1.Height = SGGrid1.Height + 900
  
End Sub

Private Sub Form_Load()
  
  loadForm Me
  viti
  Image1.Picture = LoadPicture(adresaLogo)
  lblShkolla.Caption = emerShkolla
  optMesme.Value = False
  optUlet.Value = False
  txtData.Visible = False
  'inicializoGridaNotat
End Sub

Private Sub Form_Resize()
  'SGGrid1.Width = Me.Width - 300
  'fraInfo.Left = Me.Left + Me.Width - fraInfo.Width - 200
  
  'fraKerkim.Width = Me.Width - 300
  'SGGrid1.Height = SGGrid1.Height + 900
End Sub

Private Sub inicializoGridaNotat()
    
   ' gridaNotat.Width = 14575
    'gridaNotat.Rows = 2
    'gridaNotat.Cols = 3
    'gridaNotat.FixedRows = 1
    'gridaNotat.FixedCols = 0
    'gridaNotat.Row = 0
    'gridaNotat.col = 0
    'gridaNotat.ColWidth(0) = 2000
    'gridaNotat.Text = "Nr"
    'gridaNotat.col = 1
    'gridaNotat.ColWidth(1) = 2500
    'gridaNotat.Text = "Lenda"
    'gridaNotat.col = 2
    'gridaNotat.ColWidth(2) = 10000
    'gridaNotat.Text = "Notat dhe Mungesat"
    
    
End Sub

Private Sub objectInitialization(actionName As GUI_ACTION__ENUM)
   Set objUIController = New clsUIController

   objUIController.actionName = actionName
   objUIController.ExecuteActions

   Set objUIController = Nothing
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
Private Sub mbushkomboboks()
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
