VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{04759352-AB08-11D5-863E-0010B563AF37}#1.0#0"; "Cgridv11.ocx"
Begin VB.Form frmModifikimeNota 
   Caption         =   "Modifikime Nota"
   ClientHeight    =   9960
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9960
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
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
      TabIndex        =   20
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox label3 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
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
      TabIndex        =   42
      Text            =   "Data"
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox lblLargimi 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   41
      Top             =   2040
      Width           =   3615
   End
   Begin MSComCtl2.FlatScrollBar fsbHorizontali 
      Height          =   255
      Left            =   0
      TabIndex        =   18
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
   Begin VB.TextBox label2 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   300
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   38
      Text            =   "Lendet"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox lblNotatDheMungesat 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   300
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   37
      Text            =   "Notat dhe Mungesat"
      Top             =   2400
      Width           =   2655
   End
   Begin MSComCtl2.FlatScrollBar fsbVertical 
      Height          =   5415
      Left            =   15000
      TabIndex        =   36
      Top             =   2760
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   9551
      _Version        =   393216
      Appearance      =   0
      Orientation     =   1245184
   End
   Begin VB.TextBox txtTipi 
      Height          =   375
      Left            =   8640
      TabIndex        =   34
      Top             =   2280
      Visible         =   0   'False
      Width           =   975
   End
   Begin Cgridv11.Cgrid gridaProvime 
      Height          =   2175
      Left            =   240
      TabIndex        =   32
      Top             =   3360
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   3836
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
   Begin VB.Frame fraLlojNote 
      Caption         =   "Nota :"
      ForeColor       =   &H00008000&
      Height          =   1695
      Left            =   1920
      TabIndex        =   29
      Top             =   120
      Width           =   2055
      Begin VB.OptionButton optProvimi 
         Caption         =   "Provim lirimi ose pjekurie"
         ForeColor       =   &H000000C0&
         Height          =   615
         Left            =   240
         TabIndex        =   31
         Top             =   960
         Width           =   1695
      End
      Begin VB.OptionButton optNota 
         Caption         =   "Nota te zakonshme"
         ForeColor       =   &H00C00000&
         Height          =   615
         Left            =   240
         TabIndex        =   30
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame fraInfo 
      Height          =   1575
      Left            =   12360
      TabIndex        =   22
      Top             =   120
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   26
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
         TabIndex        =   25
         Top             =   120
         Width           =   2415
      End
   End
   Begin VB.TextBox txtIndeksi 
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
      Left            =   5880
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   1320
      Width           =   975
   End
   Begin VB.Frame fraKerkim 
      Caption         =   "Kerkim sipas ..."
      ForeColor       =   &H00008000&
      Height          =   1695
      Left            =   4080
      TabIndex        =   4
      Top             =   120
      Width           =   7695
      Begin VB.CommandButton cmdKerko 
         BackColor       =   &H00FFFFFF&
         Default         =   -1  'True
         DisabledPicture =   "frmModifikimeNota.frx":6F75
         DownPicture     =   "frmModifikimeNota.frx":D3B7
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
         Left            =   5280
         Picture         =   "frmModifikimeNota.frx":137F9
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   1200
         Width           =   1695
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
         Left            =   5280
         TabIndex        =   27
         Top             =   480
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
         TabIndex        =   9
         Top             =   480
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
         TabIndex        =   8
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtKlasa 
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
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1200
         Width           =   1215
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
         Left            =   480
         TabIndex        =   6
         Top             =   480
         Width           =   1215
      End
      Begin VB.ComboBox cboVitiShkollor 
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
         ItemData        =   "frmModifikimeNota.frx":19C3B
         Left            =   3480
         List            =   "frmModifikimeNota.frx":19C3D
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Atesia"
         Height          =   255
         Left            =   5280
         TabIndex        =   28
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblEmri 
         Caption         =   "Emri"
         Height          =   255
         Left            =   1800
         TabIndex        =   15
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblMbiemri 
         Caption         =   "Mbiemri"
         Height          =   255
         Left            =   3480
         TabIndex        =   14
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblKlasa 
         Caption         =   "Klasa"
         Height          =   255
         Left            =   480
         TabIndex        =   13
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblAmza 
         Caption         =   "Numri Amzes"
         Height          =   375
         Left            =   480
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblVitiShkollor 
         Caption         =   "Viti Shkollor"
         Height          =   255
         Left            =   3480
         TabIndex        =   11
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Indeksi"
         Height          =   255
         Left            =   1800
         TabIndex        =   10
         Top             =   960
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdDil 
      BackColor       =   &H80000009&
      DisabledPicture =   "frmModifikimeNota.frx":19C3F
      DownPicture     =   "frmModifikimeNota.frx":20081
      Height          =   375
      Left            =   5760
      Picture         =   "frmModifikimeNota.frx":264C3
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9120
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000009&
      DisabledPicture =   "frmModifikimeNota.frx":2C905
      DownPicture     =   "frmModifikimeNota.frx":2FE3F
      Height          =   375
      Left            =   11760
      Picture         =   "frmModifikimeNota.frx":33379
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9120
      Width           =   2535
   End
   Begin VB.OptionButton optMesme 
      Caption         =   "Shkolla e mesme"
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   975
   End
   Begin VB.OptionButton optUlet 
      Caption         =   "Shkolla nëntëvjeçare"
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Frame fraCikli 
      Caption         =   "Cikli"
      ForeColor       =   &H00008000&
      Height          =   1695
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   2900
      Left            =   0
      TabIndex        =   39
      Top             =   8150
      Width           =   15300
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   2745
      Left            =   0
      TabIndex        =   40
      Top             =   0
      Width           =   15500
   End
   Begin Cgridv11.Cgrid Lendet 
      Height          =   945
      Left            =   0
      TabIndex        =   16
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
      TabIndex        =   17
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
   Begin VB.Label Label5 
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   240
      TabIndex        =   33
      Top             =   2880
      Width           =   3975
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
   Dim nota             As String
   nota = Cgrid1.Text(row, col)
   If nota <> "" Then
      If Not IsNull(matricaNotat(row, col)) Then
         txtData.Text = matricaNotat(row, col)
         txtTipi.Text = modifikoNotat(row, col)
      End If
      If Cgrid1.Text(row, col) <> "" Then
         Cgrid1.CellSetFocus row, col
         
         frmModifikoOseFshi.show vbModal

      End If
   End If

End Sub

Private Sub Cgrid1_CellLostFocus(row As Long, col As Long)
   ' txtData.Text = ""
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

   gridaProvime.Visible = False
   Lendet.Visible = False
   Cgrid1.Visible = False
   Label2.Visible = False
   Label3.Visible = False
   txtData.Visible = False
   lblNotatDheMungesat.Visible = False
   Label5.Caption = ""
   objektGabimi.kapGabimin
   objektGabimi.menazhim_gabimi
   If objektGabimi.mvarGabimi = 0 Then

      If optNota.Value Then
         objectInitialization SHFAQJA_E_NOTAVE       'KONSULTIME_NOTA_MOMENTALEI
      Else
         objectInitialization MODIFIKIMI_I_PROVIMEVE
      End If
   End If
   
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

Private Sub Form_Activate()


  'SGGrid1.Width = Me.Width - 300
  'SGGrid1.Height = SGGrid1.Height + 900
  
End Sub

Private Sub Form_Load()
  
  loadForm Me
  mbushkomboboks
  viti
  Image1.Picture = LoadPicture(adresaLogo)
  lblShkolla.Caption = emerShkolla
  optMesme.Value = False
  optUlet.Value = False
  txtData.Visible = False
  percaktoRendinSipasTabit
  percaktoTeDrejtat
  Pastro
  'fsbVertical.Visible = False
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
   Dim M, v             As Integer
   muai = DateTime.Month(DateTime.Now)
   viti = DateTime.Year(DateTime.date)
   M = Val(muai)
   v = Val(viti)
   cboVitiShkollor.Text = gjej_vitin(M, v)

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

Private Sub percaktoRendinSipasTabit()
    txtAmzaNo.TabIndex = 0
    txtEmri.TabIndex = 1
    txtMbiemri.TabIndex = 2
    txtAtesia.TabIndex = 3
    'cmdKerko.TabIndex = 4
End Sub




Private Sub fsbHorizontali_Change()
     Cgrid1.Left = 2400 - fsbHorizontali.Value
End Sub

Private Sub fsbHorizontali_Scroll()
     Cgrid1.Left = 2400 - fsbHorizontali.Value
End Sub

Private Sub fsbVertical_Change()
    Cgrid1.Top = 2760 - fsbVertical.Value
    Lendet.Top = 2760 - fsbVertical.Value
End Sub

Private Sub fsbVertical_Scroll()
    Cgrid1.Top = 2760 - fsbVertical.Value
    Lendet.Top = 2760 - fsbVertical.Value
End Sub

Private Sub gridaProvime_CellClick(row As Long, col As Long)
   Dim I, j             As Integer
   I = row
   j = col
   Dim nota     As String
   If I = 1 Then
      Exit Sub
   End If
   
   rreshtiModifiko = row
   shtyllaModifiko = col
   txtData.Text = ""
  
   nota = gridaProvime.Text(row, col)
   If nota <> "" Then
      
      If gridaProvime.Text(row, col) <> "" Then
         gridaProvime.CellSetFocus row, col
         
         frmModifikoOseFshi.show vbModal

      End If
   End If
   
End Sub

Private Sub percaktoTeDrejtat()
    
    Select Case statusi
        
        Case "SupervizorEmesme"
            optMesme.Visible = False
            optUlet.Visible = False
            optMesme.Value = True
            optUlet.Value = False
            optProvimi.Caption = "Provimi i matures"
            fraCikli.Visible = False
        
        Case "SupervizorTetevjecare"
            optMesme.Visible = False
            optUlet.Visible = False
            optMesme.Value = False
            optUlet.Value = True
            optProvimi.Caption = "Provimi i lirimit"
            fraCikli.Visible = False
        Case Else
            
    End Select
    
            
            
End Sub

Private Sub Pastro()
   txtAmzaNo.Text = ""
   txtEmri.Text = ""
   txtMbiemri.Text = ""
   txtAtesia.Text = ""
   txtKlasa.Text = ""
   txtIndeksi.Text = ""
   lblLargimi.Text = ""
   viti
   txtData.Visible = False
   Label3.Visible = False
   Label2.Visible = False
   lblNotatDheMungesat.Visible = False
   Lendet.Visible = False
   Cgrid1.Visible = False
   gridaProvime.Visible = False
   fsbHorizontali.Visible = False
   fsbVertical.Visible = False
   Label5.Caption = ""

End Sub

Private Sub optMesme_Click()
    Pastro
End Sub

Private Sub optNota_Click()
    Pastro
    tipiModifikoProvime = True
End Sub

Private Sub optProvimi_Click()
    Pastro
    tipiModifikoProvime = False
End Sub

Private Sub optUlet_Click()
    Pastro
End Sub
