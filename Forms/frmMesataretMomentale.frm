VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmMesataretMomentale 
   Caption         =   "Statistika - Mesataret momentale sipas klasave"
   ClientHeight    =   9405
   ClientLeft      =   60
   ClientTop       =   450
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
   ScaleHeight     =   165.894
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   268.817
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   8535
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   113
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid amzaEmerMbiemer 
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   1920
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   450
      _Version        =   393216
      BackColor       =   -2147483626
      Rows            =   1
      FixedRows       =   0
      FixedCols       =   0
      ScrollBars      =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Lendet 
      Height          =   255
      Left            =   3000
      TabIndex        =   23
      Top             =   1920
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   450
      _Version        =   393216
      BackColor       =   -2147483626
      FixedCols       =   0
      ScrollBars      =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H80000009&
      DisabledPicture =   "frmMesataretMomentale.frx":0000
      DownPicture     =   "frmMesataretMomentale.frx":353A
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
      Left            =   5880
      Picture         =   "frmMesataretMomentale.frx":6A74
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   9000
      Width           =   2415
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
      Height          =   1575
      Left            =   12480
      TabIndex        =   13
      Top             =   0
      Width           =   2655
      Begin VB.CommandButton cmdWebsiste 
         BackColor       =   &H80000009&
         Caption         =   "Website"
         Height          =   375
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton cmdEmail 
         BackColor       =   &H80000009&
         Caption         =   "E-mail"
         Height          =   375
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   960
         Width           =   1095
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   120
         Picture         =   "frmMesataretMomentale.frx":9FAE
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
         TabIndex        =   17
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
         TabIndex        =   16
         Top             =   120
         Width           =   2415
      End
   End
   Begin VB.CommandButton cmdDil 
      BackColor       =   &H80000009&
      DisabledPicture =   "frmMesataretMomentale.frx":10F23
      DownPicture     =   "frmMesataretMomentale.frx":1445D
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
      Left            =   12600
      Picture         =   "frmMesataretMomentale.frx":17997
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   9000
      Width           =   2175
   End
   Begin VB.ComboBox cboIndeksi 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      ForeColor       =   &H80000004&
      Height          =   315
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1080
      Width           =   735
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
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   8055
      Begin VB.Frame fraOptions 
         Caption         =   "Mesataret momentale"
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
         Height          =   1335
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   2415
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
            Height          =   255
            Left            =   360
            TabIndex        =   4
            Top             =   840
            Width           =   1695
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
            Height          =   255
            Left            =   360
            TabIndex        =   3
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.CommandButton cmdKerko 
         BackColor       =   &H00FFFFFF&
         DisabledPicture =   "frmMesataretMomentale.frx":1AED1
         DownPicture     =   "frmMesataretMomentale.frx":21313
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
         Left            =   5520
         Picture         =   "frmMesataretMomentale.frx":27755
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   720
         Width           =   1815
      End
      Begin VB.ComboBox cboVitiShkollor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         ForeColor       =   &H80000004&
         Height          =   315
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   480
         Width           =   1575
      End
      Begin VB.ComboBox cboKlasa 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         ForeColor       =   &H80000004&
         Height          =   315
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label1 
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
         Left            =   4080
         TabIndex        =   10
         Top             =   840
         Width           =   735
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
         Left            =   3240
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Klasa 
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
         Left            =   3240
         TabIndex        =   6
         Top             =   840
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdDalje 
      BackColor       =   &H80000009&
      DisabledPicture =   "frmMesataretMomentale.frx":2DB97
      DownPicture     =   "frmMesataretMomentale.frx":33FD9
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
      Left            =   9120
      Picture         =   "frmMesataretMomentale.frx":3A41B
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9000
      Width           =   2535
   End
   Begin MSComCtl2.FlatScrollBar fsbVertikali 
      Height          =   6015
      Left            =   15030
      TabIndex        =   19
      Top             =   1920
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   10610
      _Version        =   393216
      Appearance      =   0
      Orientation     =   1245184
   End
   Begin MSComCtl2.FlatScrollBar fsbHorizontali 
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   7680
      Visible         =   0   'False
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Arrows          =   65536
      Orientation     =   1245185
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   1984
      Left            =   120
      TabIndex        =   25
      Top             =   15
      Width           =   15135
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid emrat 
      Height          =   4815
      Left            =   120
      TabIndex        =   22
      Top             =   2160
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   8493
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      ScrollBars      =   0
      Appearance      =   0
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid notat 
      Height          =   4815
      Left            =   3000
      TabIndex        =   24
      Top             =   2160
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   8493
      _Version        =   393216
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frmMesataretMomentale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objektGabimi As New clsErrorHandler
Dim objPrintim As Object
Dim objPrintForm As New frmPrintEvidenca
Dim objPrintClass As Object
Dim TabelNotash(20, 3) As String
Dim Gabim As Boolean


Private Sub cboIndeksi_Click()
    Pastro
End Sub

Private Sub cboKlasa_Click()
    Pastro
End Sub

Private Sub cboVitiShkollor_Click()
    Pastro
End Sub

Private Sub cmdDalje_Click()
   Unload Me
   Set active_form = Nothing
End Sub

Private Sub cmdEmail_Click()
    If Not OpenEmailProgram(email) Then
        MsgBox "Ju nuk keni asnje program te instaluar per te derguar email", vbExclamation, "Gabim ne dergimin e emailit"
    End If
End Sub

Private Sub cmdKerko_Click()
    Me.emrat.Visible = False
    Me.amzaEmerMbiemer.Visible = False
    Me.Lendet.Visible = False
    Me.notat.Visible = False
   objektGabimi.kapGabimin
   objektGabimi.menazhim_gabimi
   If objektGabimi.mvarGabimi = 0 Then
      objectInitialization VIZUALIZO_MESATARET_MOMENTALE
      Me.cmdPrint.Enabled = True
   End If
End Sub


Private Sub cmdWebsiste_Click()
    If website <> "" Then
        GoToWeb website
    Else
        MsgBox "Ju nuk e keni dhene adresen e faqes tuaj te web-it.", vbInformation
    End If
End Sub


Private Sub emrat_Click()
    Dim nr As Integer
    nr = emrat.row
    notat.row = nr
    notat.RowSel = nr
    notat.SetFocus
End Sub

Private Sub Form_Load()
   loadForm Me
   mbushKomboBox
   viti
   Image1.Picture = LoadPicture(adresaLogo)
   lblShkolla.Caption = emerShkolla
   'txtEmri(1).Visible = False
   fsbVertikali.Visible = False
   fsbHorizontali.Visible = False
   cmdPrint.Enabled = False
   notat.Visible = False
   emrat.Visible = False
   Lendet.Visible = False
   amzaEmerMbiemer.Visible = False
End Sub

Private Sub fsbHorizontali_Change()
    Dim t                As Integer
    Dim vlera            As Integer
    t = 50 + 2
    vlera = fsbHorizontali.Value
    Lendet.Left = t - vlera
    Lendet.Width = 6 * CDbl(1850 / 56.7) + vlera
    notat.Left = t - vlera
    notat.Width = 6 * CDbl(1850 / 56.7) + vlera
End Sub

Private Sub fsbHorizontali_Scroll()
    Dim t                As Integer
    Dim vlera            As Integer
    t = 50 + 2
    vlera = fsbHorizontali.Value
    Lendet.Left = t - vlera
    Lendet.Width = 6 * CDbl(1850 / 56.7) + vlera
    notat.Left = t - vlera
    notat.Width = 6 * CDbl(1850 / 56.7) + vlera
End Sub

Private Sub fsbVertikali_Change()
    Dim t As Integer
    t = 35 + CDbl(300 / 56.7)
    
    notat.Top = t - fsbVertikali.Value
    emrat.Top = t - fsbVertikali.Value
    notat.Height = 20 * CDbl(300 / 56.7) + fsbVertikali.Value
    emrat.Height = 20 * CDbl(300 / 56.7) + fsbVertikali.Value
End Sub

Private Sub fsbVertikali_Scroll()
    Dim t As Integer
    t = 35 + CDbl(300 / 56.7)
    
    notat.Top = t - fsbVertikali.Value
    emrat.Top = t - fsbVertikali.Value
    notat.Height = 20 * CDbl(300 / 56.7) + fsbVertikali.Value
    emrat.Height = 20 * CDbl(300 / 56.7) + fsbVertikali.Value
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
    
    cboIndeksi.AddItem "A"
    cboIndeksi.AddItem "B"
    cboIndeksi.AddItem "C"
    cboIndeksi.AddItem "D"
    cboIndeksi.AddItem "E"
    cboIndeksi.AddItem "F"
    
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

Private Sub notat_Click()
Dim nr As Integer
nr = notat.row
emrat.row = nr
emrat.RowSel = nr
emrat.SetFocus
End Sub

Private Sub optSemestri1_Click()
    Pastro
End Sub

Private Sub Pastro()
    notat.Visible = False
    emrat.Visible = False
    amzaEmerMbiemer.Visible = False
    Lendet.Visible = False
    fsbVertikali.Visible = False
    fsbHorizontali.Visible = False
    scrbar = False
    scrbarver = False
    cmdPrint.Enabled = False
End Sub

Private Sub optSemestri2_Click()
    Pastro
End Sub

