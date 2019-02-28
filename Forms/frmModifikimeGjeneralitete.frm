VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmModifikimeGjeneralitete 
   Caption         =   "Modifikime - Gjeneralitete"
   ClientHeight    =   9720
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9720
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame fraInfo 
      Height          =   1575
      Left            =   12360
      TabIndex        =   24
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
         TabIndex        =   26
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
         TabIndex        =   25
         Top             =   960
         Width           =   1095
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   120
         Picture         =   "frmModifikimeGjeneralitete.frx":0000
         Stretch         =   -1  'True
         Top             =   600
         Width           =   960
      End
      Begin VB.Label lblWebsite 
         Height          =   375
         Left            =   1560
         TabIndex        =   28
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
         TabIndex        =   27
         Top             =   120
         Width           =   2415
      End
   End
   Begin VB.OptionButton optUlet 
      Caption         =   "Shkolla nëntëvjeçare"
      ForeColor       =   &H000000C0&
      Height          =   855
      Left            =   240
      TabIndex        =   22
      Top             =   1080
      Width           =   1335
   End
   Begin VB.OptionButton optMesme 
      Caption         =   "Shkolla e mesme"
      ForeColor       =   &H00C00000&
      Height          =   855
      Left            =   240
      TabIndex        =   21
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000009&
      DisabledPicture =   "frmModifikimeGjeneralitete.frx":6F75
      DownPicture     =   "frmModifikimeGjeneralitete.frx":A4AF
      Height          =   375
      Left            =   11520
      Picture         =   "frmModifikimeGjeneralitete.frx":D9E9
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   9120
      Width           =   2055
   End
   Begin VB.Frame fraKerkim 
      Caption         =   "Te dhenat e amzes ..."
      ForeColor       =   &H00008000&
      Height          =   2175
      Left            =   1800
      TabIndex        =   8
      Top             =   0
      Width           =   10335
      Begin VB.TextBox txtMbulo 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   1440
         Width           =   1935
      End
      Begin VB.ComboBox cboVitiShkollor 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6480
         TabIndex        =   33
         Top             =   1440
         Width           =   1935
      End
      Begin VB.ComboBox cboIndeksi 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5400
         TabIndex        =   32
         Top             =   1440
         Width           =   975
      End
      Begin VB.ComboBox cboKlasa 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4200
         TabIndex        =   31
         Top             =   1440
         Width           =   1095
      End
      Begin VB.ComboBox cboSeksi 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8520
         TabIndex        =   30
         Top             =   600
         Width           =   1575
      End
      Begin VB.CommandButton cmdKerko 
         BackColor       =   &H00FFFFFF&
         Default         =   -1  'True
         DisabledPicture =   "frmModifikimeGjeneralitete.frx":10F23
         DownPicture     =   "frmModifikimeGjeneralitete.frx":17365
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
         Left            =   8520
         Picture         =   "frmModifikimeGjeneralitete.frx":1D7A7
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1440
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
         Left            =   1560
         TabIndex        =   1
         Top             =   600
         Width           =   1695
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
         Left            =   3360
         TabIndex        =   2
         Top             =   600
         Width           =   1695
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
         Width           =   1335
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
         TabIndex        =   3
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtMemesia 
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
         TabIndex        =   4
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtVendlindja 
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
         TabIndex        =   5
         Top             =   1440
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker txtDatelindja 
         Height          =   375
         Left            =   2160
         TabIndex        =   29
         Top             =   1440
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   0
         CalendarTitleForeColor=   -2147483634
         Format          =   20119553
         UpDown          =   -1  'True
         CurrentDate     =   38308
      End
      Begin VB.Label Label1 
         Caption         =   "Viti shkollor"
         Height          =   255
         Left            =   6480
         TabIndex        =   34
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Indeksi"
         Height          =   255
         Left            =   5400
         TabIndex        =   20
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label lblEmri 
         Caption         =   "Emri"
         Height          =   255
         Left            =   1560
         TabIndex        =   17
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblMbiemri 
         Caption         =   "Mbiemri"
         Height          =   255
         Left            =   3360
         TabIndex        =   16
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblKlasa 
         Caption         =   "Klasa"
         Height          =   255
         Left            =   4200
         TabIndex        =   15
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lblAmzaNo 
         Caption         =   "Numri Amzes"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblAtesia 
         Caption         =   "Atesia"
         Height          =   255
         Left            =   5160
         TabIndex        =   13
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblMemesia 
         Caption         =   "Memesia"
         Height          =   255
         Left            =   6960
         TabIndex        =   12
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblSeksi 
         Caption         =   "Seksi"
         Height          =   255
         Left            =   8520
         TabIndex        =   11
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblVendlindja 
         Caption         =   "Vendlindja"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label lblDatelindja 
         Caption         =   "Datelindja"
         Height          =   255
         Left            =   2160
         TabIndex        =   9
         Top             =   1200
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H80000009&
      DisabledPicture =   "frmModifikimeGjeneralitete.frx":23BE9
      DownPicture     =   "frmModifikimeGjeneralitete.frx":2A02B
      Height          =   375
      Left            =   3840
      Picture         =   "frmModifikimeGjeneralitete.frx":3046D
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9120
      Width           =   2175
   End
   Begin VB.CommandButton cmdDil 
      BackColor       =   &H80000009&
      DisabledPicture =   "frmModifikimeGjeneralitete.frx":368AF
      DownPicture     =   "frmModifikimeGjeneralitete.frx":3CCF1
      Height          =   375
      Left            =   7920
      Picture         =   "frmModifikimeGjeneralitete.frx":43133
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9120
      Width           =   2055
   End
   Begin VB.Frame fraCikli 
      Caption         =   "Cikli"
      ForeColor       =   &H00008000&
      Height          =   2175
      Left            =   120
      TabIndex        =   23
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "frmModifikimeGjeneralitete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objUIController As clsUIController
Dim objektGabimi As New clsErrorHandler










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
   cmdKerko.Enabled = False
   objektGabimi.kapGabimin
   If objektGabimi.mvarGabimi = 0 Or objektGabimi.mvarGabimi = 19 Or objektGabimi.mvarGabimi = 20 Then
      
      objectInitialization MODIFIKO_GJENERALITETE1
      'txtMbulo.Visible = False
   Else
      objektGabimi.menazhim_gabimi
   End If
   cmdKerko.Enabled = True
End Sub


Private Sub cmdOK_Click()
   Dim data             As String
   data = date
   objektGabimi.kapGabimin
   objektGabimi.menazhim_gabimi
   If objektGabimi.mvarGabimi = 0 Then
      objectInitialization MODIFIKO_GJENERALITETE2
      If data_modifikimi = "" Or data <> data_modifikimi Then
         objectInitialization MODIFIKIMI_I_DATES
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

Private Sub Form_Load()
  loadForm Me
  'viti
  mbushKomboBox
  Image1.Picture = LoadPicture(adresaLogo)
  lblShkolla.Caption = emerShkolla
  percaktoRendinSipasTabit
  'inicializoGridaAmza
  'inicializoGridaLendet
  txtDatelindja.Value = Now
  percaktoTeDrejtat
  cmdOK.Enabled = False
End Sub

Private Sub Form_Resize()
   'fraKerkim.Width = Me.Width - 300
  ' SGGrid1.Width = Me.Width - 300
  ' SGGrid2.Width = Me.Width - 300
   'rtbShenime.Width = Me.Width - 300
   fraInfo.Left = Me.Left + Me.Width - fraInfo.Width - 200
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


Private Sub mbushKomboBox()
    
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
    
    cboSeksi.AddItem "mashkull"
    cboSeksi.AddItem "femer"
End Sub

Private Sub percaktoRendinSipasTabit()
    
    txtAmzaNo.TabIndex = 0
    txtEmri.TabIndex = 1
    txtMbiemri.TabIndex = 2
    txtAtesia.TabIndex = 3
    txtMemesia.TabIndex = 4
    cboSeksi.TabIndex = 5
    txtVendlindja.TabIndex = 6
    txtDatelindja.TabIndex = 7
    
    cboKlasa.TabIndex = 8
    cboIndeksi.TabIndex = 9
    cmdKerko.TabIndex = 10
End Sub

Private Sub percaktoTeDrejtat()
    
    Select Case statusi
        
        Case "SupervizorEmesme"
            optMesme.Visible = False
            optUlet.Visible = False
            optMesme.Value = True
            optUlet.Value = False
            
            fraCikli.Visible = False
            mbushKomboMesme
        
        Case "SupervizorTetevjecare"
            optMesme.Visible = False
            optUlet.Visible = False
            optMesme.Value = False
            optUlet.Value = True
            
            fraCikli.Visible = False
            mbushKomboTetevjecare
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
    txtAmzaNo.Text = ""
    txtEmri.Text = ""
    txtMbiemri.Text = ""
    txtAtesia.Text = ""
    txtMemesia.Text = ""
    cboSeksi.Text = ""
    txtVendlindja.Text = ""
    txtMbulo.Visible = True
    cboKlasa.Text = ""
    cboIndeksi.Text = ""
    cboVitiShkollor.Text = ""
End Sub

Private Sub optMesme_Click()
    Pastro
End Sub

Private Sub optUlet_Click()
    Pastro
End Sub
