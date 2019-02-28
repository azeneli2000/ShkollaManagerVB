VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{04759352-AB08-11D5-863E-0010B563AF37}#1.0#0"; "Cgridv11.ocx"
Begin VB.Form frmHedhjeGjeneralitete 
   Caption         =   "Regjistrimi i nxenesit"
   ClientHeight    =   8505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13785
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8505
   ScaleWidth      =   13785
   WindowState     =   2  'Maximized
   Begin VB.Frame fraInfo 
      Height          =   1455
      Left            =   10800
      TabIndex        =   30
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
         TabIndex        =   32
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
         TabIndex        =   31
         Top             =   960
         Width           =   1095
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   120
         Picture         =   "frmHedhjeGjeneralitete.frx":0000
         Stretch         =   -1  'True
         Top             =   600
         Width           =   960
      End
      Begin VB.Label lblWebsite 
         Height          =   375
         Left            =   1560
         TabIndex        =   34
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
         TabIndex        =   33
         Top             =   120
         Width           =   2415
      End
   End
   Begin Cgridv11.Cgrid Provimet 
      Height          =   855
      Left            =   240
      TabIndex        =   24
      Top             =   3600
      Visible         =   0   'False
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   1508
      GridMode        =   0
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000009&
      DisabledPicture =   "frmHedhjeGjeneralitete.frx":6F75
      DownPicture     =   "frmHedhjeGjeneralitete.frx":A4AF
      Height          =   375
      Left            =   11760
      Picture         =   "frmHedhjeGjeneralitete.frx":D9E9
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   7920
      Width           =   2055
   End
   Begin VB.CommandButton cmdDil 
      BackColor       =   &H80000009&
      DisabledPicture =   "frmHedhjeGjeneralitete.frx":10F23
      DownPicture     =   "frmHedhjeGjeneralitete.frx":17365
      Height          =   375
      Left            =   6360
      Picture         =   "frmHedhjeGjeneralitete.frx":1D7A7
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7920
      Width           =   2535
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H80000009&
      DisabledPicture =   "frmHedhjeGjeneralitete.frx":23BE9
      DownPicture     =   "frmHedhjeGjeneralitete.frx":2A02B
      Height          =   375
      Left            =   3120
      Picture         =   "frmHedhjeGjeneralitete.frx":3046D
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7920
      Width           =   2535
   End
   Begin VB.Frame fraKerkim 
      Caption         =   "Te dhenat e amzes ..."
      ForeColor       =   &H00008000&
      Height          =   2175
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   10575
      Begin VB.TextBox txtData 
         BackColor       =   &H80000003&
         ForeColor       =   &H80000004&
         Height          =   375
         Left            =   4200
         TabIndex        =   35
         Top             =   1440
         Width           =   1935
      End
      Begin VB.ComboBox cboSeksi 
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
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   1440
         Width           =   1935
      End
      Begin VB.ComboBox cboIndeksi 
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
         Left            =   9480
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   1440
         Width           =   975
      End
      Begin VB.ComboBox cboKlasa 
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
         Left            =   8280
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   1440
         Width           =   1095
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
         Left            =   6240
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1440
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker data 
         Height          =   375
         Left            =   4200
         TabIndex        =   6
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
         Format          =   20185089
         UpDown          =   -1  'True
         CurrentDate     =   38308
      End
      Begin VB.TextBox txtVendlindja 
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
         Left            =   2160
         TabIndex        =   5
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox txtMemesia 
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
         Left            =   8280
         TabIndex        =   4
         Top             =   600
         Width           =   2175
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
         Left            =   6240
         TabIndex        =   3
         Top             =   600
         Width           =   1935
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
         Width           =   1935
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
         Left            =   4200
         TabIndex        =   2
         Top             =   600
         Width           =   1935
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
         Left            =   2160
         TabIndex        =   1
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Indeksi"
         Height          =   255
         Left            =   9480
         TabIndex        =   21
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblVitiShkollor 
         Caption         =   "Viti Shkollor"
         Height          =   255
         Left            =   6240
         TabIndex        =   20
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblDatelindja 
         Caption         =   "Datelindja"
         Height          =   255
         Left            =   4200
         TabIndex        =   17
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lblVendlindja 
         Caption         =   "Vendlindja"
         Height          =   255
         Left            =   2160
         TabIndex        =   16
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label lblSeksi 
         Caption         =   "Seksi"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label lblMemesia 
         Caption         =   "Memesia"
         Height          =   255
         Left            =   8280
         TabIndex        =   14
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblAtesia 
         Caption         =   "Atesia"
         Height          =   255
         Left            =   6240
         TabIndex        =   13
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblAmzaNo 
         Caption         =   "Numri Amzes"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblKlasa 
         Caption         =   "Klasa"
         Height          =   255
         Left            =   8280
         TabIndex        =   11
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lblMbiemri 
         Caption         =   "Mbiemri"
         Height          =   255
         Left            =   4200
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblEmri 
         Caption         =   "Emri"
         Height          =   255
         Left            =   2160
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
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
      Left            =   3360
      TabIndex        =   26
      Top             =   3000
      Visible         =   0   'False
      Width           =   1095
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
      Left            =   2520
      TabIndex        =   25
      Top             =   3000
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Notat ne provimet e"
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
      Left            =   120
      TabIndex        =   23
      Top             =   3000
      Visible         =   0   'False
      Width           =   2385
   End
End
Attribute VB_Name = "frmHedhjeGjeneralitete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim objektGabimi As New clsErrorHandler
Dim objUIController As clsUIController



Private Sub cmdDil_Click()
  Unload Me
  Set active_form = Nothing
End Sub


Private Sub cmdEmail_Click()
    If Not OpenEmailProgram(email) Then
        MsgBox "Ju nuk keni asnje program te instaluar per te derguar email", vbExclamation, "Gabim ne dergimin e emailit"
    End If

End Sub

Private Sub cmdOK_Click()
   cmdOK.Enabled = False
   Dim data             As String
   data = date
   objektGabimi.kapGabimin
   objektGabimi.menazhim_gabimi
   If objektGabimi.mvarGabimi = 0 Then
      objectInitialization HEDHJE_GJENERALITETE
      If data_modifikimi = "" Or data <> data_modifikimi Then
         objectInitialization MODIFIKIMI_I_DATES
      End If
   End If
   cmdOK.Enabled = True
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

Private Sub Command2_Click()
objectInitialization HEDHJE_TE_DHENASH_PROVIME
End Sub

Private Sub Form_Load()

  loadForm Me
  Image1.Picture = LoadPicture(adresaLogo)
  lblShkolla.Caption = emerShkolla
  mbushKomboBox
  percaktoRendinSipasTabit
  percaktoTeDrejtat
End Sub

Private Sub Form_Resize()
   'fraKerkim.Width = Me.Width - 300
  ' SGGrid1.Width = Me.Width - 300
  ' SGGrid2.Width = Me.Width - 300
   'rtbShenime.Width = Me.Width - 300
    fraInfo.Left = Me.Left + Me.Width - fraInfo.Width - 200
End Sub

Private Sub initializeGrid()
  
 ' initGride
  
 ' createUIController HEDHJE_TE_DHENASH_GJENERALITETE_FORM_LOAD
  
  'destroyUIController
 
End Sub

Private Sub pastroFushat()
   txtAmzaNo.Text = ""
   txtAtesia.Text = ""
   txtEmri.Text = ""
   txtKlasa.Text = ""
   txtMbiemri.Text = ""
   txtMemesia.Text = ""
   txtSeksi.Text = ""
   txtVendlindja.Text = ""
   txtVendlindja.Text = ""
End Sub

Private Sub initGride()
  ' SGGrid1.Columns.RemoveAll False
  ' SGGrid1.Rows.RemoveAll False
  ' SGGrid1.DataRowCount = 1
  '
  ' SGGrid2.Columns.RemoveAll False
  ' SGGrid2.Rows.RemoveAll False
  ' SGGrid2.DataRowCount = 1
End Sub


Private Function validate() As Boolean

  validate = True
  
  If Trim(txtAmzaNo.Text) = "" Then
     validate = False
     Exit Function
  End If
  
  If Trim(txtKlasa.Text) = "" Then
     validate = False
     Exit Function
  End If
  
  If Trim(txtEmri.Text) = "" Then
     validate = False
     Exit Function
  End If
  
  If Trim(txtMbiemri.Text) = "" Then
     validate = False
     Exit Function
  End If
  
End Function

Private Function createUIController(actionName As GUI_ACTION__ENUM)
   Set objUIController = New clsUIController
   objUIController.actionName = actionName
   objUIController.ExecuteActions False
End Function

Private Function destroyUIController()
   Set objUIController = Nothing
End Function

Private Sub txtKlasa_LostFocus()
 ' If Trim(txtKlasa.Text) <> "" Then
  '   Me.SGGrid1.CellAt(0, 0) = Trim$(txtKlasa.Text)
  'End If
End Sub

'Private Sub inicializoGridaAmza()
'
'    gridaAmza.Width = 14575
'    gridaAmza.Rows = 6
'    gridaAmza.Cols = 6
'    gridaAmza.FixedRows = 1
'    gridaAmza.FixedCols = 0
'    gridaAmza.row = 0
''    gridaAmza.col = 0
'    gridaAmza.ColWidth(0) = 1500
'    gridaAmza.Text = "Klasa"'
'    gridaAmza.col = 1
'    gridaAmza.ColWidth(1) = 1500
'    gridaAmza.Text = "Viti shkollor"
'    gridaAmza.col = 2
'    gridaAmza.ColWidth(2) = 1500
'    gridaAmza.Text = "Matematike"
'    gridaAmza.col = 3
'    gridaAmza.ColWidth(3) = 1500
'    gridaAmza.Text = "Fizike"
'    gridaAmza.col = 4
'    gridaAmza.ColWidth(4) = 1500
'    gridaAmza.Text = "Letersi"
'    gridaAmza.col = 5
'    gridaAmza.ColWidth(5) = 7000
'    gridaAmza.Text = "Vrejtje"
    
'End Sub

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
    data.TabIndex = 7
    cboVitiShkollor.TabIndex = 8
    cboKlasa.TabIndex = 9
    cboIndeksi.TabIndex = 10
    cmdOK.TabIndex = 11
End Sub


Private Sub percaktoTeDrejtat()
    
    Select Case statusi
        
        Case "SupervizorEmesme"
            mbushKomboMesme
        
        Case "SupervizorTetevjecare"
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

Private Sub txtData_GotFocus()
    txtData.Visible = False
End Sub
