VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmEleminoNxenes 
   Caption         =   "Elemino nxenes"
   ClientHeight    =   9780
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14820
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9780
   ScaleWidth      =   14820
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdHelp 
      BackColor       =   &H00FFFFFF&
      DisabledPicture =   "frmEleminoNxenes.frx":0000
      DownPicture     =   "frmEleminoNxenes.frx":353A
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
      Left            =   10560
      Picture         =   "frmEleminoNxenes.frx":6A74
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   9120
      Width           =   2535
   End
   Begin VB.TextBox txtDitet 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Left            =   8100
      Locked          =   -1  'True
      TabIndex        =   37
      Text            =   "  Hen     Mar      Mer      Enjt     Pre      Shtu     Diel"
      Top             =   900
      Width           =   3700
   End
   Begin VB.CommandButton cmdDataSot 
      BackColor       =   &H00FFFFFF&
      DisabledPicture =   "frmEleminoNxenes.frx":9FAE
      DownPicture     =   "frmEleminoNxenes.frx":103F0
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
      Left            =   8640
      Picture         =   "frmEleminoNxenes.frx":16832
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   3240
      Width           =   1935
   End
   Begin VB.ComboBox cboMuaji 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9560
      TabIndex        =   35
      Top             =   555
      Width           =   1455
   End
   Begin VB.TextBox txtMuaji 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8160
      TabIndex        =   34
      Top             =   555
      Width           =   1335
   End
   Begin VB.Frame fraInfo 
      Height          =   1695
      Left            =   12000
      TabIndex        =   28
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
         TabIndex        =   30
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
         TabIndex        =   29
         Top             =   960
         Width           =   1095
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   120
         Picture         =   "frmEleminoNxenes.frx":1CC74
         Stretch         =   -1  'True
         Top             =   600
         Width           =   960
      End
      Begin VB.Label lblWebsite 
         Height          =   375
         Left            =   1560
         TabIndex        =   32
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
         TabIndex        =   31
         Top             =   120
         Width           =   2415
      End
   End
   Begin VB.TextBox txtDatelindja 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   5160
      Width           =   2535
   End
   Begin VB.TextBox txtVendlindja 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   4320
      Width           =   2535
   End
   Begin VB.TextBox txtSeksi 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   3480
      Width           =   2535
   End
   Begin VB.TextBox txtMemesia 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   6000
      Width           =   2535
   End
   Begin VB.TextBox txtVitiShkollor 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   5160
      Width           =   2535
   End
   Begin VB.TextBox txtIndeksi 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   4320
      Width           =   2535
   End
   Begin VB.TextBox txtKlasa 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   3480
      Width           =   2535
   End
   Begin VB.Frame frmIdentifiko 
      Caption         =   "Te dhenat identifikuese"
      ForeColor       =   &H00008000&
      Height          =   2655
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   7695
      Begin VB.TextBox txtAtesia 
         BackColor       =   &H80000003&
         ForeColor       =   &H80000004&
         Height          =   375
         Left            =   5760
         TabIndex        =   39
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CommandButton cmdKerko 
         BackColor       =   &H00FFFFFF&
         DisabledPicture =   "frmEleminoNxenes.frx":23BE9
         DownPicture     =   "frmEleminoNxenes.frx":2A02B
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
         Left            =   4920
         Picture         =   "frmEleminoNxenes.frx":3046D
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   600
         Width           =   2175
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
         Left            =   4080
         TabIndex        =   8
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
         Left            =   2400
         TabIndex        =   7
         Top             =   1440
         Width           =   1575
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
         Height          =   360
         Left            =   2400
         TabIndex        =   5
         Top             =   645
         Width           =   1575
      End
      Begin VB.OptionButton optTetevjecare 
         Caption         =   "Shkolla nëntëvjeçare"
         ForeColor       =   &H000000C0&
         Height          =   855
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   1935
      End
      Begin VB.OptionButton optMesme 
         Caption         =   "Shkolla e mesme"
         ForeColor       =   &H00C00000&
         Height          =   615
         Left            =   240
         TabIndex        =   3
         Top             =   550
         Width           =   1935
      End
      Begin VB.Frame fraCikli 
         Caption         =   "Cikli"
         ForeColor       =   &H00008000&
         Height          =   2175
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label7 
         Caption         =   "Atesia"
         Height          =   255
         Left            =   5760
         TabIndex        =   38
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Mbiemri"
         Height          =   255
         Left            =   4080
         TabIndex        =   10
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Emri "
         Height          =   255
         Left            =   2400
         TabIndex        =   9
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Numri i amzes"
         Height          =   255
         Left            =   2400
         TabIndex        =   6
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdDil 
      BackColor       =   &H80000009&
      DisabledPicture =   "frmEleminoNxenes.frx":368AF
      DownPicture     =   "frmEleminoNxenes.frx":3CCF1
      Height          =   375
      Left            =   6720
      Picture         =   "frmEleminoNxenes.frx":43133
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9120
      Width           =   2415
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H80000009&
      DisabledPicture =   "frmEleminoNxenes.frx":49575
      DownPicture     =   "frmEleminoNxenes.frx":4F9B7
      Height          =   375
      Left            =   3240
      Picture         =   "frmEleminoNxenes.frx":55DF9
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9120
      Width           =   2175
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   2535
      Left            =   8040
      TabIndex        =   33
      Top             =   480
      Width           =   3855
      _Version        =   524288
      _ExtentX        =   6800
      _ExtentY        =   4471
      _StockProps     =   1
      BackColor       =   16777152
      Year            =   2005
      Month           =   3
      Day             =   20
      DayLength       =   1
      MonthLength     =   1
      DayFontColor    =   0
      FirstDay        =   7
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
   Begin VB.Label lblDataLargimi 
      Alignment       =   2  'Center
      Caption         =   "Data e largimit"
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
      Left            =   8400
      TabIndex        =   26
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label11 
      Caption         =   "Datelindja"
      Height          =   255
      Left            =   3840
      TabIndex        =   24
      Top             =   4920
      Width           =   2055
   End
   Begin VB.Label Label10 
      Caption         =   "Vendlindja"
      Height          =   255
      Left            =   3840
      TabIndex        =   22
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label Label9 
      Caption         =   "Seksi"
      Height          =   255
      Left            =   3840
      TabIndex        =   20
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label8 
      Caption         =   "Memesia"
      Height          =   255
      Left            =   840
      TabIndex        =   18
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "Viti shkollor"
      Height          =   255
      Left            =   840
      TabIndex        =   16
      Top             =   4920
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Indeksi"
      Height          =   255
      Left            =   840
      TabIndex        =   14
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Klasa"
      Height          =   255
      Left            =   840
      TabIndex        =   12
      Top             =   3240
      Width           =   1575
   End
End
Attribute VB_Name = "frmEleminoNxenes"
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


Private Sub cmdEmail_Click()
    If Not OpenEmailProgram(email) Then
        MsgBox "Ju nuk keni asnje program te instaluar per te derguar email", vbExclamation, "Gabim ne dergimin e emailit"
    End If
End Sub

Private Sub cmdHelp_Click()
CallHelp indeksHelp
End Sub

Private Sub cmdKerko_Click()
   cmdKerko.Enabled = False
   objektGabimi.kapGabimin
   objektGabimi.menazhim_gabimi
   If objektGabimi.mvarGabimi = 0 Then
      objectInitialization ELIMINO_NXENES_KERKO
   End If
   cmdKerko.Enabled = True
End Sub

Private Sub cmdOK_Click()
   Dim data As String
   data = date
   objektGabimi.kapGabimin
   objektGabimi.menazhim_gabimi
   If objektGabimi.mvarGabimi = 0 Then
      objectInitialization ELIMINO_NXENES_LARGO
      If data_modifikimi = "" Or data <> data_modifikimi Then
         objectInitialization MODIFIKIMI_I_DATES
      End If
   End If
   cmdOK.Enabled = False
End Sub

Private Sub cmdWebsiste_Click()
    If website <> "" Then
        GoToWeb website
    Else
        MsgBox "Ju nuk e keni dhene adresen e faqes tuaj te web-it.", vbInformation
    End If
End Sub

Private Sub Form_Load()
    loadForm Me
    Image1.Picture = LoadPicture(adresaLogo)
    lblShkolla.Caption = emerShkolla
    Calendar1.Today
    perkthe_muajin
    mbush_kombo_muaji
    cboMuaji.Locked = True
    percaktoRendinSipasTabit
    cmdOK.Enabled = False
    percaktoTeDrejtat
End Sub

Private Sub objectInitialization(actionName As GUI_ACTION__ENUM)
   Set objUIController = New clsUIController

   objUIController.actionName = actionName
   objUIController.ExecuteActions

   Set objUIController = Nothing
End Sub
Private Sub Form_Resize()

   
   fraInfo.Left = Me.Left + Me.Width - fraInfo.Width - 200
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
            cboMuaji.Text = "Nentor"
        Case 11
            txtMuaji.Text = "Nentor"
            cboMuaji.Text = "Nentor"
        Case 12
            txtMuaji.Text = "Dhjetor"
            cboMuaji.Text = "Dhjetor"
    End Select
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


Private Sub percaktoRendinSipasTabit()
    txtAmzaNo.TabIndex = 0
    txtEmri.TabIndex = 1
    txtMbiemri.TabIndex = 2
    txtAtesia.TabIndex = 3
    cmdKerko.TabIndex = 4
End Sub

Private Sub percaktoTeDrejtat()
    
    Select Case statusi
        
        Case "SupervizorEmesme"
            optMesme.Visible = False
            optTetevjecare.Visible = False
            optMesme.Value = True
            optTetevjecare.Value = False
            
            fraCikli.Visible = False
            
        
        Case "SupervizorTetevjecare"
            optMesme.Visible = False
            optTetevjecare.Visible = False
            optMesme.Value = False
            optTetevjecare.Value = True
            
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
    txtSeksi.Text = ""
    txtIndeksi.Text = ""
    txtVendlindja.Text = ""
    txtVitiShkollor.Text = ""
    txtDatelindja.Text = ""
    txtMemesia.Text = ""
    
End Sub

Private Sub optMesme_Click()
    Pastro
End Sub

Private Sub optTetevjecare_Click()
    Pastro
End Sub
