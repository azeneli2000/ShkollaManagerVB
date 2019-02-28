VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmShenimeTePerkohshme 
   Caption         =   "Shenime te perkohshme"
   ClientHeight    =   9300
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14505
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9300
   ScaleWidth      =   14505
   WindowState     =   2  'Maximized
   Begin RichTextLib.RichTextBox rtbSjellja3 
      Height          =   1335
      Left            =   7560
      TabIndex        =   25
      Top             =   6600
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   2355
      _Version        =   393217
      TextRTF         =   $"frmShenimeTePerkohshme.frx":0000
   End
   Begin RichTextLib.RichTextBox rtbSjellja2 
      Height          =   1335
      Left            =   7560
      TabIndex        =   24
      Top             =   4920
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   2355
      _Version        =   393217
      TextRTF         =   $"frmShenimeTePerkohshme.frx":0082
   End
   Begin RichTextLib.RichTextBox rtbSjellja1 
      Height          =   1335
      Left            =   7560
      TabIndex        =   23
      Top             =   3240
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   2355
      _Version        =   393217
      TextRTF         =   $"frmShenimeTePerkohshme.frx":0104
   End
   Begin VB.CommandButton cmdDil 
      BackColor       =   &H80000009&
      DisabledPicture =   "frmShenimeTePerkohshme.frx":0186
      DownPicture     =   "frmShenimeTePerkohshme.frx":65C8
      Height          =   375
      Left            =   7440
      Picture         =   "frmShenimeTePerkohshme.frx":CA0A
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   8640
      Width           =   2535
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H80000009&
      DisabledPicture =   "frmShenimeTePerkohshme.frx":12E4C
      DownPicture     =   "frmShenimeTePerkohshme.frx":1928E
      Height          =   375
      Left            =   4080
      Picture         =   "frmShenimeTePerkohshme.frx":1F6D0
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   8640
      Width           =   2535
   End
   Begin VB.CommandButton cmdNdihme 
      BackColor       =   &H00FFFFFF&
      DisabledPicture =   "frmShenimeTePerkohshme.frx":25B12
      DownPicture     =   "frmShenimeTePerkohshme.frx":2904C
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
      Left            =   11040
      Picture         =   "frmShenimeTePerkohshme.frx":2C586
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   8640
      Width           =   2535
   End
   Begin VB.Frame fraInfo 
      Height          =   1455
      Left            =   12360
      TabIndex        =   11
      Top             =   0
      Width           =   2775
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   15
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label lblWebsite 
         Height          =   375
         Left            =   1560
         TabIndex        =   14
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   120
         Picture         =   "frmShenimeTePerkohshme.frx":2FAC0
         Stretch         =   -1  'True
         Top             =   600
         Width           =   960
      End
   End
   Begin VB.Frame fraKerkimi 
      Caption         =   "Kerko sipas"
      ForeColor       =   &H00008000&
      Height          =   1935
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   7335
      Begin VB.TextBox txtVitiShkollor 
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
         Height          =   315
         Left            =   240
         TabIndex        =   16
         Top             =   720
         Width           =   2000
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
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   720
         Width           =   2000
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
         Height          =   315
         Left            =   4800
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   720
         Width           =   2000
      End
      Begin VB.CommandButton cmdKerko 
         BackColor       =   &H80000009&
         DisabledPicture =   "frmShenimeTePerkohshme.frx":36A35
         DownPicture     =   "frmShenimeTePerkohshme.frx":3CE77
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
         Left            =   4800
         Picture         =   "frmShenimeTePerkohshme.frx":432B9
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1320
         Width           =   2000
      End
      Begin VB.Label Label1 
         Caption         =   "Viti shkollor"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Klasa"
         Height          =   255
         Left            =   2400
         TabIndex        =   7
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Indeksi"
         Height          =   255
         Left            =   4800
         TabIndex        =   6
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.ListBox lista 
      Appearance      =   0  'Flat
      Height          =   5100
      Left            =   360
      TabIndex        =   1
      Top             =   2880
      Width           =   4095
   End
   Begin VB.ListBox listAmza 
      Appearance      =   0  'Flat
      Height          =   4905
      Left            =   4920
      TabIndex        =   0
      Top             =   2760
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Shënim i përkohshëm për mesin e semestrit të dytë"
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   7560
      TabIndex        =   19
      Top             =   6360
      Width           =   3735
   End
   Begin VB.Label Label62 
      Caption         =   "Shënim i përkohshëm për fundin e semestrit të parë"
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   7560
      TabIndex        =   18
      Top             =   4680
      Width           =   4095
   End
   Begin VB.Label Label61 
      Caption         =   "Shënim i përkohshëm për mesin e semestrit të parë"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   7560
      TabIndex        =   17
      Top             =   3000
      Width           =   3735
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Emrat e nxenesve"
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
      Left            =   360
      TabIndex        =   10
      Top             =   2520
      Width           =   4095
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Shënimet për nxënësin sipas llojit"
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
      Left            =   7560
      TabIndex        =   9
      Top             =   2520
      Width           =   4455
   End
End
Attribute VB_Name = "frmShenimeTePerkohshme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
    rtbSjellja1.Text = ""
    rtbSjellja2.Text = ""
    rtbSjellja3.Text = ""
    lista.Clear
    Me.listAmza.Clear
    objectInitialization HEDHJE_SHENIM_PERKOHSHEM_KERKO
End Sub

Private Sub cmdOK_Click()
   Dim data             As String
   data = date
   objectInitialization SHENIME_PERKOHSHME_HIDH_OK
   If data_modifikimi = "" Or data <> data_modifikimi Then
      objectInitialization MODIFIKIMI_I_DATES
   End If
   rtbSjellja1.Text = ""
   rtbSjellja2.Text = ""
   rtbSjellja3.Text = ""
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
  mbushKomboBox
  viti
  lista.ListIndex = -1
  percaktoTeDrejtat
  'cmdOK.Enabled = False
End Sub


Private Sub viti()
   Dim d, muai, viti    As String
   d = Now
   Dim M, v             As Integer
   muai = DateTime.Month(DateTime.Now)
   viti = DateTime.Year(DateTime.date)
   M = Val(muai)
   v = Val(viti)
   Me.txtVitiShkollor.Text = gjej_vitin(M, v)

End Sub
Private Sub mbushKomboBox()
    
    cboIndeksi.AddItem "A"
    cboIndeksi.AddItem "B"
    cboIndeksi.AddItem "C"
    cboIndeksi.AddItem "D"
    cboIndeksi.AddItem "E"
    cboIndeksi.AddItem "F"
  
End Sub

Private Sub percaktoTeDrejtat()
    
    Select Case statusi
        
        Case "SupervizorEmesme"
            mbushKomboMesme
        
        Case "SupervizorTetevjecare"
            mbushKomboTetevjecare
        
        Case "Administrator"
            mbushKomboAdministrator
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
Private Sub mbushKomboAdministrator()
   cboKlasa.Clear
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
Private Sub objectInitialization(actionName As GUI_ACTION__ENUM)
   Set objUIController = New clsUIController

   objUIController.actionName = actionName
   objUIController.ExecuteActions

   Set objUIController = Nothing
End Sub
