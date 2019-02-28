VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmInformacione 
   Caption         =   "Informacione"
   ClientHeight    =   9555
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12630
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9555
   ScaleWidth      =   12630
   WindowState     =   2  'Maximized
   Begin VB.Frame frmInfo 
      Height          =   6255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   10815
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
         TabIndex        =   19
         Top             =   720
         Width           =   4215
      End
      Begin VB.TextBox txtTelefoni 
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
         TabIndex        =   18
         Top             =   4320
         Width           =   4215
      End
      Begin VB.TextBox txtAdresa 
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
         Top             =   1320
         Width           =   4215
      End
      Begin VB.TextBox txtWebsite 
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
         TabIndex        =   8
         Top             =   3120
         Width           =   4215
      End
      Begin VB.CommandButton cmdZgjidhLogo 
         BackColor       =   &H80000009&
         Caption         =   "Zgjidh Logo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8040
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2760
         Width           =   1575
      End
      Begin VB.TextBox txtEmail 
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
         TabIndex        =   6
         Top             =   3720
         Width           =   4215
      End
      Begin VB.TextBox txtQyteti 
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
         TabIndex        =   5
         Top             =   1920
         Width           =   4215
      End
      Begin VB.TextBox txtRrethi 
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
         TabIndex        =   4
         Top             =   2520
         Width           =   4215
      End
      Begin VB.Label lblEmail 
         Caption         =   "Nr. Telefoni:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   17
         Top             =   4320
         Width           =   1335
      End
      Begin VB.Label lblEmail 
         Caption         =   "Email :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   16
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Label lblAdresa 
         Caption         =   "Adresa :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label lblWebSite 
         Caption         =   "Website :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Image imgLogo 
         BorderStyle     =   1  'Fixed Single
         Height          =   1695
         Left            =   7680
         Stretch         =   -1  'True
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Emri  :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Qyteti :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Rrethi :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Logo :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   6720
         TabIndex        =   10
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdDil 
      BackColor       =   &H80000009&
      DisabledPicture =   "frmInformacione.frx":0000
      DownPicture     =   "frmInformacione.frx":353A
      Height          =   375
      Left            =   10920
      Picture         =   "frmInformacione.frx":6A74
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9120
      Width           =   2535
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H80000009&
      Default         =   -1  'True
      DisabledPicture =   "frmInformacione.frx":9FAE
      DownPicture     =   "frmInformacione.frx":103F0
      Height          =   375
      Left            =   3000
      Picture         =   "frmInformacione.frx":16832
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9120
      Width           =   2535
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   1440
      Top             =   6720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdDalje 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      DisabledPicture =   "frmInformacione.frx":1CC74
      DownPicture     =   "frmInformacione.frx":230B6
      Height          =   375
      Left            =   6840
      Picture         =   "frmInformacione.frx":294F8
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9120
      Width           =   2535
   End
End
Attribute VB_Name = "frmInformacione"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdDalje_Click()
  Unload Me
  Set active_form = Nothing
End Sub


Function MerrFile(cdlg As CommonDialog, s As String) As String
   MerrFile = ""
   cdlg.fileName = ""
   cdlg.CancelError = True
   
   On Error GoTo dil
   
   cdlg.Filter = s
   cdlg.ShowOpen
   MerrFile = cdlg.fileName
dil:
End Function

Private Sub cmdDil_Click()
    CallHelp indeksHelp
End Sub

Private Sub cmdOK_Click()
   Dim data             As String
   data = Date
   objectInitialization INFORMACIONE_MBI_SHKOLLEN
   If data_modifikimi = "" Or data <> data_modifikimi Then
      objectInitialization MODIFIKIMI_I_DATES
   End If
End Sub

Private Sub cmdZgjidhLogo_Click()
   Dim s                As String, s1 As String

   s = MerrFile(cdlg, "Figura|*.bmp;*.ico;*.cur;*.jpg;*.gif")
   adresaLogo = s
   If s = "" Then Exit Sub
   Set imgLogo.Picture = LoadPicture(s)
End Sub

Private Sub objectInitialization(actionName As GUI_ACTION__ENUM)
   Set objUIController = New clsUIController

   objUIController.actionName = actionName
   objUIController.ExecuteActions

   Set objUIController = Nothing
End Sub



Private Sub Form_Load()
    objectInitialization SHFAQ_INFORMACION_MBI_SHKOLLEN
    imgLogo.Picture = LoadPicture(adresaLogo)
    percaktoRendinSipasTabit
End Sub

Private Sub percaktoRendinSipasTabit()
        
    txtEmri.TabIndex = 0
    txtAdresa.TabIndex = 1
    txtQyteti.TabIndex = 2
    txtRrethi.TabIndex = 3
    txtWebsite.TabIndex = 4
    txtEmail.TabIndex = 5
    Me.txtTelefoni.TabIndex = 6
    cmdZgjidhLogo.TabIndex = 7
    cmdOK.TabIndex = 8
    cmdDalje.TabIndex = 9
    cmdDil.TabIndex = 10
    
End Sub

Private Sub lblTelNr_Click(Index As Integer)

End Sub
