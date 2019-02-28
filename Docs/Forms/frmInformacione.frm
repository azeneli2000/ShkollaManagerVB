VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmInformacione 
   Caption         =   "Informacione"
   ClientHeight    =   7890
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12630
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7890
   ScaleWidth      =   12630
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdDil 
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
      TabIndex        =   11
      Top             =   7080
      Width           =   2055
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
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7080
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
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7080
      Width           =   2535
   End
   Begin VB.Frame frmInfo 
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10815
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
         Left            =   1680
         TabIndex        =   13
         Top             =   720
         Width           =   4215
      End
      Begin VB.TextBox txtEmail 
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
         Left            =   1680
         TabIndex        =   10
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
         TabIndex        =   6
         Top             =   2760
         Width           =   1575
      End
      Begin VB.TextBox txtWebsite 
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
         Left            =   1680
         TabIndex        =   5
         Top             =   2280
         Width           =   4215
      End
      Begin VB.TextBox txtAdresa 
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
         Left            =   1680
         TabIndex        =   2
         Top             =   1440
         Width           =   4215
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
         Left            =   360
         TabIndex        =   12
         Top             =   720
         Width           =   1095
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
         Left            =   360
         TabIndex        =   9
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Image imgLogo 
         BorderStyle     =   1  'Fixed Single
         Height          =   1695
         Left            =   7680
         Stretch         =   -1  'True
         Top             =   480
         Width           =   2175
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
         Left            =   360
         TabIndex        =   4
         Top             =   2280
         Width           =   1215
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
         TabIndex        =   3
         Top             =   480
         Width           =   855
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
         Left            =   360
         TabIndex        =   1
         Top             =   1440
         Width           =   1215
      End
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
   cdlg.FileName = ""
   cdlg.CancelError = True
   
   On Error GoTo dil
   
   cdlg.Filter = s
   cdlg.ShowOpen
   MerrFile = cdlg.FileName
dil:
End Function

Private Sub cmdDil_Click()
    CallHelp indeksHelp
End Sub

Private Sub cmdOK_Click()
    objectInitialization INFORMACIONE_MBI_SHKOLLEN
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
End Sub
