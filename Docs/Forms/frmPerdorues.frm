VERSION 5.00
Begin VB.Form frmPerdorues 
   Caption         =   "Perdoruesi"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7500
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame fraInfo 
      Height          =   1575
      Left            =   12360
      TabIndex        =   21
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
         TabIndex        =   23
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
         TabIndex        =   22
         Top             =   960
         Width           =   1095
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   120
         Picture         =   "frmPerdorues.frx":0000
         Stretch         =   -1  'True
         Top             =   600
         Width           =   960
      End
      Begin VB.Label lblWebsite 
         Height          =   375
         Left            =   1560
         TabIndex        =   25
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
         TabIndex        =   24
         Top             =   120
         Width           =   2415
      End
   End
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
      TabIndex        =   9
      Top             =   6960
      Width           =   2055
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
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6960
      Width           =   2535
   End
   Begin VB.Frame frmUser 
      Height          =   5055
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   10095
      Begin VB.TextBox txtPerdorues 
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
         TabIndex        =   12
         Top             =   600
         Width           =   2055
      End
      Begin VB.CommandButton cmdVerifiko 
         BackColor       =   &H80000009&
         Caption         =   "Verifiko"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3240
         Width           =   2535
      End
      Begin VB.ComboBox txtUserName 
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
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox Text3 
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
         IMEMode         =   3  'DISABLE
         Left            =   7440
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   2040
         Width           =   2055
      End
      Begin VB.TextBox Text2 
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
         Left            =   7440
         TabIndex        =   7
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox Text1 
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
         IMEMode         =   3  'DISABLE
         Left            =   7440
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   1320
         Width           =   2055
      End
      Begin VB.TextBox txtPassword1 
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
         IMEMode         =   3  'DISABLE
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   2040
         Width           =   2055
      End
      Begin VB.CommandButton cmdModifikoUser 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Modifiko "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3240
         Width           =   1335
      End
      Begin VB.CommandButton cmdKrijoUser 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Krijo "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   3240
         Width           =   2415
      End
      Begin VB.TextBox txtPassword 
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
         IMEMode         =   3  'DISABLE
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Frame fraModifiko 
         Caption         =   "Modifiko perdorues :"
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
         Height          =   3855
         Left            =   360
         TabIndex        =   13
         Top             =   240
         Width           =   3735
         Begin VB.CommandButton cmdElimino 
            BackColor       =   &H80000009&
            Caption         =   "Elimino"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2040
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   3000
            Width           =   1215
         End
         Begin VB.Label lbPassword1 
            Caption         =   "Perserit fjalekalimin :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   495
            Left            =   120
            TabIndex        =   16
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label lbPassword 
            Caption         =   "Fjalekalimi :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   375
            Left            =   120
            TabIndex        =   15
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label Label4 
            Caption         =   "Perdorues :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   375
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame fraKrijo 
         Caption         =   "Krijo perdorues :"
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
         Height          =   3855
         Left            =   5640
         TabIndex        =   17
         Top             =   240
         Width           =   4095
         Begin VB.Label Label3 
            Caption         =   "Perserit fjalekalimin :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   495
            Left            =   360
            TabIndex        =   20
            Top             =   1680
            Width           =   1455
         End
         Begin VB.Label Label2 
            Caption         =   "Fjalekalimi :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   375
            Left            =   360
            TabIndex        =   19
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Perdorues :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   375
            Left            =   360
            TabIndex        =   18
            Top             =   360
            Width           =   1455
         End
      End
   End
End
Attribute VB_Name = "frmPerdorues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdDalje_Click()
  Unload Me
  Set active_form = Nothing
End Sub

Private Sub cmdDil_Click()
    CallHelp indeksHelp
End Sub

Private Sub cmdElimino_Click()

   objectInitialization ELIMINIMI_I_PERDORUESIT
   txtUserName.Clear
   objectInitialization KONFIGURIME_PERDORUES_KERKO
   txtPerdorues.Text = ""
   txtPassword.Text = ""
   txtPassword1.Text = ""
   txtPerdorues.Visible = False
   txtUserName.Visible = True
   txtPassword1.Visible = False
   lbPassword1.Visible = False
   cmdVerifiko.Visible = True
   cmdModifikoUser.Visible = False
End Sub

Private Sub cmdKrijoUser_Click()

   If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Then
      MsgBox "Ju duhet te jepni te gjitha tedhenat e kerkuara para se te krijoni nje perdorues te ri.", , "Konfigurimi i perdoruesve"
   Else
      Dim kodi1            As String
      Dim kodi2            As String
      kodi1 = Text1.Text
      kodi2 = Text3.Text
      If kodi1 = kodi2 Then
         objectInitialization KONFIGURIME_PERDORUES_I_RI
         txtUserName.Clear
         objectInitialization KONFIGURIME_PERDORUES_KERKO
      Else
         MsgBox " Fjalekalimi i perseritur nuk eshte i njejte me ate qe dhate heren e pare.Jepni perseri fjalekalimin.", , "Konfigurimi i perdoruesve."
         Text1.Text = ""
         Text3.Text = ""
      End If
   End If
End Sub

Private Sub cmdModifikoUser_Click()

   If txtPerdorues.Text = "" Or txtPassword.Text = "" Or txtPassword1.Text = "" Then
      MsgBox "Ju duhet te zgjidhni nje perdorues ekzistent dhe te jepni fjalekalimin e tij per te patur te mundur modifikimin e ketij perdoruesi.", , "Konfigurimi i perdoruesve."
   Else
      Dim kodi1            As String
      Dim kodi2            As String
      kodi1 = txtPassword.Text
      kodi2 = txtPassword1.Text
      If kodi1 = kodi2 Then
         objectInitialization KONFIGURIME_PERDORUES_MODIFIKO

         MsgBox "Perdoruesi u modifikua."
         txtPerdorues.Text = ""
         txtPassword.Text = ""
         txtPassword1.Text = ""
         txtPerdorues.Visible = False
         txtUserName.Visible = True
         txtPassword1.Visible = False
         lbPassword1.Visible = False
         cmdVerifiko.Visible = True
         cmdModifikoUser.Visible = False
      Else
         MsgBox " Fjalekalimi i perseritur nuk eshte i njejte me ate qe dhate heren e pare.Jepni perseri fjalekalimin.", , "Konfigurimi i perdoruesve."
         txtPassword.Text = ""
         txtPassword1.Text = ""
      End If
   End If
End Sub




Private Sub cmdVerifiko_Click()
   If txtUserName = "" Or txtPassword.Text = "" Then
      MsgBox "Ju duhet te jepni te gjitha te dhenat e kerkuara per te modifikuar njeperdorues te ri.", , "Konfigurimii perdoruesve."
   Else
      objectInitialization VERIFIKIMI_I_PERDORUESIT
   End If
End Sub



Private Sub Form_Load()
  loadForm Me
  Image1.Picture = LoadPicture(adresaLogo)
  lblShkolla.Caption = emerShkolla
  txtPerdorues.Visible = False
  txtPassword1.Visible = False
  lbPassword1.Visible = False
  objectInitialization KONFIGURIME_PERDORUES_KERKO
  
End Sub

Private Sub Form_Resize()
  fraInfo.Left = Me.Left + Me.Width - fraInfo.Width - 200
End Sub

Private Sub objectInitialization(actionName As GUI_ACTION__ENUM)
   Set objUIController = New clsUIController

   objUIController.actionName = actionName
   objUIController.ExecuteActions

   Set objUIController = Nothing
End Sub


Private Sub txtUserName_Click()
   ' txtEmri.Text = txtUserName.List(txtUserName.ListIndex)
End Sub
