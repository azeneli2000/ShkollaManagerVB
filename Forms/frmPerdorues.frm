VERSION 5.00
Begin VB.Form frmPerdorues 
   Caption         =   "Perdoruesi"
   ClientHeight    =   8610
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8610
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
         TabIndex        =   24
         Top             =   120
         Width           =   2415
      End
   End
   Begin VB.CommandButton cmdDil 
      BackColor       =   &H80000009&
      DisabledPicture =   "frmPerdorues.frx":6F75
      DownPicture     =   "frmPerdorues.frx":A4AF
      Height          =   375
      Left            =   9360
      Picture         =   "frmPerdorues.frx":D9E9
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9120
      Width           =   2535
   End
   Begin VB.CommandButton cmdDalje 
      BackColor       =   &H00FFFFFF&
      DisabledPicture =   "frmPerdorues.frx":10F23
      DownPicture     =   "frmPerdorues.frx":17365
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
      Left            =   5160
      Picture         =   "frmPerdorues.frx":1D7A7
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9120
      Width           =   2535
   End
   Begin VB.Frame frmUser 
      Height          =   5775
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   11055
      Begin VB.CommandButton cmdVerifiko 
         BackColor       =   &H80000009&
         DisabledPicture =   "frmPerdorues.frx":23BE9
         DownPicture     =   "frmPerdorues.frx":2A02B
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
         Left            =   1440
         Picture         =   "frmPerdorues.frx":3046D
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   4440
         Width           =   2295
      End
      Begin VB.CommandButton cmdElimino 
         BackColor       =   &H80000009&
         Caption         =   "Elemino"
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
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   4440
         Width           =   1095
      End
      Begin VB.TextBox txtPerdorues 
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
         Left            =   1920
         TabIndex        =   12
         Top             =   840
         Width           =   2055
      End
      Begin VB.ComboBox txtUserName 
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
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox Text3 
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
         IMEMode         =   3  'DISABLE
         Left            =   7440
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   2280
         Width           =   2055
      End
      Begin VB.TextBox Text2 
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
         Left            =   7440
         TabIndex        =   7
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox Text1 
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
         IMEMode         =   3  'DISABLE
         Left            =   7440
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox txtPassword1 
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
         IMEMode         =   3  'DISABLE
         Left            =   1920
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   2280
         Width           =   2055
      End
      Begin VB.CommandButton cmdModifikoUser 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Modifiko "
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
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   4440
         Width           =   1215
      End
      Begin VB.CommandButton cmdKrijoUser 
         BackColor       =   &H00FFFFFF&
         DisabledPicture =   "frmPerdorues.frx":368AF
         DownPicture     =   "frmPerdorues.frx":3CCF1
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
         Left            =   7440
         Picture         =   "frmPerdorues.frx":43133
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   4560
         Width           =   2055
      End
      Begin VB.TextBox txtPassword 
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
         IMEMode         =   3  'DISABLE
         Left            =   1920
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Frame fraModifiko 
         Caption         =   "Modifiko perdorues :"
         ForeColor       =   &H00008000&
         Height          =   5175
         Left            =   360
         TabIndex        =   13
         Top             =   240
         Width           =   4215
         Begin VB.ComboBox cboWorkstationModifiko 
            BackColor       =   &H80000003&
            ForeColor       =   &H80000004&
            Height          =   315
            ItemData        =   "frmPerdorues.frx":49575
            Left            =   1560
            List            =   "frmPerdorues.frx":49577
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   3360
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.ComboBox cboStatusiModifiko 
            BackColor       =   &H80000003&
            ForeColor       =   &H80000004&
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   2640
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Label lblWorkstationModifiko 
            Caption         =   "Vendi i hyrjes :"
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   240
            TabIndex        =   34
            Top             =   3375
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label Label5 
            Caption         =   "Statusi i ri :"
            ForeColor       =   &H00C00000&
            Height          =   375
            Left            =   240
            TabIndex        =   30
            Top             =   2640
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label lbPassword1 
            Caption         =   "Perserit fjalekalimin :"
            ForeColor       =   &H00C00000&
            Height          =   495
            Left            =   240
            TabIndex        =   16
            Top             =   1920
            Width           =   1215
         End
         Begin VB.Label lbPassword 
            Caption         =   "Fjalekalimi :"
            ForeColor       =   &H00C00000&
            Height          =   375
            Left            =   240
            TabIndex        =   15
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label4 
            Caption         =   "Perdorues :"
            ForeColor       =   &H00C00000&
            Height          =   375
            Left            =   240
            TabIndex        =   14
            Top             =   600
            Width           =   1335
         End
      End
      Begin VB.Frame fraKrijo 
         Caption         =   "Krijo perdorues :"
         ForeColor       =   &H00008000&
         Height          =   5175
         Left            =   5760
         TabIndex        =   17
         Top             =   240
         Width           =   4095
         Begin VB.ComboBox cboWorkstation 
            BackColor       =   &H80000003&
            ForeColor       =   &H80000004&
            Height          =   315
            ItemData        =   "frmPerdorues.frx":49579
            Left            =   1680
            List            =   "frmPerdorues.frx":4957B
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   3480
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.ComboBox cboStatusi 
            BackColor       =   &H80000003&
            ForeColor       =   &H80000004&
            Height          =   360
            ItemData        =   "frmPerdorues.frx":4957D
            Left            =   1680
            List            =   "frmPerdorues.frx":4957F
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   2760
            Width           =   2055
         End
         Begin VB.Label lblWorkstation 
            Caption         =   "Vendi i hyrjes :"
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   360
            TabIndex        =   31
            Top             =   3500
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label lblStatusi 
            Caption         =   "Statusi :"
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   360
            TabIndex        =   28
            Top             =   2760
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Perserit fjalekalimin :"
            ForeColor       =   &H00C00000&
            Height          =   495
            Left            =   360
            TabIndex        =   20
            Top             =   2040
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Fjalekalimi :"
            ForeColor       =   &H00C00000&
            Height          =   375
            Left            =   360
            TabIndex        =   19
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Perdorues :"
            ForeColor       =   &H00C00000&
            Height          =   375
            Left            =   360
            TabIndex        =   18
            Top             =   600
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



Private Sub cboStatusi_Click()
    If (cboStatusi.Text = "Vizitor") Then
        lblWorkstation.Visible = True
        cboWorkstation.Visible = True
    Else
        lblWorkstation.Visible = False
        cboWorkstation.Visible = False
    End If
End Sub


Private Sub cboStatusiModifiko_Click()
    If (cboStatusiModifiko.Text = "Vizitor") Then
        lblWorkstationModifiko.Visible = True
        cboWorkstationModifiko.Visible = True
    Else
        lblWorkstationModifiko.Visible = False
        cboWorkstationModifiko.Visible = False
    End If
End Sub

Private Sub cmdDalje_Click()
  Unload Me
  Set active_form = Nothing
End Sub

Private Sub cmdDil_Click()
    CallHelp indeksHelp
End Sub

Private Sub cmdElimino_Click()
   Dim data             As String
   data = date
   objectInitialization ELIMINIMI_I_PERDORUESIT
   If data_modifikimi = "" Or data <> data_modifikimi Then
      objectInitialization MODIFIKIMI_I_DATES
   End If
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
   cmdElimino.Visible = False
   cboStatusiModifiko.Visible = False
   Label5.Visible = False
   Me.lblWorkstationModifiko.Visible = False
   Me.cboWorkstationModifiko.Visible = False
End Sub

Private Sub cmdEmail_Click()
    If Not OpenEmailProgram(email) Then
        MsgBox "Ju nuk keni asnje program te instaluar per te derguar email", vbExclamation, "Gabim ne dergimin e emailit"
    End If
End Sub

Private Sub cmdKrijoUser_Click()

   If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or cboStatusi.Text = "" Then
      MsgBox "Ju duhet te jepni te gjitha te dhenat e kerkuara " & Chr(10) & "para se te krijoni nje perdorues te ri.", vbInformation, "Konfigurimi i perdoruesve"
   ElseIf cboStatusi.Text = "Vizitor" And cboWorkstation.Text = "" Then
      MsgBox "Për konfigurimin e një vizitori përveç të dhënave të tjera," & Chr(10) & "duhet të caktoni edhe emrin e kompjuterit prej të cilit do të hyjë në program!", vbInformation, "Konfigurimi i perdoruesve"
   Else
      Dim kodi1            As String
      Dim kodi2            As String
      Dim data             As String
      data = date
      kodi1 = Text1.Text
      kodi2 = Text3.Text
      If kodi1 = kodi2 Then
         objectInitialization KONFIGURIME_PERDORUES_I_RI
         txtUserName.Clear
         objectInitialization KONFIGURIME_PERDORUES_KERKO
         If data_modifikimi = "" Or data <> data_modifikimi Then
            objectInitialization MODIFIKIMI_I_DATES
         End If
      Else
         MsgBox " Fjalekalimi i perseritur nuk eshte i njejte me ate qe dhate heren e pare." & Chr(10) & "Jepni perseri fjalekalimin.", vbInformation, "Konfigurimi i perdoruesve."
         Text1.Text = ""
         Text3.Text = ""
      End If
   End If
End Sub

Private Sub cmdModifikoUser_Click()

   If txtPerdorues.Text = "" Or txtPassword.Text = "" Or txtPassword1.Text = "" Then
      MsgBox "Ju duhet te zgjidhni nje perdorues ekzistent dhe te jepni fjalekalimin e tij " & Chr(10) & "per te patur te mundur modifikimin e ketij perdoruesi.", vbInformation, "Konfigurimi i perdoruesve."
   ElseIf cboStatusiModifiko.Text = "Vizitor" And cboWorkstationModifiko.Text = "" Then
      MsgBox "Për konfigurimin e një vizitori përveç të dhënave të tjera," & Chr(10) & "duhet të caktoni edhe emrin e kompjuterit prej të cilit do të hyjë në program!", vbInformation, "Konfigurimi i perdoruesve"
   Else
      Dim kodi1            As String
      Dim kodi2            As String
      Dim data As String
      data = date
      kodi1 = txtPassword.Text
      kodi2 = txtPassword1.Text
      If kodi1 = kodi2 Then
         objectInitialization KONFIGURIME_PERDORUES_MODIFIKO
         If data_modifikimi = "" Or data <> data_modifikimi Then
            objectInitialization MODIFIKIMI_I_DATES
         End If

         MsgBox "Perdoruesi u modifikua.", vbInformation, "Konfigurimi i perdoruesve."
         Me.txtUserName.Clear
         objectInitialization KONFIGURIME_PERDORUES_KERKO
         txtPerdorues.Text = ""
         txtPassword.Text = ""
         txtPassword1.Text = ""
         txtUserName.ListIndex = -1
         txtPerdorues.Visible = False
         txtUserName.Visible = True
         txtPassword1.Visible = False
         lbPassword1.Visible = False
         cmdVerifiko.Visible = True
         cmdModifikoUser.Visible = False
         cmdElimino.Visible = False
         cboStatusiModifiko.Visible = False
         Label5.Visible = False
         lblWorkstationModifiko.Visible = False
         Me.cboWorkstationModifiko.Visible = False
      Else
         MsgBox " Fjalekalimi i perseritur nuk eshte i njejte me ate qe dhate heren e pare." & Chr(10) & "Jepni perseri fjalekalimin.", vbInformation, "Konfigurimi i perdoruesve."
         txtPassword.Text = ""
         txtPassword1.Text = ""
      End If
   End If
   
   
   
End Sub




Private Sub cmdVerifiko_Click()
   If txtUserName = "" Or txtPassword.Text = "" Then
      MsgBox "Ju duhet te jepni te gjitha te dhenat e kerkuara per te modifikuar njeperdorues te ri.", vbInformation, "Konfigurimii perdoruesve."
   Else
      objectInitialization VERIFIKIMI_I_PERDORUESIT
      
   End If
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
  cmdElimino.Visible = False
  cmdModifikoUser.Visible = False
  txtPerdorues.Visible = False
  txtPassword1.Visible = False
  lbPassword1.Visible = False
  objectInitialization KONFIGURIME_PERDORUES_KERKO
  percaktoRendinSipasTabit
  mbushKomboStatusi
  MbushKomboWorkstation
  percaktoTeDrejtat
  
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

Private Sub percaktoRendinSipasTabit()
    Text2.TabIndex = 0
    Text1.TabIndex = 1
    Text3.TabIndex = 2
    cmdKrijoUser.TabIndex = 3
    
End Sub
    
Private Sub mbushKomboStatusi()
    cboStatusi.AddItem "Administrator"
    cboStatusi.AddItem "SupervizorEmesme"
    cboStatusi.AddItem "SupervizorTetevjecare"
    cboStatusi.AddItem "Vizitor"
    
    cboStatusiModifiko.AddItem "Administrator"
    cboStatusiModifiko.AddItem "SupervizorEmesme"
    cboStatusiModifiko.AddItem "SupervizorTetevjecare"
    cboStatusiModifiko.AddItem "Vizitor"
End Sub

Private Sub percaktoTeDrejtat()
    
    Select Case statusi
       
        Case "Administrator"
            
            
        
        Case "SupervizorEmesme"
        
            fraKrijo.Enabled = False
            Text1.Enabled = False
            Text2.Enabled = False
            Text3.Enabled = False
            cboStatusi.Enabled = False
            cmdKrijoUser.Enabled = False
            cboStatusiModifiko.Visible = False
            
            Label5.Visible = False
            
        Case "SupervizorTetevjecare"
        
            fraKrijo.Enabled = False
            Text1.Enabled = False
            Text2.Enabled = False
            Text3.Enabled = False
            cboStatusi.Enabled = False
            cmdKrijoUser.Enabled = False
            cboStatusiModifiko.Visible = False
            Label5.Visible = False
            
        Case "Vizitor"
        
            fraKrijo.Enabled = False
            Text1.Enabled = False
            Text2.Enabled = False
            Text3.Enabled = False
            cboStatusi.Enabled = False
            cmdKrijoUser.Enabled = False
            cboStatusiModifiko.Visible = False
            Label5.Visible = False
            
        Case Else
    End Select
           
End Sub

Private Sub MbushKomboWorkstation()
    GetServers (1)
End Sub

Private Function GetServers(dwServerType As Long) As Long

  'lists all servers running the specified
  'type of software that are visible in a domain
   Dim bufptr          As Long
   Dim dwEntriesread   As Long
   Dim dwTotalentries  As Long
   Dim dwResumehandle  As Long
   Dim se100           As SERVER_INFO_100
   Dim success         As Long
   Dim nStructSize     As Long
   Dim cnt             As Long

  'Call passing MAX_PREFERRED_LENGTH to have the
  'API allocate required memory for the return values.
  '
  'The call is enumerating all machines on the
  'network (SV_TYPE_ALL); however, by Or'ing
  'specific bit masks for defined types you can
  'customize the returned data. For example, a
  'value of 0x00000003 combines the bit masks for
  'SV_TYPE_WORKSTATION (0x00000001) and
  'SV_TYPE_SERVER (0x00000002).
  '
  'dwServerName must be Null. The level parameter
  '(101 here) specifies the data structure being
  'used (in this case a SERVER_INFO_101 structure).
  '
  'The domain member is passed as Null, indicating
  'machines on the primary domain are to be retrieved.
  'If you decide to use this member to enumerate
  'specific domains, pass StrPtr("YourDomainName"),
  'not a string directly.
   success = NetServerEnum(0&, _
                           100, _
                           bufptr, _
                           MAX_PREFERRED_LENGTH, _
                           dwEntriesread, _
                           dwTotalentries, _
                           dwServerType, _
                           0&, _
                           dwResumehandle)

  'if all goes well
   If success = NERR_SUCCESS Then
      
      nStructSize = LenB(se100)
      
     'loop through the returned data, adding
     'each machine to the list
      For cnt = 0 To dwEntriesread - 1
         
        'get one chunk of data and cast
        'into an SERVER_INFO_101 struct
        'in order to add the name to a list
         CopyMemory se100, ByVal bufptr + (nStructSize * cnt), nStructSize
            
         Me.cboWorkstation.AddItem GetPointerToByteStringW(se100.sv100_name)
         Me.cboWorkstationModifiko.AddItem GetPointerToByteStringW(se100.sv100_name)
      Next
      
   End If
   
  'clean up, regardless of success
  'and return number of entries read
  'as sign of success
   Call NetApiBufferFree(bufptr)
   GetServers = dwEntriesread

End Function
