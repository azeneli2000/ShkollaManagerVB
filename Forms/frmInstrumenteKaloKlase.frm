VERSION 5.00
Begin VB.Form frmInstrumenteKaloKlase 
   Appearance      =   0  'Flat
   Caption         =   "Kalo Klase"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15000
   WindowState     =   2  'Maximized
   Begin VB.ListBox lista1 
      Appearance      =   0  'Flat
      Height          =   3930
      Left            =   840
      MultiSelect     =   2  'Extended
      TabIndex        =   32
      Top             =   4080
      Width           =   4695
   End
   Begin VB.ListBox lista2Amza 
      Appearance      =   0  'Flat
      Height          =   1590
      Left            =   5760
      TabIndex        =   31
      Top             =   6960
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame fraCikli 
      Caption         =   "Cikli"
      ForeColor       =   &H00008000&
      Height          =   1455
      Left            =   840
      TabIndex        =   28
      Top             =   120
      Width           =   2415
      Begin VB.OptionButton optUlet 
         Caption         =   "Shkolla nëntëvjeçare"
         ForeColor       =   &H000000C0&
         Height          =   615
         Left            =   240
         TabIndex        =   30
         Top             =   720
         Width           =   1815
      End
      Begin VB.OptionButton optMesme 
         Caption         =   "Shkolla e mesme"
         ForeColor       =   &H00800000&
         Height          =   615
         Left            =   240
         TabIndex        =   29
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.ListBox listaAmza 
      Appearance      =   0  'Flat
      Height          =   2760
      Left            =   5880
      TabIndex        =   27
      Top             =   1800
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame fraInfo 
      Height          =   1695
      Left            =   12240
      TabIndex        =   22
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
         Picture         =   "frmInstrumenteKaloKlase.frx":0000
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
   Begin VB.CommandButton cmdKtheMbrapsht 
      BackColor       =   &H80000009&
      DisabledPicture =   "frmInstrumenteKaloKlase.frx":6F75
      DownPicture     =   "frmInstrumenteKaloKlase.frx":703F
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      Picture         =   "frmInstrumenteKaloKlase.frx":7109
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6240
      Width           =   1695
   End
   Begin VB.ListBox lista2 
      Appearance      =   0  'Flat
      Height          =   4320
      Left            =   7560
      MultiSelect     =   2  'Extended
      TabIndex        =   12
      Top             =   4080
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000009&
      DisabledPicture =   "frmInstrumenteKaloKlase.frx":71D3
      DownPicture     =   "frmInstrumenteKaloKlase.frx":A70D
      Height          =   375
      Left            =   11520
      Picture         =   "frmInstrumenteKaloKlase.frx":DC47
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8880
      Width           =   2295
   End
   Begin VB.ComboBox cboVitiShkollor2 
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
      Left            =   7800
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   3000
      Width           =   2055
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
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   3000
      Width           =   2055
   End
   Begin VB.CommandButton cmdKaloNeListe 
      BackColor       =   &H80000009&
      DisabledPicture =   "frmInstrumenteKaloKlase.frx":11181
      DownPicture     =   "frmInstrumenteKaloKlase.frx":16A0B
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      Picture         =   "frmInstrumenteKaloKlase.frx":1C295
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Frame fraKlasa 
      Caption         =   "Klasa e re :"
      ForeColor       =   &H00008000&
      Height          =   1935
      Left            =   7560
      TabIndex        =   2
      Top             =   1680
      Width           =   4335
      Begin VB.ComboBox cboIndeksi2 
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
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   600
         Width           =   855
      End
      Begin VB.ComboBox cboKlasa2 
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
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Viti shkollor :"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Indeksi :"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1440
         TabIndex        =   6
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblKlasa 
         Caption         =   "Klasa :"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H8000000E&
      DisabledPicture =   "frmInstrumenteKaloKlase.frx":21B1F
      DownPicture     =   "frmInstrumenteKaloKlase.frx":27F61
      Height          =   375
      Left            =   3600
      Picture         =   "frmInstrumenteKaloKlase.frx":2E3A3
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8880
      Width           =   2655
   End
   Begin VB.CommandButton cmdDil 
      BackColor       =   &H8000000E&
      DisabledPicture =   "frmInstrumenteKaloKlase.frx":347E5
      DownPicture     =   "frmInstrumenteKaloKlase.frx":3AC27
      Height          =   375
      Left            =   7080
      Picture         =   "frmInstrumenteKaloKlase.frx":41069
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8880
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Caption         =   "Klasa e vjeter :"
      ForeColor       =   &H00008000&
      Height          =   1935
      Left            =   840
      TabIndex        =   15
      Top             =   1680
      Width           =   4695
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
         ItemData        =   "frmInstrumenteKaloKlase.frx":474AB
         Left            =   120
         List            =   "frmInstrumenteKaloKlase.frx":474AD
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   600
         Width           =   855
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
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Viti shkollor :"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Klasa :"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblIndeksi 
         Caption         =   "Indeksi :"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1320
         TabIndex        =   18
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nxenesit  e  klases se re :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   7560
      TabIndex        =   13
      Top             =   3720
      Width           =   4335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nxenesit e klases se vjeter :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   840
      TabIndex        =   11
      Top             =   3720
      Width           =   4695
   End
End
Attribute VB_Name = "frmInstrumenteKaloKlase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
 Dim nr                As Integer
 Dim objektGabimi As New clsErrorHandler






Private Sub cboIndeksi_Click()
   Dim Klasa            As String
   Dim indeksi          As String
   Dim vitishkollor     As String
   Klasa = cboKlasa.Text
   indeksi = cboIndeksi.Text
   vitishkollor = cboVitiShkollor.Text
   If cboKlasa.Text <> "" And cboIndeksi.Text <> "" And cboVitiShkollor.Text <> "" And cboKlasa2.Text <> "" And cboIndeksi2.Text <> "" And cboVitiShkollor2.Text <> "" Then
      If cboVitiShkollor.Text <> cboVitiShkollor2.Text Then
        cmdKtheMbrapsht.Enabled = False
      Else
        cmdKtheMbrapsht.Enabled = True
      End If
   Else
      cmdKtheMbrapsht.Enabled = True
   End If
   If Klasa <> "" And indeksi <> "" And vitishkollor <> "" Then
      lista1.Clear
      objektGabimi.kapGabimin
      If objektGabimi.mvarGabimi = 0 Then
        If gabimCikli(1) = False Then
         objectInitialization INSTRUMENTE_KALO_KLASE_KERKO
        End If
      Else
         objektGabimi.menazhim_gabimi
      End If
   End If
End Sub


Private Sub cboIndeksi2_Click()
   Dim Klasa            As String
   Dim indeksi          As String
   Dim vitishkollor     As String
   Klasa = cboKlasa2.Text
   indeksi = cboIndeksi2.Text
   vitishkollor = cboVitiShkollor2.Text
   If cboKlasa.Text <> "" And cboIndeksi.Text <> "" And cboVitiShkollor.Text <> "" And cboKlasa2.Text <> "" And cboIndeksi2.Text <> "" And cboVitiShkollor2.Text <> "" Then
      If cboVitiShkollor.Text <> cboVitiShkollor2.Text Then
        cmdKtheMbrapsht.Enabled = False
      Else
        cmdKtheMbrapsht.Enabled = True
      End If
   Else
      cmdKtheMbrapsht.Enabled = True
   End If
   If Klasa <> "" And indeksi <> "" And vitishkollor <> "" Then
      lista2.Clear
      objektGabimi.kapGabimin
      If objektGabimi.mvarGabimi = 0 Then
        If gabimCikli(2) = False Then
         objectInitialization INSTRUMENTE_KERKO
        End If
      Else
         objektGabimi.menazhim_gabimi
      End If
   End If
End Sub

Private Sub cboKlasa_Click()
   Dim Klasa            As String
   Dim indeksi          As String
   Dim vitishkollor     As String
   Klasa = cboKlasa.Text
   indeksi = cboIndeksi.Text
   vitishkollor = cboVitiShkollor.Text
   If cboKlasa.Text <> "" And cboIndeksi.Text <> "" And cboVitiShkollor.Text <> "" And cboKlasa2.Text <> "" And cboIndeksi2.Text <> "" And cboVitiShkollor2.Text <> "" Then
      If cboVitiShkollor.Text <> cboVitiShkollor2.Text Then
        cmdKtheMbrapsht.Enabled = False
      Else
        cmdKtheMbrapsht.Enabled = True
      End If
   Else
      cmdKtheMbrapsht.Enabled = True
   End If
   If Klasa <> "" And indeksi <> "" And vitishkollor <> "" Then
   
      lista1.Clear
      objektGabimi.kapGabimin
      If objektGabimi.mvarGabimi = 0 Then
        If gabimCikli(1) = False Then
         objectInitialization INSTRUMENTE_KALO_KLASE_KERKO
        End If
      Else
         objektGabimi.menazhim_gabimi
      End If
   End If
End Sub


Private Sub cboKlasa2_Click()
   Dim Klasa            As String
   Dim indeksi          As String
   Dim vitishkollor     As String
   Klasa = cboKlasa2.Text
   indeksi = cboIndeksi2.Text
   vitishkollor = cboVitiShkollor2.Text
   If cboKlasa.Text <> "" And cboIndeksi.Text <> "" And cboVitiShkollor.Text <> "" And cboKlasa2.Text <> "" And cboIndeksi2.Text <> "" And cboVitiShkollor2.Text <> "" Then
      If cboVitiShkollor.Text <> cboVitiShkollor2.Text Then
        cmdKtheMbrapsht.Enabled = False
      Else
        cmdKtheMbrapsht.Enabled = True
      End If
   Else
      cmdKtheMbrapsht.Enabled = True
   End If
   If Klasa <> "" And indeksi <> "" And vitishkollor <> "" Then
      lista2.Clear
      objektGabimi.kapGabimin
      If objektGabimi.mvarGabimi = 0 Then
        If gabimCikli(2) = False Then
         objectInitialization INSTRUMENTE_KERKO
        End If
      Else
         objektGabimi.menazhim_gabimi
      End If
    End If
   End Sub

Private Sub cboVitiShkollor_Click()
   Dim Klasa            As String
   Dim indeksi          As String
   Dim vitishkollor     As String
   Klasa = cboKlasa.Text
   indeksi = cboIndeksi.Text
   vitishkollor = cboVitiShkollor.Text
   If cboKlasa.Text <> "" And cboIndeksi.Text <> "" And cboVitiShkollor.Text <> "" And cboKlasa2.Text <> "" And cboIndeksi2.Text <> "" And cboVitiShkollor2.Text <> "" Then
      If cboVitiShkollor.Text <> cboVitiShkollor2.Text Then
        cmdKtheMbrapsht.Enabled = False
      Else
        cmdKtheMbrapsht.Enabled = True
      End If
   Else
      cmdKtheMbrapsht.Enabled = True
   End If
   If Klasa <> "" And indeksi <> "" And vitishkollor <> "" Then
      
      lista1.Clear
      objektGabimi.kapGabimin
      If objektGabimi.mvarGabimi = 0 Then
        If gabimCikli(1) = False Then
         objectInitialization INSTRUMENTE_KALO_KLASE_KERKO
        End If
      Else
         objektGabimi.menazhim_gabimi
      End If
   End If
End Sub


Private Sub cboVitiShkollor2_Click()
   Dim Klasa            As String
   Dim indeksi          As String
   Dim vitishkollor     As String
   Klasa = cboKlasa2.Text
   indeksi = cboIndeksi2.Text
   vitishkollor = cboVitiShkollor2.Text
   If cboKlasa.Text <> "" And cboIndeksi.Text <> "" And cboVitiShkollor.Text <> "" And cboKlasa2.Text <> "" And cboIndeksi2.Text <> "" And cboVitiShkollor2.Text <> "" Then
      If cboVitiShkollor.Text <> cboVitiShkollor2.Text Then
        cmdKtheMbrapsht.Enabled = False
      Else
        cmdKtheMbrapsht.Enabled = True
      End If
   Else
      cmdKtheMbrapsht.Enabled = True
   End If
   If Klasa <> "" And indeksi <> "" And vitishkollor <> "" Then
      
      lista2.Clear
      objektGabimi.kapGabimin
      If objektGabimi.mvarGabimi = 0 Then
        If gabimCikli(2) = False Then
         objectInitialization INSTRUMENTE_KERKO
        End If
      Else
         objektGabimi.menazhim_gabimi
      End If
   End If
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

Private Sub cmdKaloNeListe_Click()
   Dim klasa1, klasa2   As Integer

   If cboKlasa.Text = "" Or cboIndeksi.Text = "" Or cboVitiShkollor.Text = "" Or cboKlasa2.Text = "" Or cboIndeksi2.Text = "" Or cboVitiShkollor2.Text = "" Then
      MsgBox "Ju duhet te percaktoni te dhenat identifikuese te klases se re dhe te vjeter" & Chr(10) & "per te kaluar nxenesit nga nje klase ne nje tjeter.", vbInformation, "Konfigurimi i klasave."
      Exit Sub
   End If

   klasa1 = Val(cboKlasa.Text)
   klasa2 = Val(cboKlasa2.Text)
   If klasa1 > klasa2 Then
      MsgBox "Ju nuk mund te kaloni nxenes ne nje klase me te ulet." & Chr(10) & "Jepni perseri klasat.", vbInformation, "Konfigurimi i klasave."
      cboKlasa2.ListIndex = -1
      cboIndeksi2.ListIndex = -1
      cboVitiShkollor2.ListIndex = -1
      lista2.Clear
      lista2Amza.Clear
      Exit Sub
   End If
   If klasa2 > klasa1 + 1 Then
      MsgBox "Ju nuk mund te kaloni nje nxenes ne nje klase me te larte se klasa pasardhese." & Chr(10) & "Jepni perseri klasen.", vbInformation, "Konfigurimi i klasave."
      cboKlasa2.ListIndex = -1
      cboIndeksi2.ListIndex = -1
      cboVitiShkollor2.ListIndex = -1
      lista2.Clear
      lista2Amza.Clear
      Exit Sub
   End If
   If cboVitiShkollor.Text > cboVitiShkollor2.Text Then

      MsgBox "Ju nuk mund te kaloni nxenes ne nje vit shkollor  paraardhes." & Chr(10) & "Jepni perseri vitin shkollor.", vbInformation, "Konfigurimi i klasave."
      cboKlasa2.ListIndex = -1
      cboIndeksi2.ListIndex = -1
      cboVitiShkollor2.ListIndex = -1
      lista2.Clear
      lista2Amza.Clear
      Exit Sub
   End If
   If (cboVitiShkollor.Text = cboVitiShkollor2.Text) And (cboKlasa.Text <> cboKlasa2.Text) Then
      MsgBox "Ju nuk mund te hidhni levizni nxenesit ne klasa te ndryshme ( si numer ) per vitin shkollor te njejte.", vbExclamation, "Konfigurimi i klasave."
      cboKlasa2.ListIndex = -1
      cboIndeksi2.ListIndex = -1
      cboVitiShkollor2.ListIndex = -1
      lista2.Clear
      lista2Amza.Clear
      Exit Sub
   End If
   If (cboVitiShkollor2.ListIndex > cboVitiShkollor.ListIndex + 1) Then
      MsgBox "Ju nuk mund te hidhni nxenesit ne nje vit shkollor " & Chr(10) & "me te madh se pasardhesi.", vbInformation, "Konfigurimi i klasave."
      cboKlasa2.ListIndex = -1
      cboIndeksi2.ListIndex = -1
      cboVitiShkollor2.ListIndex = -1
      lista2.Clear
      lista2Amza.Clear
      Exit Sub
   End If
   objectInitialization INSTRUMENTE_NXENES_KALO_NE_LISTE


End Sub


Private Sub cmdKtheMbrapsht_Click()

   Dim klasa1           As String
   Dim klasa2           As String
   Dim vitiShkollor1    As String
   Dim vitiShkollor2    As String
   Dim indeksi1         As String
   Dim indeksi2         As String
   Dim numerAmze        As String
   Dim strCikli         As String
   Dim emri             As String
   Dim ugjet            As Boolean
   If optMesme.Value Then
      strCikli = "TRUE"
   Else
      strCikli = "FALSE"
   End If
   klasa1 = cboKlasa.Text
   klasa2 = cboKlasa2.Text
   vitiShkollor1 = cboVitiShkollor.Text
   vitiShkollor2 = cboVitiShkollor2.Text
   indeksi1 = cboIndeksi.Text
   indeksi2 = cboIndeksi2.Text
   Dim I                As Integer
   If klasa1 = "" Or vitiShkollor1 = "" Or indeksi1 = "" Then
      objectInitialization KONTROLLI_I_REGJISTRIMIT
   Else

      I = 0
      Do While (I <= lista2.ListCount - 1)
         If lista2.Selected(I) Then
            emri = lista2.List(I)
            numerAmze = lista2Amza.List(I)
            If Not gjendetNeListe(emri, lista1) Then
               lista2.RemoveItem (I)
               lista2Amza.RemoveItem (I)
               lista1.AddItem emri
               listaAmza.AddItem numerAmze
               I = I - 1
            End If
         End If
         I = I + 1
      Loop
   End If
End Sub

Private Sub cmdOK_Click()

   Dim klasa1           As String
   Dim klasa2           As String
   Dim vitiShkollor1    As String
   Dim vitiShkollor2    As String
   Dim indeksi1         As String
   Dim indeksi2         As String
   Dim data             As String
   data = date
   klasa1 = cboKlasa.Text
   klasa2 = cboKlasa2.Text
   vitiShkollor1 = cboVitiShkollor.Text
   vitiShkollor2 = cboVitiShkollor2.Text
   indeksi1 = cboIndeksi.Text
   indeksi2 = cboIndeksi2.Text
   If klasa2 = "" Or indeksi2 = "" Or vitiShkollor2 = "" Or klasa1 = "" Or indeksi1 = "" Or vitiShkollor1 = "" Then
      Exit Sub
   Else

      objectInitialization INSTRUMENTE_KALO_KLASE
      If data_modifikimi = "" Or data <> data_modifikimi Then
         objectInitialization MODIFIKIMI_I_DATES
      End If
   End If
   listaAmza.Clear
   lista2Amza.Clear
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
   mbushKomboBox
   optMesme.Value = True
   Image1.Picture = LoadPicture(adresaLogo)
   lblShkolla.Caption = emerShkolla
   percaktoTeDrejtat
End Sub

Private Sub Form_Resize()
  fraInfo.Left = Me.Left + Me.Width - fraInfo.Width - 200
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
    
    cboVitiShkollor2.AddItem "1990-1991"
    cboVitiShkollor2.AddItem "1991-1992"
    cboVitiShkollor2.AddItem "1992-1993"
    cboVitiShkollor2.AddItem "1993-1994"
    cboVitiShkollor2.AddItem "1994-1995"
    cboVitiShkollor2.AddItem "1995-1996"
    cboVitiShkollor2.AddItem "1996-1997"
    cboVitiShkollor2.AddItem "1997-1998"
    cboVitiShkollor2.AddItem "1998-1999"
    cboVitiShkollor2.AddItem "1999-2000"
    cboVitiShkollor2.AddItem "2000-2001"
    cboVitiShkollor2.AddItem "2001-2002"
    cboVitiShkollor2.AddItem "2002-2003"
    cboVitiShkollor2.AddItem "2003-2004"
    cboVitiShkollor2.AddItem "2004-2005"
    cboVitiShkollor2.AddItem "2005-2006"
    cboVitiShkollor2.AddItem "2006-2007"
    cboVitiShkollor2.AddItem "2007-2008"
    cboVitiShkollor2.AddItem "2008-2009"
    cboVitiShkollor2.AddItem "2009-2010"
    cboVitiShkollor2.AddItem "2010-2011"
    cboVitiShkollor2.AddItem "2011-2012"
    cboVitiShkollor2.AddItem "2012-2012"
    cboVitiShkollor2.AddItem "2013-2014"
    cboVitiShkollor2.AddItem "2014-2015"
    cboVitiShkollor2.AddItem "2015-2016"
    cboVitiShkollor2.AddItem "2016-2017"
    cboVitiShkollor2.AddItem "2017-2018"
    cboVitiShkollor2.AddItem "2018-2019"
    cboVitiShkollor2.AddItem "2019-2020"
    
    cboKlasa.AddItem "9"
    cboKlasa.AddItem "10"
    cboKlasa.AddItem "11"
    cboKlasa.AddItem "12"
    
    
    cboKlasa2.AddItem "9"
    cboKlasa2.AddItem "10"
    cboKlasa2.AddItem "11"
    cboKlasa2.AddItem "12"
    
    cboIndeksi.AddItem "A"
    cboIndeksi.AddItem "B"
    cboIndeksi.AddItem "C"
    cboIndeksi.AddItem "D"
    cboIndeksi.AddItem "E"
    cboIndeksi.AddItem "F"
    
    cboIndeksi2.AddItem "A"
    cboIndeksi2.AddItem "B"
    cboIndeksi2.AddItem "C"
    cboIndeksi2.AddItem "D"
    cboIndeksi2.AddItem "E"
    cboIndeksi2.AddItem "F"
    
    
    
End Sub


Private Function gjendetNeListe(emri As String, liste As ListBox) As Boolean
    Dim j As Integer
    Dim ugjet As Boolean
    ugjet = False
    j = 0
    Do While (j <= liste.ListCount - 1) And (ugjet = False)
        If LCase(emri) = LCase(liste.List(j)) Then
            ugjet = True
        Else
            j = j + 1
        End If
    Loop
    
    gjendetNeListe = ugjet
     
End Function

Private Sub objectInitialization(actionName As GUI_ACTION__ENUM)
   Set objUIController = New clsUIController

   objUIController.actionName = actionName
   objUIController.ExecuteActions

   Set objUIController = Nothing
End Sub

Private Sub mbushKomboUlet()
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
    
    cboKlasa2.Clear
    cboKlasa2.AddItem "1"
    cboKlasa2.AddItem "2"
    cboKlasa2.AddItem "3"
    cboKlasa2.AddItem "4"
    cboKlasa2.AddItem "5"
    cboKlasa2.AddItem "6"
    cboKlasa2.AddItem "7"
    cboKlasa2.AddItem "8"
    cboKlasa2.AddItem "9"
End Sub

Private Sub mbushKomboMesme()
    cboKlasa.Clear
    cboKlasa.AddItem "9"
    cboKlasa.AddItem "10"
    cboKlasa.AddItem "11"
    cboKlasa.AddItem "12"
    
    cboKlasa2.Clear
    cboKlasa2.AddItem "9"
    cboKlasa2.AddItem "10"
    cboKlasa2.AddItem "11"
    cboKlasa2.AddItem "12"
End Sub



Private Sub optMesme_Click()
    mbushKomboMesme
    Pastro
End Sub

Private Sub optUlet_Click()
    mbushKomboUlet
    Pastro
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
            mbushKomboUlet
        Case Else
            
    End Select
    
            
            
End Sub

Private Sub Pastro()

    cboKlasa.ListIndex = -1
    cboIndeksi.ListIndex = -1
    cboVitiShkollor.ListIndex = -1
    
    cboKlasa2.ListIndex = -1
    cboIndeksi2.ListIndex = -1
    cboVitiShkollor2.ListIndex = -1
    
    lista1.Clear
    listaAmza.Clear
    lista2Amza.Clear
    lista2.Clear
End Sub

Private Function gabimCikli(I As Integer) As Boolean
    Dim cikliZgjedhur As Boolean
    Dim strCikli  As String
    strCikli = " shkollës nëntëvjeçare!"
    cikliZgjedhur = False
    If Me.optMesme.Value Then
        cikliZgjedhur = True
        strCikli = " shkollës së mesme!"
    End If
    If (I = 1) Then
        If (cikliZgjedhur <> ktheCiklin(Me.cboKlasa.Text, CStr(Me.cboVitiShkollor.Text))) Then
            MsgBox "Për vitin shkollor " + Me.cboVitiShkollor.Text + " klasa e " + cboKlasa.Text + "-të nuk i përket " + strCikli, vbExclamation, "Kalo klase"
            cboKlasa.ListIndex = -1
            cboIndeksi.ListIndex = -1
            cboVitiShkollor.ListIndex = -1
            gabimCikli = True
        Else
            gabimCikli = False
        End If
    Else
        If (cikliZgjedhur <> ktheCiklin(Me.cboKlasa2.Text, CStr(Me.cboVitiShkollor2.Text))) Then
            MsgBox "Për vitin shkollor " + Me.cboVitiShkollor2.Text + " klasa e " + cboKlasa2.Text + "-të nuk i përket " + strCikli, vbExclamation, "Kalo klase"
            cboKlasa2.ListIndex = -1
            cboIndeksi2.ListIndex = -1
            cboVitiShkollor2.ListIndex = -1
            gabimCikli = True
        Else
            gabimCikli = False
        End If
    End If
End Function




