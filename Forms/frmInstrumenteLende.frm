VERSION 5.00
Begin VB.Form frmInstrumenteLende 
   Caption         =   "Konfigurimi i lendeve"
   ClientHeight    =   8565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14685
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
   ForeColor       =   &H00C00000&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8565
   ScaleWidth      =   14685
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.ListBox listaNumri 
      Height          =   1425
      Left            =   4800
      TabIndex        =   32
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox listaNrM 
      Height          =   1425
      Left            =   3720
      TabIndex        =   31
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox listaNrU 
      Height          =   1425
      Left            =   2520
      TabIndex        =   30
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
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
      Height          =   1695
      Left            =   11880
      TabIndex        =   25
      Top             =   0
      Width           =   2655
      Begin VB.CommandButton cmdWebsiste 
         BackColor       =   &H80000009&
         Caption         =   "Website"
         Height          =   375
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton cmdEmail 
         BackColor       =   &H80000009&
         Caption         =   "E-mail"
         Height          =   375
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   960
         Width           =   1095
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   120
         Picture         =   "frmInstrumenteLende.frx":0000
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
         TabIndex        =   29
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
         TabIndex        =   28
         Top             =   120
         Width           =   2415
      End
   End
   Begin VB.CommandButton cmdPrano 
      BackColor       =   &H80000009&
      Height          =   375
      Left            =   11400
      Picture         =   "frmInstrumenteLende.frx":6F75
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6360
      Width           =   2295
   End
   Begin VB.CommandButton cmdHiq 
      BackColor       =   &H80000009&
      DisabledPicture =   "frmInstrumenteLende.frx":D3B7
      DownPicture     =   "frmInstrumenteLende.frx":137F9
      Height          =   375
      Left            =   11400
      Picture         =   "frmInstrumenteLende.frx":19C3B
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5760
      Width           =   2295
   End
   Begin VB.CommandButton cmdShto 
      BackColor       =   &H80000009&
      DisabledPicture =   "frmInstrumenteLende.frx":2007D
      DownPicture     =   "frmInstrumenteLende.frx":264BF
      Height          =   375
      Left            =   11400
      Picture         =   "frmInstrumenteLende.frx":2C901
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5160
      Width           =   2295
   End
   Begin VB.ListBox listaProvimet 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1785
      Left            =   11280
      TabIndex        =   14
      Top             =   2640
      Width           =   2775
   End
   Begin VB.OptionButton optTetevjecare 
      Caption         =   "Shkolla nëntëvjecare"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   615
      Left            =   360
      TabIndex        =   13
      Top             =   1080
      Width           =   1455
   End
   Begin VB.OptionButton optMesme 
      Caption         =   "Shkolla e mesme"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   360
      TabIndex        =   12
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton cmdKtheMbrapsht 
      BackColor       =   &H80000009&
      DisabledPicture =   "frmInstrumenteLende.frx":32D43
      DownPicture     =   "frmInstrumenteLende.frx":32E0D
      Height          =   495
      Left            =   4200
      Picture         =   "frmInstrumenteLende.frx":32ED7
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3735
      Width           =   2175
   End
   Begin VB.ListBox lista2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3540
      Left            =   6720
      TabIndex        =   10
      Top             =   2640
      Width           =   3375
   End
   Begin VB.ListBox lista1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3540
      Left            =   360
      MultiSelect     =   2  'Extended
      TabIndex        =   9
      Top             =   2640
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000009&
      DisabledPicture =   "frmInstrumenteLende.frx":32FA1
      DownPicture     =   "frmInstrumenteLende.frx":364DB
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
      Left            =   11400
      Picture         =   "frmInstrumenteLende.frx":39A15
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9120
      Width           =   2535
   End
   Begin VB.ComboBox cboVitiShkollor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      ForeColor       =   &H80000004&
      Height          =   360
      Left            =   6360
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1320
      Width           =   2055
   End
   Begin VB.CommandButton cmdShtoLende 
      BackColor       =   &H80000009&
      DisabledPicture =   "frmInstrumenteLende.frx":3CF4F
      DownPicture     =   "frmInstrumenteLende.frx":45461
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
      Left            =   4200
      Picture         =   "frmInstrumenteLende.frx":4D973
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4455
      Width           =   2175
   End
   Begin VB.CommandButton cmdDil 
      BackColor       =   &H8000000E&
      DisabledPicture =   "frmInstrumenteLende.frx":55E85
      DownPicture     =   "frmInstrumenteLende.frx":5C2C7
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
      Left            =   6720
      Picture         =   "frmInstrumenteLende.frx":62709
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9120
      Width           =   2535
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H8000000E&
      DisabledPicture =   "frmInstrumenteLende.frx":68B4B
      DownPicture     =   "frmInstrumenteLende.frx":6EF8D
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
      Left            =   3960
      Picture         =   "frmInstrumenteLende.frx":753CF
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6960
      Width           =   2655
   End
   Begin VB.Frame fraKlasa 
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
      Left            =   6240
      TabIndex        =   1
      Top             =   120
      Width           =   4215
      Begin VB.ComboBox cboKlasa 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         ForeColor       =   &H80000004&
         Height          =   360
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Viti shkollor :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label lblKlasa 
         Caption         =   "Klasa :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdKaloNeListe 
      BackColor       =   &H80000009&
      DisabledPicture =   "frmInstrumenteLende.frx":7B811
      DownPicture     =   "frmInstrumenteLende.frx":8109B
      Height          =   495
      Left            =   4200
      Picture         =   "frmInstrumenteLende.frx":86925
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Frame fraCikli 
      Caption         =   "Cikli"
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
      Height          =   1695
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   1815
   End
   Begin VB.Frame fraLendet 
      Caption         =   "Konfigurimi i lendeve :"
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
      Height          =   5655
      Left            =   120
      TabIndex        =   20
      Top             =   1920
      Width           =   10335
      Begin VB.TextBox txtLendaRe 
         Height          =   375
         Left            =   6840
         TabIndex        =   34
         Text            =   "Text1"
         Top             =   4800
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton cmdModifikoLende 
         BackColor       =   &H80000009&
         DisabledPicture =   "frmInstrumenteLende.frx":8C1AF
         DownPicture     =   "frmInstrumenteLende.frx":93721
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
         Left            =   4080
         Picture         =   "frmInstrumenteLende.frx":9AC93
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   3240
         Width           =   2175
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Lendet e klases se zgjedhur :"
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   6600
         TabIndex        =   22
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Lendet sipas ciklit :"
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   3495
      End
   End
   Begin VB.Frame fraProvimet 
      Caption         =   "Konfigurimi i provimeve :"
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
      Height          =   4935
      Left            =   10680
      TabIndex        =   23
      Top             =   1920
      Width           =   3975
      Begin VB.Label lblProvimet 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Provimet sipas ciklit :"
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   600
         TabIndex        =   24
         Top             =   360
         Width           =   2775
      End
   End
End
Attribute VB_Name = "frmInstrumenteLende"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim first As Boolean

Private Sub cboKlasa_Click()
  listaNumri.Clear
  If first = False Then
    If cboVitiShkollor.Text <> "" And (optMesme.Value Or optTetevjecare.Value) Then
       objectInitialization INSTRUMENTE_LENDE_KLASA_KLIK
    Else
       MsgBox "Per te pare lendet qe zhvillon kjo klase ju duhet te zgjidhni vitin shkollor perkates dhe ciklin.", , "Konfigurimi i lendeve."
    End If
    Else
    first = False
  End If
    
End Sub





Private Sub cboVitiShkollor_Click()
    
   listaProvimet.Clear
   listaNumri.Clear
   'objectInitialization INSTRUMENTE_VITI_SHKOLLOR_KLIK
   If cboKlasa.Text <> "" And (optMesme.Value Or optTetevjecare.Value) Then
        objectInitialization INSTRUMENTE_LENDE_KLASA_KLIK
   End If
   If (optMesme.Value Or optTetevjecare.Value) Then
        objectInitialization SHFAQ_PROVIME
   End If
   Dim Klasa As String
   Klasa = Me.cboKlasa.Text
   If (Me.optMesme.Value) Then
        mbushKomboMesme
   Else
        mbushKomboTetevjecare
   End If
   Dim vitFillimi As Integer
   vitFillimi = CInt(Mid(Me.cboVitiShkollor.Text, 1, 4))
   If (Klasa <> "") Then
        If (Klasa = "9") Then
            If (optMesme.Value And vitFillimi >= 2008) Then
                Exit Sub
            End If
            If (Not optMesme.Value And vitFillimi <= 2007) Then
                Exit Sub
            End If
        End If
        Me.cboKlasa.Text = Klasa
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

Private Sub cmdHiq_Click()
   If optMesme.Value = False And optTetevjecare.Value = False Then
      MsgBox "Ju duhet te percaktoni me pare ciklin e shkolles.", vbExclamation, "Konfigurimi i lendeve."
      Exit Sub
   End If
   If cboVitiShkollor.Text = "" Then
      MsgBox "Ju duhet te zgjidhni vitin shkollor.", vbExclamation, "Konfigurimi i lendeve."
      Exit Sub
   End If
   
   
   objectInitialization FSHIRJA_E_PROVIMEVE
   
End Sub

Private Sub cmdKaloNeListe_Click()

   If optMesme.Value = False And optTetevjecare.Value = False Then
      MsgBox "Ju duhet te percaktoni me pare ciklin e shkolles.", vbExclamation, "Konfigurimi i lendeve."
      Exit Sub
   End If
   If cboKlasa.Text = "" Then
      MsgBox "Ju duhet te zgjidhni klasen lendet e se ciles do te hidhni.", vbExclamation, "Konfigurimi i lendeve."
      Exit Sub
   End If
   If cboVitiShkollor.Text = "" Then
      MsgBox "Ju duhet te zgjidhni vitin shkollor.", vbExclamation, "Konfigurimi i lendeve."
      Exit Sub
   End If

   Dim cikli            As Boolean
   Dim I                As Integer
   Dim emri             As String
   For I = 0 To lista1.ListCount - 1
      If (lista1.Selected(I) = True) Then
         emri = lista1.List(I)

         If Not gjendetNeListe(emri, lista2) Then
            If lista2.ListCount >= 20 Then
               MsgBox "Ju nuk mun t'i shenjoni nje klase me shume se 20 lende.", vbInformation, "Konfigurimi i lendeve."
               Exit Sub
            End If
            lista2.AddItem emri

            If optMesme.Value Then
               listaNumri.AddItem listaNrM.List(I)
            End If
            If optTetevjecare.Value Then
               listaNumri.AddItem listaNrU.List(I)
            End If
         End If
      End If
   Next I
End Sub

Private Sub cmdKtheMbrapsht_Click()

   If optMesme.Value = False And optTetevjecare.Value = False Then
      MsgBox "Ju duhet te percaktoni me pare ciklin e shkolles.", vbExclamation, "Konfigurimi i lendeve."
      Exit Sub
   End If
   If cboKlasa.Text = "" Then
      MsgBox "Ju duhet te zgjidhni klasen lendet e se ciles do te hidhni.", vbExclamation, "Konfigurimi i lendeve."
      Exit Sub
   End If
   If cboVitiShkollor.Text = "" Then
      MsgBox "Ju duhet te zgjidhni vitin shkollor.", vbExclamation, "Konfigurimi i lendeve."
      Exit Sub
   End If
   objectInitialization HEQJA_E_LENDEVE
End Sub



Private Sub cmdOK_Click()

   Dim data             As String
   data = date
   If optMesme.Value = False And optTetevjecare.Value = False Then
      MsgBox "Ju duhet te percaktoni me pare ciklin e shkolles.", vbExclamation, "Konfigurimi i lendeve."
      Exit Sub
   End If
   If cboKlasa.Text = "" Then
      MsgBox "Ju duhet te zgjidhni klasen lendet e se ciles do te hidhni.", vbExclamation, "Konfigurimi i lendeve."
      Exit Sub
   End If
   If cboVitiShkollor.Text = "" Then
      MsgBox "Ju duhet te zgjidhni vitin shkollor.", vbExclamation, "Konfigurimi i lendeve."
      Exit Sub
   End If
   objectInitialization INSTRUMENTE_LENDE
   If data_modifikimi = "" Or data <> data_modifikimi Then
      objectInitialization MODIFIKIMI_I_DATES
   End If
   MsgBox "Lendet per klasen e zgjedhur nga ju u modifikuan.", , "Konfigurimi i lendeve."

End Sub



Private Sub cmdPrano_Click()
    If optMesme.Value = False And optTetevjecare.Value = False Then
        MsgBox "Ju duhet te percaktoni me pare ciklin e shkolles.", vbExclamation, "Konfigurimi i lendeve."
        Exit Sub
    End If
    If cboVitiShkollor.Text = "" Then
        MsgBox "Ju duhet te zgjidhni vitin shkollor.", vbExclamation, "Konfigurimi i lendeve."
        Exit Sub
    End If
    objectInitialization INSTRUMENTE_LENDE_HIDH_PROVIME
    MsgBox "Provimet u hodhen.", , "Konfigurimi  lendeve."
    'listaProvimet.Clear
  
End Sub

Private Sub cmdShto_Click()
   If optMesme.Value = False And optTetevjecare.Value = False Then
      MsgBox "Ju duhet te percaktoni me pare ciklin e shkolles.", vbExclamation, "Konfigurimi i lendeve."
      Exit Sub
   End If
   If cboVitiShkollor.Text = "" Then
      MsgBox "Ju duhet te zgjidhni vitin shkollor.", vbExclamation, "Konfigurimi i lendeve."
      Exit Sub
   End If
   Dim provimi          As String
   provimi = InputBox("Jepni emrin e provimit :", "Konfigurimi i lendeve.")
   If provimi <> "" Then
      If Not gjendetNeListe(provimi, listaProvimet) Then
         listaProvimet.AddItem provimi
      Else
         MsgBox "Ky provim tashme ekziston.", , "Konfigurimi i lendeve."
      End If
   End If
End Sub

Private Sub cmdShtoLende_Click()

   If optMesme.Value = False And optTetevjecare.Value = False Then
      MsgBox "Ju duhet te percaktoni me pare ciklin e shkolles.", vbExclamation, "Konfigurimi i lendeve."
      Exit Sub
   End If
   If cboKlasa.Text = "" Then
      MsgBox "Ju duhet te zgjidhni klasen lendet e se ciles do te hidhni.", vbExclamation, "Konfigurimi i lendeve."
      Exit Sub
   End If
   If cboVitiShkollor.Text = "" Then
      MsgBox "Ju duhet te zgjidhni vitin shkollor.", vbExclamation, "Konfigurimi i lendeve."
      Exit Sub
   End If
   If lista2.ListCount >= 27 Then
      MsgBox "Ju nuk mund t'i shtoni nje klase me shume se 27 lende.", vbInformation, "Konfigurimi i lendeve."
      Exit Sub
   End If
   Dim lenda            As String
   lenda = InputBox("Jepni emrin e lendes qe doni te shtoni", "Lenda e re")
   If lenda <> "" Then
      If Not gjendetNeListe(lenda, lista2) Then
         lista2.AddItem lenda
         listaNumri.AddItem "100"
      End If

   End If

End Sub

Private Sub cmdModifikoLende_Click()
    If optMesme.Value = False And optTetevjecare.Value = False Then
      MsgBox "Ju duhet te percaktoni me pare ciklin e shkolles.", vbExclamation, "Konfigurimi i lendeve."
      Exit Sub
   End If
   If cboKlasa.Text = "" Then
      MsgBox "Ju duhet te zgjidhni klasen lendet e se ciles do te hidhni.", vbExclamation, "Konfigurimi i lendeve."
      Exit Sub
   End If
   If cboVitiShkollor.Text = "" Then
      MsgBox "Ju duhet te zgjidhni vitin shkollor.", vbExclamation, "Konfigurimi i lendeve."
      Exit Sub
   End If
   If lista2.ListCount >= 27 Then
      MsgBox "Ju nuk mund t'i shtoni nje klase me shume se 27 lende.", vbInformation, "Konfigurimi i lendeve."
      Exit Sub
   End If
   If (lista2.ListIndex <= 0) Then
      MsgBox "Zgjidhni njërën prej lëndëve para se të modifikoni.", vbInformation, "Konfigurimi i lendeve."
      Exit Sub
   End If
   Dim lendaVjeter, lendaRe As String
   lendaVjeter = CStr(lista2.List(lista2.ListIndex))
   lendaRe = InputBox("Jepni emrin e ri të lëndës që doni të modifikoni!", "Modifiko lëndë", lendaVjeter)
   If (Trim(lendaRe) <> "" And Trim(lendaRe) <> lendaVjeter) Then
        Me.txtLendaRe.Text = lendaRe
        objectInitialization MODIFIKO_LENDE
        Me.txtLendaRe.Text = ""
        objectInitialization INSTRUMENTE_LENDE_KLASA_KLIK
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
   first = True
   loadForm Me
   mbushKomboBox
   viti
   Image1.Picture = LoadPicture(adresaLogo)
   'cboKlasa.Text = 1
   lblShkolla.Caption = emerShkolla
   percaktoTeDrejtat
   listaNrU.Visible = False
   listaNrM.Visible = False
   listaNumri.Visible = False
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


Private Sub lendetEmesme()
    lista1.AddItem "Letersi dhe gjuhe shqipe"
    lista1.AddItem "Anglisht"
    lista1.AddItem "Histori"
    lista1.AddItem "Edukim artistik"
    lista1.AddItem "Gjeografi"
    lista1.AddItem "Njohuri per shoqerine"
    lista1.AddItem "Njohuri per ekonomine"
    lista1.AddItem "Psikologji"
    lista1.AddItem "Filozofi"
    lista1.AddItem "Matematike"
    lista1.AddItem "Fizike"
    lista1.AddItem "Kimi"
    lista1.AddItem "Biologji"
    lista1.AddItem "Teknologji"
    lista1.AddItem "Informatike"
    lista1.AddItem "Edukim fizik"
    
    listaNrM.AddItem "1"
    listaNrM.AddItem "2"
    listaNrM.AddItem "3"
    listaNrM.AddItem "4"
    listaNrM.AddItem "5"
    listaNrM.AddItem "6"
    listaNrM.AddItem "7"
    listaNrM.AddItem "8"
    listaNrM.AddItem "9"
    listaNrM.AddItem "10"
    listaNrM.AddItem "11"
    listaNrM.AddItem "12"
    listaNrM.AddItem "13"
    listaNrM.AddItem "14"
    listaNrM.AddItem "15"
    listaNrM.AddItem "16"
    
End Sub

Private Sub lendetTetevjecare()

    lista1.AddItem "Abetare"
    lista1.AddItem "Gjuhe shqipe"
    lista1.AddItem "Lexim letrar"
    lista1.AddItem "Anglisht"
    lista1.AddItem "Histori"
    lista1.AddItem "Dituri natyre"
    lista1.AddItem "Gjeografi"
    lista1.AddItem "Matematike"
    lista1.AddItem "Fizike"
    lista1.AddItem "Kimi"
    lista1.AddItem "Biologji"
    lista1.AddItem "Edukate shoqerore"
    lista1.AddItem "Edukim muzikor"
    lista1.AddItem "Edukim figurativ"
    lista1.AddItem "Mesim pune"
    lista1.AddItem "Edukim fizik"
    
    listaNrU.AddItem "1"
    listaNrU.AddItem "2"
    listaNrU.AddItem "3"
    listaNrU.AddItem "4"
    listaNrU.AddItem "5"
    listaNrU.AddItem "6"
    listaNrU.AddItem "7"
    listaNrU.AddItem "8"
    listaNrU.AddItem "9"
    listaNrU.AddItem "10"
    listaNrU.AddItem "11"
    listaNrU.AddItem "12"
    listaNrU.AddItem "13"
    listaNrU.AddItem "14"
    listaNrU.AddItem "15"
    listaNrU.AddItem "16"
End Sub







Private Sub optMesme_Click()
    
    cboKlasa.Clear
    mbushKomboMesme
    lista1.Clear
    lista2.Clear
    lendetEmesme
    lblProvimet.Caption = "Provimet e pjekurise :"
    listaProvimet.Clear
    listaNumri.Clear
    objectInitialization INSTRUMENTE_VITI_SHKOLLOR_KLIK
    
End Sub

Private Sub optTetevjecare_Click()
    lista1.Clear
    cboKlasa.Clear
    lista2.Clear
    mbushKomboTetevjecare
    lendetTetevjecare
    lblProvimet.Caption = "Provimet e lirimit :"
    listaProvimet.Clear
    listaNumri.Clear
    objectInitialization INSTRUMENTE_VITI_SHKOLLOR_KLIK
    'objectInitialization INSTRUMENTE_LENDE_HIDH_PROVIME
End Sub

Private Function gjendetNeListe(emri As String, liste As ListBox)
    Dim j As Integer
    Dim ugjet As Boolean
    ugjet = False
    j = 0
    Do While (j <= liste.ListCount - 1) And (ugjet = False)
        If LCase(emri) = LCase(liste.List(j)) Then
            ugjet = True
        End If
        j = j + 1
    Loop
    
    gjendetNeListe = ugjet
        
End Function

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


Private Sub percaktoTeDrejtat()
    
    Select Case statusi
        
        Case "SupervizorEmesme"
            optMesme.Visible = False
            optTetevjecare.Visible = False
            optMesme.Value = True
            optTetevjecare.Value = False
            
            fraCikli.Visible = False
            mbushKomboMesme
        
        Case "SupervizorTetevjecare"
            optMesme.Visible = False
            optTetevjecare.Visible = False
            optMesme.Value = False
            optTetevjecare.Value = True
            
            fraCikli.Visible = False
            mbushKomboTetevjecare
        Case Else
            
    End Select
    
            
            
End Sub

Private Sub mbushKomboMesme()
    cboKlasa.Clear
    Dim vitFillimi As Integer
    vitFillimi = CInt(Mid(Me.cboVitiShkollor.Text, 1, 4))
    If (vitFillimi <= 2007) Then
         cboKlasa.AddItem "9"
         'cboKlasa.Text = "9"
    End If
    cboKlasa.AddItem "10"
    cboKlasa.AddItem "11"
    cboKlasa.AddItem "12"
End Sub

Private Sub mbushKomboTetevjecare()
   cboKlasa.Clear
   
   cboKlasa.AddItem "1"
   'cboKlasa.Text = "1"
   cboKlasa.AddItem "2"
   cboKlasa.AddItem "3"
   cboKlasa.AddItem "4"
   cboKlasa.AddItem "5"
   cboKlasa.AddItem "6"
   cboKlasa.AddItem "7"
   cboKlasa.AddItem "8"
   'per vitin shkollor 2008-2009
   Dim vitFillimi As Integer
   vitFillimi = CInt(Mid(Me.cboVitiShkollor.Text, 1, 4))
   If (vitFillimi >= 2008) Then
        cboKlasa.AddItem "9"
   End If
End Sub

