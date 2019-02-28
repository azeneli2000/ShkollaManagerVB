VERSION 5.00
Begin VB.Form frmInstrumenteLende 
   Caption         =   "Instrumente - Lende"
   ClientHeight    =   8565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14685
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
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
   WindowState     =   2  'Maximized
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
         TabIndex        =   27
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
         TabIndex        =   28
         Top             =   120
         Width           =   2415
      End
   End
   Begin VB.CommandButton cmdPrano 
      BackColor       =   &H80000009&
      Caption         =   "Ok"
      Height          =   375
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6360
      Width           =   2295
   End
   Begin VB.CommandButton cmdHiq 
      BackColor       =   &H80000009&
      Caption         =   "Elimino provim"
      Height          =   375
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5760
      Width           =   2295
   End
   Begin VB.CommandButton cmdShto 
      BackColor       =   &H80000009&
      Caption         =   "Shto provim"
      Height          =   375
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5160
      Width           =   2295
   End
   Begin VB.ListBox listaProvimet 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1950
      Left            =   11280
      TabIndex        =   14
      Top             =   2640
      Width           =   2775
   End
   Begin VB.OptionButton optTetevjecare 
      Caption         =   "Shkolla tetevjecare"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   600
      TabIndex        =   13
      Top             =   1080
      Width           =   1575
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
      Left            =   600
      TabIndex        =   12
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton cmdKtheMbrapsht 
      BackColor       =   &H80000009&
      Caption         =   "<<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4440
      Width           =   1695
   End
   Begin VB.ListBox lista2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3630
      Left            =   6240
      MultiSelect     =   2  'Extended
      TabIndex        =   10
      Top             =   2640
      Width           =   3375
   End
   Begin VB.ListBox lista1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3870
      Left            =   480
      MultiSelect     =   2  'Extended
      TabIndex        =   9
      Top             =   2640
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
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
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7920
      Width           =   2295
   End
   Begin VB.ComboBox cboVitiShkollor 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Height          =   360
      Left            =   6360
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1320
      Width           =   2055
   End
   Begin VB.CommandButton cmdShtoLende 
      BackColor       =   &H80000009&
      Caption         =   "Shto Lende"
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
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5160
      Width           =   1695
   End
   Begin VB.CommandButton cmdDil 
      BackColor       =   &H8000000E&
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
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7920
      Width           =   2655
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H8000000E&
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
      Left            =   3600
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
      Width           =   3135
      Begin VB.ComboBox cboKlasa 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Height          =   360
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Viti shkollor :"
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
            Size            =   9.75
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
      Caption         =   ">>>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Frame fraCikli 
      Caption         =   "Cikli"
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
      Height          =   1695
      Left            =   480
      TabIndex        =   18
      Top             =   120
      Width           =   2055
   End
   Begin VB.Frame fraLendet 
      Caption         =   "Konfigurimi i lendeve :"
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
      Height          =   5655
      Left            =   120
      TabIndex        =   20
      Top             =   1920
      Width           =   9855
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Lendet e klases se zgjedhur :"
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   6120
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
         Left            =   360
         TabIndex        =   21
         Top             =   360
         Width           =   3495
      End
   End
   Begin VB.Frame fraProvimet 
      Caption         =   "Konfigurimi i provimeve :"
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
 

Private Sub cboKlasa_Click()
  If cboVitiShkollor.Text <> "" Then
     objectInitialization INSTRUMENTE_LENDE_KLASA_KLIK
  Else
     MsgBox "Per te pare lendet qe zhvillon kjo klase ju duhet te zgjidhni vitin shkollor perkates.", , "Konfigurimi i lendeve."
  End If
  
    
    
End Sub





Private Sub cboVitiShkollor_Click()
   listaProvimet.Clear
   objectInitialization INSTRUMENTE_VITI_SHKOLLOR_KLIK
End Sub

Private Sub cmdDil_Click()

   Unload Me
   Set active_form = Nothing
   
End Sub


Private Sub cmdHiq_Click()
   Dim i                As Integer
   Dim j                As Integer
   j = listaProvimet.ListIndex

   If j > -1 Then
      listaProvimet.RemoveItem j
   End If
End Sub

Private Sub cmdKaloNeListe_Click()
    Dim i As Integer
    Dim emri As String
    For i = 0 To lista1.ListCount - 1
        If (lista1.Selected(i) = True) Then
            emri = lista1.List(i)
            If Not gjendetNeListe(emri, lista2) Then
                lista2.AddItem emri
                
            End If
        End If
    Next i
End Sub

Private Sub cmdKtheMbrapsht_Click()
     Dim i As Integer
     i = 0
     Do While (i <= lista2.ListCount - 1)
        If lista2.Selected(i) Then
            lista2.RemoveItem (i)
            i = i - 1
        End If
        i = i + 1
     Loop
End Sub

Private Sub cmdOK_Click()
   Dim Klasa            As String
   Klasa = cboKlasa.Text
   If Klasa <> "" Then
      objectInitialization INSTRUMENTE_LENDE
      MsgBox "Lendet per klasen e zgjedhur nga ju u modifikuan.", , "Konfigurimi i lendeve."
   Else
      MsgBox "Ju duhet te jepni me pare numrin e klases.", , "Konfigurimi i lendeve."
   End If

End Sub



Private Sub cmdPrano_Click()
  If cboVitiShkollor.Text = "" Or (optMesme.Value And optTetevjecare.Value) Then
    MsgBox "Ju duhet te percaktoni vitin shkollor dhe ciklin e shkolles para se te hidhni provimet.", , "Konfigurimi i lendeve."
  Else
    objectInitialization INSTRUMENTE_LENDE_HIDH_PROVIME
    MsgBox "Provimet u hodhen.", , "Konfigurimi  lendeve."
    'listaProvimet.Clear
  End If
End Sub

Private Sub cmdShto_Click()

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
   If cboVitiShkollor.Text = "" Or cboKlasa.Text = "" Then
      MsgBox "Ju duhet te jepni te dhenat identifikuese te klases , te cilat jane numri i klases dhe viti shkollor , para se te shtoni nje lende ne kete klase.", , "Kalo lende."
   Else
      Dim lenda            As String
      lenda = InputBox("Jepni emrin e lendes qe doni te shtoni", "Lenda e re")
      If lenda <> "" Then
         If Not gjendetNeListe(lenda, lista2) Then
            lista2.AddItem lenda
         End If

      End If
   End If
End Sub

Private Sub Command1_Click()
    CallHelp indeksHelp
End Sub

Private Sub Form_Load()
   loadForm Me
   mbushKomboBox
   viti
   Image1.Picture = LoadPicture(adresaLogo)
   lblShkolla.Caption = emerShkolla
   
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
    lista1.AddItem "Gjuhe e huaj"
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
    lista1.AddItem "Edukim Fizik"
    
End Sub

Private Sub lendetTetevjecare()

    lista1.AddItem "Abetare"
    lista1.AddItem "Gjuhe shqipe"
    lista1.AddItem "Lexim letrar"
    lista1.AddItem "Gjuhe e huaj"
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
End Sub



Private Sub optMesme_Click()
    
    cboKlasa.Clear
    mbushKomboKlasaMesme
    lista1.Clear
    lista2.Clear
    lendetEmesme
    lblProvimet.Caption = "Provimet e pjekurise :"
    listaProvimet.Clear
    objectInitialization INSTRUMENTE_VITI_SHKOLLOR_KLIK
    
End Sub

Private Sub optTetevjecare_Click()
    lista1.Clear
    cboKlasa.Clear
    lista2.Clear
    mbushKomboKlasaTetevjecare
    lendetTetevjecare
    lblProvimet.Caption = "Provimet e lirimit :"
    listaProvimet.Clear
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

Private Sub mbushKomboKlasaTetevjecare()
    
    cboKlasa.AddItem "1"
    cboKlasa.AddItem "2"
    cboKlasa.AddItem "3"
    cboKlasa.AddItem "4"
    cboKlasa.AddItem "5"
    cboKlasa.AddItem "6"
    cboKlasa.AddItem "7"
    cboKlasa.AddItem "8"
End Sub

Private Sub mbushKomboKlasaMesme()
    
    cboKlasa.AddItem "9"
    cboKlasa.AddItem "10"
    cboKlasa.AddItem "11"
    cboKlasa.AddItem "12"
End Sub

Private Sub viti()
   Dim d, muai, viti    As String
   d = Now
   Dim m, v             As Integer
   muai = Mid(d, 4, 2)
   viti = Mid(d, 7, 4)
   m = Val(muai)
   v = Val(viti)
   cboVitiShkollor.Text = gjej_vitin(m, v)

End Sub
