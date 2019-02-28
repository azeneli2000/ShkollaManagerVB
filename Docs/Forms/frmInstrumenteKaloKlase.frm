VERSION 5.00
Begin VB.Form frmInstrumenteKaloKlase 
   Caption         =   "Kalo Klase"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15000
   WindowState     =   2  'Maximized
   Begin VB.Frame fraCikli 
      Caption         =   "Cikli"
      ForeColor       =   &H00008000&
      Height          =   1455
      Left            =   840
      TabIndex        =   30
      Top             =   120
      Width           =   2415
      Begin VB.OptionButton optUlet 
         Caption         =   "Shkolla tetevjecare"
         ForeColor       =   &H000000C0&
         Height          =   615
         Left            =   240
         TabIndex        =   32
         Top             =   720
         Width           =   1815
      End
      Begin VB.OptionButton optMesme 
         Caption         =   "Shkolla e mesme"
         ForeColor       =   &H00800000&
         Height          =   615
         Left            =   240
         TabIndex        =   31
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.ListBox listaAmza 
      Appearance      =   0  'Flat
      Height          =   2760
      Left            =   5880
      TabIndex        =   29
      Top             =   1800
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame fraInfo 
      Height          =   1695
      Left            =   12240
      TabIndex        =   24
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
         TabIndex        =   26
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
         TabIndex        =   25
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
         TabIndex        =   28
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
         TabIndex        =   27
         Top             =   120
         Width           =   2415
      End
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
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6240
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
      Height          =   4350
      Left            =   7560
      MultiSelect     =   2  'Extended
      TabIndex        =   13
      Top             =   4080
      Width           =   4335
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
      Height          =   4350
      Left            =   840
      MultiSelect     =   1  'Simple
      TabIndex        =   11
      Top             =   4080
      Width           =   4695
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
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8880
      Width           =   2055
   End
   Begin VB.ComboBox cboVitiShkollor2 
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
      Left            =   7800
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   3000
      Width           =   2055
   End
   Begin VB.ComboBox cboVitiShkollor 
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
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   3000
      Width           =   2055
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
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Frame fraKlasa 
      Caption         =   "Klasa e re :"
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
      Height          =   1935
      Left            =   7560
      TabIndex        =   2
      Top             =   1680
      Width           =   4335
      Begin VB.ComboBox cboIndeksi2 
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
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   600
         Width           =   1455
      End
      Begin VB.ComboBox cboKlasa2 
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
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   600
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
         Left            =   240
         TabIndex        =   23
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Indeksi :"
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
         Left            =   2040
         TabIndex        =   6
         Top             =   360
         Width           =   1455
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
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1455
      End
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
      TabIndex        =   1
      Top             =   8880
      Width           =   2655
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
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8880
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Caption         =   "Klasa e vjeter :"
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
      Height          =   1935
      Left            =   840
      TabIndex        =   16
      Top             =   1680
      Width           =   4695
      Begin VB.ComboBox cboKlasa 
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
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   600
         Width           =   1455
      End
      Begin VB.ComboBox cboIndeksi 
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
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton cmdKerko 
         BackColor       =   &H80000009&
         Caption         =   "Kerko"
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
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label6 
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
         TabIndex        =   22
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label1 
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
         TabIndex        =   21
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblIndeksi 
         Caption         =   "Indeksi :"
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
         Left            =   1920
         TabIndex        =   20
         Top             =   360
         Width           =   1815
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
      TabIndex        =   14
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
      TabIndex        =   12
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



Private Sub cmdDil_Click()
   Unload Me
   Set active_form = Nothing
End Sub

Private Sub cmdKaloNeListe_Click()
   objektGabimi.kapGabimin
   If objektGabimi.mvarGabimi = 15 Then
      objektGabimi.mvarGabimi = 16
   End If
   objektGabimi.menazhim_gabimi
   If objektGabimi.mvarGabimi = 0 Then
      objectInitialization INSTRUMENTE_NXENES_KALO_NE_LISTE
   End If
End Sub

Private Sub cmdKerko_Click()
   lista1.Clear
   objektGabimi.kapGabimin
   If objektGabimi.mvarGabimi = 16 Or objektGabimi.mvarGabimi = 0 Then

     
         objectInitialization INSTRUMENTE_KALO_KLASE_KERKO
     
   Else
   objektGabimi.menazhim_gabimi
   End If
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

   Dim klasa1           As String
   Dim klasa2           As String
   Dim vitiShkollor1    As String
   Dim vitiShkollor2    As String
   klasa1 = cboKlasa.Text
   klasa2 = cboKlasa2.Text
   vitiShkollor1 = cboVitiShkollor.Text
   vitiShkollor2 = cboVitiShkollor2.Text
   Dim nr1              As Double
   Dim nr2              As Double
   nr1 = Val(klasa1)
   nr2 = Val(klasa2)
   Dim i, j             As Integer
   i = cboVitiShkollor.ListIndex
   j = cboVitiShkollor2.ListIndex
   If i > j Then
      MsgBox "Ju nuk mund te regjistroni nje nxenes ne nje vit shkollor paraardhes.", , "Konfigurimi i klasave."
   ElseIf (j > i + 1) Then
      MsgBox "Ju nuk mund te regjistroni nje nxenes ne nje vit shkollor me te madh se viti shkollor pasaradhes", , "Konfigurimi i klasave "
   Else
      If (klasa1 = klasa2) Or (nr2 = nr1 + 1) Then
         objectInitialization INSTRUMENTE_KALO_KLASE
      Else
         If nr1 > nr2 Then
            MsgBox "Ju nuk mund te kaloni nje nxenes ne nje klase me te vogel.", , "Konfigurimi i klasave."
         End If
         If nr1 + 1 < nr2 Then
            MsgBox "Ju nuk mund te kaloni nje nxenes ne nje klase me te larte se klasa pasardhese.", , "Konfigurimi i klasave."

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
   optMesme.Value = True
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


Private Function gjendetNeListe(emri As String, liste As ListBox)
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
    
    cboKlasa2.Clear
    cboKlasa2.AddItem "1"
    cboKlasa2.AddItem "2"
    cboKlasa2.AddItem "3"
    cboKlasa2.AddItem "4"
    cboKlasa2.AddItem "5"
    cboKlasa2.AddItem "6"
    cboKlasa2.AddItem "7"
    cboKlasa2.AddItem "8"
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
End Sub

Private Sub optUlet_Click()
    mbushKomboUlet
End Sub
