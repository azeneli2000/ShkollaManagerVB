VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmStatistikaNxenes 
   Caption         =   "Statistika - Nxenes"
   ClientHeight    =   9705
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9705
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame fraInfo 
      Height          =   1575
      Left            =   12360
      TabIndex        =   13
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
         TabIndex        =   15
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
         TabIndex        =   14
         Top             =   960
         Width           =   1095
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   120
         Picture         =   "frmStatistikaNxenes.frx":0000
         Stretch         =   -1  'True
         Top             =   600
         Width           =   960
      End
      Begin VB.Label lblWebsite 
         Height          =   375
         Left            =   1560
         TabIndex        =   17
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
         TabIndex        =   16
         Top             =   120
         Width           =   2415
      End
   End
   Begin VB.ComboBox cboOptions 
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
      Left            =   480
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   960
      Width           =   2415
   End
   Begin MSFlexGridLib.MSFlexGrid gridaNxenesit 
      Height          =   3255
      Left            =   6480
      TabIndex        =   5
      Top             =   3480
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   5741
      _Version        =   393216
      Rows            =   13
      Cols            =   3
      FixedCols       =   0
      ScrollBars      =   2
   End
   Begin MSFlexGridLib.MSFlexGrid gridaKlasat 
      Height          =   3255
      Left            =   360
      TabIndex        =   4
      Top             =   3480
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   5741
      _Version        =   393216
      Rows            =   13
      Cols            =   3
      FixedCols       =   0
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin VB.CommandButton cmdShfaq 
      BackColor       =   &H80000009&
      Caption         =   "Shfaq"
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
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   960
      Width           =   2415
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
      Left            =   3240
      TabIndex        =   2
      Text            =   "cboVitiShkollor"
      Top             =   960
      Width           =   2415
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
      TabIndex        =   1
      Top             =   9000
      Width           =   2055
   End
   Begin VB.CommandButton cmdDil 
      BackColor       =   &H80000009&
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
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9000
      Width           =   2295
   End
   Begin VB.Frame fraZgjedhje 
      Caption         =   "Statistika sipas :"
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
      Height          =   1215
      Left            =   240
      TabIndex        =   7
      Top             =   360
      Width           =   8175
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   3000
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "Percakto vitin shkollor :"
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
         Left            =   3000
         TabIndex        =   9
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Percakto semestrin :"
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
         TabIndex        =   8
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame fraStatistika 
      Height          =   5415
      Left            =   120
      TabIndex        =   10
      Top             =   2280
      Width           =   12015
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Mesatarja e nxenesve :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   6360
         TabIndex        =   12
         Top             =   720
         Width           =   5415
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Raporti meshkuj - femra :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   5535
      End
   End
End
Attribute VB_Name = "frmStatistikaNxenes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objektGabimi As New clsErrorHandler
Dim objUIController As clsUIController
Private Sub cmdDil_Click()
    Unload Me
    Set active_form = Nothing
End Sub

Private Sub cmdShfaq_Click()
   'lista1.Clear
   objektGabimi.kapGabimin
   objektGabimi.menazhim_gabimi
   If objektGabimi.mvarGabimi = 0 Then
      pastroGriden gridaKlasat
      pastroGriden gridaNxenesit
      objectInitialization STATISTIKA_NXENES_KLASAT
      'renditListe lista1
      renditGriden gridaKlasat
      formatoGrida gridaKlasat, 3255
   End If
End Sub

Private Sub Command1_Click()
    CallHelp indeksHelp
End Sub



Private Sub Form_Load()
  loadForm Me
  viti
  
  Image1.Picture = LoadPicture(adresaLogo)
  lblShkolla.Caption = emerShkolla
  loadComboItems
  mbushKomboBox
  inicializoGridat
End Sub


Private Sub loadComboItems()
   
   With cboOptions
        .Clear
        
        .AddItem "Semestri I"
        .AddItem "Semestri II"
        .AddItem "Vjetore"
        
        .ListIndex = -1
        
   End With
   
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
    
    
End Sub

Private Sub objectInitialization(actionName As GUI_ACTION__ENUM)
   Set objUIController = New clsUIController

   objUIController.actionName = actionName
   objUIController.ExecuteActions

   Set objUIController = Nothing
End Sub

Private Sub gridaKlasat_Click()
   Dim i                As Integer
   Dim Klasa            As String
   i = gridaKlasat.RowSel
   gridaKlasat.row = i
   gridaKlasat.col = 0
   Klasa = gridaKlasat.Text
   If i <> 0 Then
      If Klasa <> "" Then
         pastroGriden gridaNxenesit
         objectInitialization STATISTIKA_NXENES_MESATARE
         formatoGrida gridaNxenesit, 3255
      End If
   End If

End Sub

Private Sub lista1_Click()
    lista2.Clear
    lista3.Clear
    lista4.Clear
    objectInitialization STATISTIKA_NXENES_MESATARE
End Sub


Private Sub renditListe(liste As ListBox)
   Dim vektorStr(100)   As String
   Dim rreshti          As String
   Dim vektorKlasa()    As Double
   Dim klasaI           As Double
   Dim klasaJ           As Double
   Dim ruajRreshtin     As String
   Dim vektorIndeksi()  As String
   Dim indeksiI         As String
   Dim indeksiJ         As String
   Dim i, j             As Integer

   For i = 0 To liste.ListCount - 1
      vektorStr(i + 1) = liste.List(i)
      ' vektorKlasa(i + 1) = ktheKlase(liste.List(i))
      ' vektorIndeksi(i + 1) = ktheIndeksi(liste.List(i))
   Next i
   Dim nr               As Integer
   nr = liste.ListCount
   For i = 1 To nr - 1
      For j = i + 1 To nr
         klasaI = ktheKlase(vektorStr(i))
         klasaJ = ktheKlase(vektorStr(j))
         If klasaI > klasaJ Then
            ruajRreshtin = vektorStr(i)
            vektorStr(i) = vektorStr(j)
            vektorStr(j) = ruajRreshtin
         Else
            If klasaI = klasaJ Then
               indeksiI = ktheIndeksi(vektorStr(i))
               indeksiJ = ktheIndeksi(vektorStr(j))
               If indeksiI > indeksiJ Then
                  ruajRreshtin = vektorStr(i)
                  vektorStr(i) = vektorStr(j)
                  vektorStr(j) = ruajRreshtin
               End If
            End If
         End If
      Next j
   Next i
   liste.Clear
   For i = 1 To nr
      liste.AddItem vektorStr(i)
   Next i



End Sub

Private Function ktheKlase(rreshti As String) As Double
   Dim Klasa            As String
   s3 = Trim(rreshti)
   Dim s2               As String
   s2 = ""
   Dim ugjet            As Boolean
   Dim i, j             As Integer
   ugjet = False
   i = 1
   Do While Not ugjet
      If Mid(s3, i, 1) <> " " Then
         s2 = s2 & Mid(s3, i, 1)
         i = i + 1
      Else
         ugjet = True
      End If
   Loop

   Klasa = ""
   ugjet = False
   j = 1
   Do While Not ugjet
      If Mid(s2, j, 1) <> "-" Then
         Klasa = Klasa & Mid(s2, j, 1)
         j = j + 1
      Else
         ugjet = True
      End If
   Loop

   ktheKlase = Val(Klasa)
End Function


Private Function ktheIndeksi(rreshti As String) As String

   Dim s3               As String
   s3 = Trim(rreshti)
   Dim Klasa            As String
   Dim indeksi          As String
   Dim s2               As String
   Dim i, j             As Integer
   s2 = ""
   ugjet = False
   i = 1
   Do While Not ugjet
      If Mid(s3, i, 1) <> " " Then
         s2 = s2 & Mid(s3, i, 1)
         i = i + 1
      Else
         ugjet = True
      End If
   Loop

   Klasa = ""
   ugjet = False
   j = 1
   Do While Not ugjet
      If Mid(s2, j, 1) <> "-" Then
         Klasa = Klasa & Mid(s2, j, 1)
         j = j + 1
      Else
         ugjet = True
      End If
   Loop
   j = j + 1
   indeksi = Mid(s2, j)
   ktheIndeksi = indeksi
End Function

Private Sub inicializoGridat()
    
   
    gridaKlasat.row = 0
    gridaKlasat.col = 0
    gridaKlasat.ColWidth(0) = 1500
    gridaKlasat.ColWidth(1) = 2000
    gridaKlasat.ColWidth(2) = 2000
   
    gridaKlasat.CellFontSize = 10
    gridaKlasat.CellAlignment = vbCenter
    gridaKlasat.Text = "Klasa"
    gridaKlasat.col = 1
    
    gridaKlasat.CellFontSize = 10
    gridaKlasat.CellAlignment = vbCenter
    gridaKlasat.Text = "Meshkuj"
    gridaKlasat.col = 2
    
    gridaKlasat.CellFontSize = 10
    gridaKlasat.CellAlignment = vbCenter
    gridaKlasat.Text = "Femra"
    
    
    
    
    gridaNxenesit.row = 0
    gridaNxenesit.col = 0
    gridaNxenesit.CellFontSize = 10
    gridaNxenesit.CellAlignment = vbCenter
    gridaNxenesit.Text = "Emri"
    gridaNxenesit.col = 1
    gridaNxenesit.CellFontSize = 10
    gridaNxenesit.CellAlignment = vbCenter
    gridaNxenesit.Text = "Mbiemri"
    gridaNxenesit.col = 2
    gridaNxenesit.CellFontSize = 10
    gridaNxenesit.CellAlignment = vbCenter
    gridaNxenesit.Text = "Mesatarja"
    gridaNxenesit.ColWidth(0) = 1600
    gridaNxenesit.ColWidth(1) = 1600
    gridaNxenesit.ColWidth(2) = 2300
End Sub

Private Sub renditGriden(grida As MSFlexGrid)

   Dim gjatesia         As Integer
   Dim i                As Integer
   Dim j                As Integer
   Dim klasaI           As String
   Dim klasaJ           As String
   Dim ruajKlasa        As String
   Dim indeksiI         As String
   Dim indeksiJ         As String
   Dim ruajIndeksi      As String
   Dim meshkujI         As String
   Dim femraI           As String

   Dim meshkujJ         As String
   Dim femraJ           As String

   gjatesia = grida.Rows
   For i = 1 To gjatesia - 2
      For j = i + 1 To gjatesia - 1

         gridaKlasat.col = 0
         gridaKlasat.row = i
         klasaI = gridaKlasat.Text
         gridaKlasat.row = j
         klasaJ = gridaKlasat.Text

         If Val(klasaI) > Val(klasaJ) Then
            ruajKlasa = klasaI
            klasaI = klasaJ
            klasaJ = ruajKlasa
            gridaKlasat.col = 0
            gridaKlasat.row = i
            gridaKlasat.Text = klasaI
            gridaKlasat.row = j
            gridaKlasat.Text = klasaJ

            gridaKlasat.col = 1
            gridaKlasat.row = i
            meshkujI = gridaKlasat.Text
            grida.row = j
            meshkujJ = gridaKlasat.Text
            gridaKlasat.row = i
            gridaKlasat.Text = meshkujJ
            grida.row = j
            gridaKlasat.Text = meshkujI

            gridaKlasat.col = 2
            gridaKlasat.row = i
            femraI = gridaKlasat.Text
            grida.row = j
            femraJ = gridaKlasat.Text
            gridaKlasat.row = i
            gridaKlasat.Text = femraJ
            grida.row = j
            gridaKlasat.Text = femraI


         ElseIf Val(klasaI) = Val(klasaJ) Then
            If indeksiI > indeksiJ Then
               ruajKlasa = klasaI
               klasaI = klasaJ
               klasaJ = ruajKlasa
               gridaKlasat.col = 0
               gridaKlasat.row = i
               gridaKlasat.Text = klasaI
               gridaKlasat.row = j
               gridaKlasat.Text = klasaJ

               gridaKlasat.col = 1
               gridaKlasat.row = i
               meshkujI = gridaKlasat.Text
               grida.row = j
               meshkujJ = gridaKlasat.Text
               gridaKlasat.row = i
               gridaKlasat.Text = meshkujJ
               grida.row = j
               gridaKlasat.Text = meshkujI

               gridaKlasat.col = 2
               gridaKlasat.row = i
               femraI = gridaKlasat.Text
               grida.row = j
               femraJ = gridaKlasat.Text
               gridaKlasat.row = i
               gridaKlasat.Text = femraJ
               grida.row = j
               gridaKlasat.Text = femraI

            End If
         Else
         End If
      Next j
   Next i

End Sub

Private Sub formatoGrida(grida As MSFlexGrid, lartesia As Long)
    
    Dim l As Long
    Dim i As Integer
    i = grida.Rows
    l = i * grida.RowHeight(0) + 100
    If l < lartesia Then
        grida.Height = l
    End If
    
    
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

Private Sub pastroGriden(grida As MSFlexGrid)
    
    Dim i As Integer
    Dim j As Integer
    Dim nr_rreshta As Integer
    Dim nr_shtylla As Integer
    nr_rreshta = grida.Rows
    nr_shtylla = grida.Cols
    For i = 1 To nr_rreshta - 1
        For j = 0 To nr_shtylla - 1
            grida.row = i
            grida.col = j
            grida.Text = ""
        Next j
    Next i
    
End Sub
