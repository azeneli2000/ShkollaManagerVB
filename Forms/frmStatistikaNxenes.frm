VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
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
   Begin VB.TextBox txtKlasaZgjedhur 
      Height          =   375
      Left            =   3600
      TabIndex        =   23
      Text            =   "Text3"
      Top             =   2880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H80000009&
      DisabledPicture =   "frmStatistikaNxenes.frx":0000
      DownPicture     =   "frmStatistikaNxenes.frx":353A
      Height          =   375
      Left            =   4440
      Picture         =   "frmStatistikaNxenes.frx":6A74
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   9240
      Width           =   2535
   End
   Begin VB.Frame fraInfo 
      Height          =   1575
      Left            =   12360
      TabIndex        =   11
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
         TabIndex        =   13
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
         TabIndex        =   12
         Top             =   960
         Width           =   1095
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   120
         Picture         =   "frmStatistikaNxenes.frx":9FAE
         Stretch         =   -1  'True
         Top             =   600
         Width           =   960
      End
      Begin VB.Label lblWebsite 
         Height          =   375
         Left            =   1560
         TabIndex        =   15
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
         TabIndex        =   14
         Top             =   120
         Width           =   2415
      End
   End
   Begin VB.ComboBox cboOptions 
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
      Left            =   480
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   960
      Width           =   2415
   End
   Begin VB.CommandButton cmdShfaq 
      BackColor       =   &H80000009&
      Default         =   -1  'True
      DisabledPicture =   "frmStatistikaNxenes.frx":10F23
      DownPicture     =   "frmStatistikaNxenes.frx":17365
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
      Left            =   6240
      Picture         =   "frmStatistikaNxenes.frx":1D7A7
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   960
      Width           =   1815
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
      Left            =   3240
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   960
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000009&
      DisabledPicture =   "frmStatistikaNxenes.frx":23BE9
      DownPicture     =   "frmStatistikaNxenes.frx":27123
      Height          =   375
      Left            =   11760
      Picture         =   "frmStatistikaNxenes.frx":2A65D
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9240
      Width           =   2535
   End
   Begin VB.CommandButton cmdDil 
      BackColor       =   &H80000009&
      Height          =   375
      Left            =   8160
      Picture         =   "frmStatistikaNxenes.frx":2DB97
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9240
      Width           =   2535
   End
   Begin VB.Frame fraZgjedhje 
      Caption         =   "Statistika sipas :"
      ForeColor       =   &H00008000&
      Height          =   1215
      Left            =   240
      TabIndex        =   5
      Top             =   360
      Width           =   8175
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   3000
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "Viti Shkollor :"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   3000
         TabIndex        =   7
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Semestri :"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame fraStatistika 
      Height          =   5535
      Left            =   240
      TabIndex        =   8
      Top             =   3480
      Width           =   12015
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridaNxenesitLabel 
         Height          =   300
         Left            =   6480
         TabIndex        =   21
         Top             =   720
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridaLabelKlasat 
         Height          =   375
         Left            =   720
         TabIndex        =   20
         Top             =   720
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   661
         _Version        =   393216
         Rows            =   1
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   1
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   240
         TabIndex        =   22
         Top             =   4560
         Width           =   11295
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridaKlasat 
         Height          =   3375
         Left            =   720
         TabIndex        =   17
         Top             =   1080
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   5953
         _Version        =   393216
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridaNxenesit 
         Height          =   3375
         Left            =   6480
         TabIndex        =   18
         Top             =   1080
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   5953
         _Version        =   393216
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
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
         Left            =   6240
         TabIndex        =   10
         Top             =   120
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
         TabIndex        =   9
         Top             =   120
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
Dim objPrintClass As Object
Dim objPrintForm As New frmPrintStatistikaNxenes
Dim Klasa As String

Private Sub cmdDil_Click()
    Unload Me
    Set active_form = Nothing
End Sub

Private Sub cmdEmail_Click()
    If Not OpenEmailProgram(email) Then
        MsgBox "Ju nuk keni asnje program te instaluar per te derguar email", vbExclamation, "Gabim ne dergimin e emailit"
    End If
End Sub


Private Sub cmdPrint_Click()
    Dim zgjedhja As Integer
    objPrintForm.show vbModal
    zgjedhja = objPrintForm.choice
    Select Case zgjedhja
        Case 1
            PrintoRaportin
        Case 2
            PrintoMesataren
    End Select

End Sub

Private Sub cmdShfaq_Click()
   'lista1.Clear
   gridaKlasat.Visible = False
   objektGabimi.kapGabimin
   objektGabimi.menazhim_gabimi
   If objektGabimi.mvarGabimi = 0 Then
      pastroGriden gridaKlasat
      pastroGriden gridaNxenesit
      gridaNxenesit.Rows = 0
      gridaNxenesit.Height = 0
      inicializoGridaNxenesit gridaKlasat, gridaLabelKlasat, 12
      objectInitialization STATISTIKA_NXENES_KLASAT
      'renditListe lista1
      
      renditGriden gridaKlasat
      Dim vitFillimi As Integer
      vitFillimi = CInt(Mid(Me.cboVitiShkollor.Text, 1, 4))
      If (vitFillimi <= 2007) Then
        shtoStatistikat
      Else
        shtoStatistikat1
      End If
      ngjyrosGriden gridaKlasat
      formatoGrida gridaKlasat, 3240
      
   End If
   gridaKlasat.Visible = True
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



Private Sub Command2_Click()
    
End Sub

Private Sub Form_Load()
  loadForm Me
  mbushKomboBox
  viti
  
  Image1.Picture = LoadPicture(adresaLogo)
  lblShkolla.Caption = emerShkolla
  loadComboItems
  inicializoGridaMestarjaNxenesit gridaNxenesit, gridaNxenesitLabel, 12
  inicializoGridaNxenesit gridaKlasat, gridaLabelKlasat, 12
  ngjyrosGriden gridaKlasat
  ngjyrosGriden gridaNxenesit
  
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
  fraInfo.left = Me.left + Me.Width - fraInfo.Width - 200
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
    
    
End Sub

Private Sub objectInitialization(actionName As GUI_ACTION__ENUM)
   Set objUIController = New clsUIController

   objUIController.actionName = actionName
   objUIController.ExecuteActions

   Set objUIController = Nothing
End Sub

Private Sub gridaKlasat_Click()
   Dim i                As Integer
   i = gridaKlasat.RowSel
   gridaKlasat.row = i
   gridaKlasat.col = 0
   Klasa = gridaKlasat.Text
   Me.txtKlasaZgjedhur.Text = Klasa
   If Klasa <> "" Then
      If Not (Klasa = "Totali" Or Klasa = "Cikli i ulet" Or Klasa = "E mesmja" Or Klasa = "Tetevjecarja" Or Klasa = "Nëntëvjeçarja") Then
         gridaNxenesit.Visible = False
         pastroGriden gridaNxenesit
         objectInitialization STATISTIKA_NXENES_MESATARE
         ngjyrosGriden gridaNxenesit
         formatoGrida gridaNxenesit, 3240
         gridaNxenesit.Visible = True
         gridaNxenesitLabel.Visible = True
      Else
        gridaNxenesit.Visible = False
        gridaNxenesitLabel.Visible = False
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
   
   Dim i As Integer
   Dim j As Integer
   gridaKlasat.Cols = 3
   gridaNxenesit.Cols = 3
   gridaKlasat.FixedCols = 0
   gridaNxenesit.FixedCols = 0
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

   gridaKlasat.ScrollBars = flexScrollBarNone
   gridaNxenesit.ScrollBars = flexScrollBarNone

   gridaKlasat.Rows = 16
   gridaNxenesit.Rows = 16

   For i = 0 To gridaKlasat.Rows - 1
      For j = 0 To gridaKlasat.Cols - 1
         gridaKlasat.row = i
         gridaKlasat.col = j
         gridaKlasat.CellAlignment = vbCenter
      Next j
   Next i

   For i = 0 To gridaNxenesit.Rows - 1
      For j = 0 To gridaNxenesit.Cols - 1
         gridaNxenesit.row = i
         gridaNxenesit.col = j
         gridaNxenesit.CellAlignment = vbCenter
      Next j
   Next i
End Sub

Private Sub renditGriden(grida As MSHFlexGrid)

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
   For i = 0 To gjatesia - 2
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

Private Sub formatoGrida(grida As MSHFlexGrid, lartesia As Long)

   Dim l                As Long
   Dim i                As Integer
   i = grida.Rows
   l = i * CLng(270)
   If l < lartesia Then
      grida.Height = l
      grida.ScrollBars = flexScrollBarNone
   Else
      grida.Height = lartesia
      grida.ScrollBars = flexScrollBarVertical
   End If

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


Private Sub pastroGriden(grida As MSHFlexGrid)
    
    Dim i As Integer
    Dim j As Integer
    Dim nr_rreshta As Integer
    Dim nr_shtylla As Integer
    nr_rreshta = grida.Rows
    nr_shtylla = grida.Cols
    For i = 0 To nr_rreshta - 1
        For j = 0 To nr_shtylla - 1
            grida.row = i
            grida.col = j
            grida.Text = ""
        Next j
    Next i
    
End Sub

Private Sub ngjyrosGriden(grida As MSHFlexGrid)

   Dim i                As Integer
   Dim j                As Integer
   For i = 0 To grida.Rows - 1
      If i Mod 2 = 1 Then
         For j = 0 To grida.Cols - 1
            grida.col = j
            grida.row = i
            grida.CellBackColor = &HE0E0E0
         Next j
      End If
   Next i
End Sub

Private Sub PrintoRaportin()
    Dim semestri As String
    Dim tekst As String
    Dim fontiBold As Boolean
    Dim i, j, p  As Integer
    Dim maxRows          As Integer
    Dim currPage As Integer, totalPages As Integer
    maxRows = 0
    If gridaKlasat.Rows Mod 26 = 0 Then
        totalPages = gridaKlasat.Rows / 26
    Else
        totalPages = Fix(gridaKlasat.Rows / 26) + 1
    End If
    currPage = 1
    p = 0
    Set objPrintClass = CreateObject("PrintimComponent.clsPrintim")
    If objPrintClass.PrinterIsInstalled = False Then
        Exit Sub
    End If
        objPrintClass.PrintFont "Times New Roman"
        objPrintClass.OrientimFaqe True
        If objPrintClass.Gabim = True Then
            MsgBox "Printeri juaj nuk lejon qe te nderrohet ne menyre automatike formati ne Portrait. " & _
                "Nderrojeni formatin nga Portrait ne Landscape te preferencat e printerit tuaj.", vbCritical + vbOKOnly, _
                "Gabim ne Printim."
            Exit Sub
        End If
        gridaKlasat.row = 0
        gridaKlasat.col = 0
        If gridaKlasat.Text <> "" Then
            Do While p < gridaKlasat.Rows - 1
                If gridaKlasat.Rows - 1 - p > 26 Then
                    maxRows = 26
                Else
                    maxRows = gridaKlasat.Rows - p
                End If
                objPrintClass.PrintLeft 50, 60, "Raporti meshkuj - femra për vitin shkollor", True, 12, False
                objPrintClass.PrintLeft 134, 60, cboVitiShkollor.Text & ":", True, 12, False, True
                'objPrintClass.PrintLeft
                ' Nese perdoruesi anullon printimin
                If objPrintClass.Gabim = True Then
                    Exit Sub
                End If
        
                PrintoFormatPergj
                objPrintClass.PrintTabele 55, 70, 40, 14, 1, 1
                objPrintClass.PrintTabele 95, 70, 50, 7, 1, 1
                objPrintClass.PrintTabele 95, 77, 25, 7, 1, 2
                objPrintClass.PrintTabele 55, 84, 40, 6, maxRows, 1
                objPrintClass.PrintTabele 95, 84, 25, 6, maxRows, 2
        
                objPrintClass.PrintLeft 70, 75, "Klasa", True, 12, False
                objPrintClass.PrintLeft 113, 71, "Raporti", True, 12
                objPrintClass.PrintLeft 100, 78, "meshkuj", True, 12
                objPrintClass.PrintLeft 127, 78, "femra", True, 12
                j = 0
                For i = p To maxRows - 1 + p
                    gridaKlasat.col = 0
                    gridaKlasat.row = i
                    tekst = gridaKlasat.Text
                    objPrintClass.PrintLeft 65, 85.1 + 6 * j, tekst, False, 12, False
                    gridaKlasat.col = 1
                    tekst = gridaKlasat.Text
                    objPrintClass.PrintLeft 105, 85.1 + 6 * j, tekst, False, 12, False
                    gridaKlasat.col = 2
                    tekst = gridaKlasat.Text
                    objPrintClass.PrintLeft 130, 85.1 + 6 * j, tekst, False, 12, False
                    j = j + 1
                Next
                p = p + maxRows
                If p < gridaKlasat.Rows Then
                    'maxRows = maxRows + 26
                    objPrintClass.PrintLeft 191, 275, CStr(currPage) & "/" & CStr(totalPages)
                    objPrintClass.NewPage
                End If
            Loop
            objPrintClass.PrintLeft 191, 275, CStr(currPage) & "/" & CStr(totalPages)
            objPrintClass.EndDoc
    Else
        MsgBox "Ju nuk keni perzgjedhur asnje klase per te pare mesataret e lendeve", vbOKOnly + vbInformation, "Zgjidhni klasen"
    End If
End Sub

Private Sub PrintoMesataren()
    Dim tekst As String
    Dim fontiBold As Boolean
    Dim maxRows          As Integer
    Dim totalPages As Integer, currPage As String
    maxRows = 0
    If gridaNxenesit.Rows Mod 26 = 0 Then
        totalPages = gridaNxenesit.Rows / 27
    Else
        totalPages = Fix(gridaNxenesit.Rows / 27) + 1
    End If
    currPage = 1
    p = 0
    Set objPrintClass = CreateObject("PrintimComponent.clsPrintim")
    objPrintClass.PrintFont "Times New Roman"
    objPrintClass.OrientimFaqe True
    If objPrintClass.Gabim = True Then
        MsgBox "Printeri juaj nuk lejon qe te nderrohet ne menyre automatike formati ne Landscape. " & _
            "Nderrojeni formatin nga Landscape ne Portrait te preferencat e printerit tuaj.", vbCritical + vbOKOnly, _
            "Gabim ne Printim."
        Exit Sub
    End If
    gridaNxenesit.row = 0
    gridaNxenesit.col = 0
    If gridaNxenesit.Text <> "" Then
        Do While p <= gridaNxenesit.Rows - 1
            If gridaNxenesit.Rows - 1 - p > 26 Then
                maxRows = 26
            Else
                maxRows = gridaNxenesit.Rows - p
            End If
            objPrintClass.OrientimFaqe True
            objPrintClass.PrintLeft 55, 70, "Klasa ", True, 12, False
            objPrintClass.PrintLeft 68, 70, Klasa, True, 12, False, True
            objPrintClass.PrintLeft 86, 70, "Viti shkollor", True, 12, False, False
            objPrintClass.PrintLeft 111, 70, cboVitiShkollor.Text, True, 12, False, True
            ' Nese perdoruesi anullon printimin
            If objPrintClass.Gabim = True Then
                Exit Sub
            End If
        
            PrintoFormatPergj
            
            objPrintClass.PrintTabele 45, 85, 11, 6, maxRows + 1, 1
            objPrintClass.PrintTabele 56, 85, 60, 6, maxRows + 1, 1
            objPrintClass.PrintTabele 116, 85, 30, 6, maxRows + 1, 1
            
            ' Printimi i kokes se tabeles
            objPrintClass.PrintLeft 47, 86, "Nr", True, 12, False
            objPrintClass.PrintLeft 68, 86, "Emri  Mbiemri", True, 12, False
            objPrintClass.PrintLeft 118, 86, "Mesatarja", True, 12, False
            ' Printimi i emrave dhe mbiemrave dhe notave
            j = 0
            For i = p To maxRows - 1 + p
                objPrintClass.PrintLeft 47, 92 + 6 * j, CStr(i + 1), False, 12
                gridaNxenesit.row = i
                gridaNxenesit.col = 0
                tekst = gridaNxenesit.Text
                gridaNxenesit.col = 1
                tekst = tekst + Space(2) + gridaNxenesit.Text
                objPrintClass.PrintLeft 57, 92 + 6 * j, tekst, False, 12, False
                gridaNxenesit.col = 2
                tekst = gridaNxenesit.Text
                objPrintClass.PrintLeft 121, 92 + 6 * j, tekst, False, 12, False
                j = j + 1
            Next
            p = p + maxRows
            If p < gridaNxenesit.Rows Then
                'maxRows = maxRows + 26
                objPrintClass.NewPage
                objPrintClass.PrintLeft 191, 275, CStr(currPage) & "/" & CStr(totalPages)
                currPage = currPage + 1
            End If
        Loop
        objPrintClass.PrintLeft 191, 275, CStr(currPage) & "/" & CStr(totalPages)
        objPrintClass.EndDoc
    Else
        MsgBox "Ju nuk keni perzgjedhur asnje klase per te pare mesataret e lendeve", vbOKOnly + vbInformation, "Zgjidhni klasen"
    End If
End Sub

Private Sub shtoStatistikat()
   Dim i                As Integer, j As Integer
   Dim Klasa            As String, indeksi As String, rreshti As String
   Dim kaCikliUlet      As Boolean
   Dim kaCikliTetevjecare As Boolean
   Dim kaCikliMesem     As Boolean
   Dim nrRReshtash      As Integer
   Dim lartesia         As Integer
   Dim nrU              As Integer
   Dim nrT              As Integer
   Dim nrM              As Integer
   Dim rrU              As Integer
   Dim rrT              As Integer
   Dim rrM              As Integer
   Dim nrMeshkujM       As Integer, nrFemraM As Integer
   Dim nrMeshkujT       As Integer, nrFemraT As Integer
   Dim nrMeshkujU       As Integer, nrFemraU As Integer
   Dim nrMeshkuj        As Integer, nrFemra As Integer
   Dim kol0             As String, kol1 As String, kol2 As String, kol3 As String
   Dim nrSmes           As Integer
   Dim nrSu             As Integer
   Dim nrSt             As Integer
   Dim nrs              As Integer
   Dim iM               As Integer
   Dim iT               As Integer
   nrsm = 0
   nrSt = 0
   nrSu = 0
   nrs = 0
   lartesia = gridaKlasat.Rows

   If lartesia = 0 Then
      Exit Sub
   End If
   nrU = 0
   nrM = 0
   nrT = 0
   For i = 0 To lartesia - 1
      gridaKlasat.row = i
      gridaKlasat.col = j
      rreshti = gridaKlasat.Text
      Klasa = ktheKlasenString(rreshti)
      If Klasa = "1" Or Klasa = "2" Or Klasa = "3" Or Klasa = "4" Then
         kaCikliUlet = True
         nrU = nrU + 1
         gridaKlasat.col = 1
         nrMeshkujU = nrMeshkujU + CInt(gridaKlasat.Text)
         nrMeshkuj = nrMeshkuj + CInt(gridaKlasat.Text)
         gridaKlasat.col = 2
         nrFemraU = nrFemraU + CInt(gridaKlasat.Text)
         nrFemra = nrFemra + CInt(gridaKlasat.Text)
      End If
      If Klasa = "5" Or Klasa = "6" Or Klasa = "7" Or Klasa = "8" Then

         kaCikliTetevjecare = True
         nrT = nrT + 1

         

         gridaKlasat.col = 1
         nrMeshkujT = nrMeshkujT + CInt(gridaKlasat.Text)
         nrMeshkuj = nrMeshkuj + CInt(gridaKlasat.Text)
         gridaKlasat.col = 2
         nrFemraT = nrFemraT + CInt(gridaKlasat.Text)
         nrFemra = nrFemra + CInt(gridaKlasat.Text)
      End If
      If Klasa = "9" Or Klasa = "10" Or Klasa = "11" Or Klasa = "12" Then

         kaCikliMesem = True
         nrM = nrM + 1

         gridaKlasat.col = 1
         nrMeshkujM = nrMeshkujM + CInt(gridaKlasat.Text)
         nrMeshkuj = nrMeshkuj + CInt(gridaKlasat.Text)
         gridaKlasat.col = 2
         nrFemraM = nrFemraM + CInt(gridaKlasat.Text)
         nrFemra = nrFemra + CInt(gridaKlasat.Text)
      End If

   Next i

   rrU = 0
   rrT = 0
   rrM = 0
   iT = 0
   iM = 0
   If kaCikliUlet Then
      gridaKlasat.Rows = gridaKlasat.Rows + 1
      
      rrU = nrU + 1
      rrT = rrU
      rrM = rrU
      iT = iT + 1
      iM = iM + 1
   End If

   If kaCikliTetevjecare Then
      gridaKlasat.Rows = gridaKlasat.Rows + 1
      
      rrT = rrT + nrT + 1
      rrM = rrM + nrT + 1
      iM = iM + 1
   End If

   If kaCikliMesem Then
      gridaKlasat.Rows = gridaKlasat.Rows + 1
      
      rrM = rrM + nrM + 1
   End If
   gridaKlasat.Rows = gridaKlasat.Rows + 1

   If kaCikliMesem Then

      For i = nrU + nrM + nrT - 1 To nrU + nrT Step -1

         gridaKlasat.row = i
         gridaKlasat.col = 0
         kol0 = gridaKlasat.Text

         gridaKlasat.row = i
         gridaKlasat.col = 1
         kol1 = gridaKlasat.Text

         gridaKlasat.row = i
         gridaKlasat.col = 2
         kol2 = gridaKlasat.Text

         

         gridaKlasat.row = i + iM
         gridaKlasat.col = 0
         gridaKlasat.Text = kol0

         gridaKlasat.row = i + iM
         gridaKlasat.col = 1
         gridaKlasat.Text = kol1

         gridaKlasat.row = i + iM
         gridaKlasat.col = 2
         gridaKlasat.Text = kol2

         
      Next i

      gridaKlasat.row = rrM - 1

      gridaKlasat.col = 0
      gridaKlasat.Text = "E mesmja"
      gridaKlasat.CellForeColor = &HC0&

      gridaKlasat.col = 1
      gridaKlasat.Text = str(nrMeshkujM)
      gridaKlasat.CellForeColor = &HC0&

      gridaKlasat.col = 2
      gridaKlasat.Text = str(nrFemraM)
      gridaKlasat.CellForeColor = &HC0&

     

   End If

   If kaCikliTetevjecare Then

      For i = nrU + nrT - 1 To nrU Step -1

         gridaKlasat.row = i
         gridaKlasat.col = 0
         kol0 = gridaKlasat.Text

         gridaKlasat.row = i
         gridaKlasat.col = 1
         kol1 = gridaKlasat.Text

         gridaKlasat.row = i
         gridaKlasat.col = 2
         kol2 = gridaKlasat.Text

         

         gridaKlasat.row = i + iT
         gridaKlasat.col = 0
         gridaKlasat.Text = kol0

         gridaKlasat.row = i + iT
         gridaKlasat.col = 1
         gridaKlasat.Text = kol1

         gridaKlasat.row = i + iT
         gridaKlasat.col = 2
         gridaKlasat.Text = kol2

         

      Next i

      gridaKlasat.row = rrT - 1
      gridaKlasat.col = 0
      gridaKlasat.Text = "Tetevjecarja"
      gridaKlasat.CellForeColor = &HC0&
      gridaKlasat.col = 1
      gridaKlasat.Text = str(nrMeshkujT)
      gridaKlasat.CellForeColor = &HC0&

      gridaKlasat.col = 2
      gridaKlasat.Text = str(nrFemraT)
      gridaKlasat.CellForeColor = &HC0&

     
   End If

   If kaCikliUlet Then

      For i = nrU - 1 To 0 Step -1

         gridaKlasat.row = i
         gridaKlasat.col = 0
         kol0 = gridaKlasat.Text

         gridaKlasat.row = i
         gridaKlasat.col = 1
         kol1 = gridaKlasat.Text

         gridaKlasat.row = i
         gridaKlasat.col = 2
         kol2 = gridaKlasat.Text

        

         gridaKlasat.row = i
         gridaKlasat.col = 0
         gridaKlasat.Text = kol0

         gridaKlasat.row = i
         gridaKlasat.col = 1
         gridaKlasat.Text = kol1

         gridaKlasat.row = i
         gridaKlasat.col = 2
         gridaKlasat.Text = kol2

        
      Next i

      gridaKlasat.row = rrU - 1
      gridaKlasat.col = 0
      gridaKlasat.Text = "Cikli i ulet"
      gridaKlasat.CellForeColor = &HC0&
      gridaKlasat.col = 1
      gridaKlasat.Text = str(nrMeshkujU)
      gridaKlasat.CellForeColor = &HC0&

      gridaKlasat.col = 2
      gridaKlasat.Text = str(nrFemraU)
      gridaKlasat.CellForeColor = &HC0&

     
   End If

   gridaKlasat.row = gridaKlasat.Rows - 1
   gridaKlasat.col = 0
   gridaKlasat.Text = "Totali"
   gridaKlasat.CellForeColor = &H8000&
   gridaKlasat.col = 1
   gridaKlasat.Text = str(nrMeshkuj)
   gridaKlasat.CellForeColor = &H8000&
   gridaKlasat.col = 2
   gridaKlasat.Text = str(nrFemra)
   gridaKlasat.CellForeColor = &H8000&
 

   For i = 0 To gridaKlasat.Rows - 1
      gridaKlasat.RowHeight(i) = 270
      For j = 0 To gridaKlasat.Cols - 1
         gridaKlasat.row = i
         gridaKlasat.col = j
         gridaKlasat.CellAlignment = 1
      Next j
   Next i
End Sub

Private Sub shtoStatistikat1()
   Dim i                As Integer, j As Integer
   Dim Klasa            As String, indeksi As String, rreshti As String
   Dim kaCikliUlet      As Boolean
   Dim kaCikliTetevjecare As Boolean
   Dim kaCikliMesem     As Boolean
   Dim nrRReshtash      As Integer
   Dim lartesia         As Integer
   Dim nrU              As Integer
   Dim nrT              As Integer
   Dim nrM              As Integer
   Dim rrU              As Integer
   Dim rrT              As Integer
   Dim rrM              As Integer
   Dim nrMeshkujM       As Integer, nrFemraM As Integer
   Dim nrMeshkujT       As Integer, nrFemraT As Integer
   Dim nrMeshkujU       As Integer, nrFemraU As Integer
   Dim nrMeshkuj        As Integer, nrFemra As Integer
   Dim kol0             As String, kol1 As String, kol2 As String, kol3 As String
   Dim nrSmes           As Integer
   Dim nrSu             As Integer
   Dim nrSt             As Integer
   Dim nrs              As Integer
   Dim iM               As Integer
   Dim iT               As Integer
   nrsm = 0
   nrSt = 0
   nrSu = 0
   nrs = 0
   lartesia = gridaKlasat.Rows

   If lartesia = 0 Then
      Exit Sub
   End If
   nrU = 0
   nrM = 0
   nrT = 0
   For i = 0 To lartesia - 1
      gridaKlasat.row = i
      gridaKlasat.col = j
      rreshti = gridaKlasat.Text
      Klasa = ktheKlasenString(rreshti)
      If Klasa = "1" Or Klasa = "2" Or Klasa = "3" Or Klasa = "4" Or Klasa = "5" Then
         kaCikliUlet = True
         nrU = nrU + 1
         gridaKlasat.col = 1
         nrMeshkujU = nrMeshkujU + CInt(gridaKlasat.Text)
         nrMeshkuj = nrMeshkuj + CInt(gridaKlasat.Text)
         gridaKlasat.col = 2
         nrFemraU = nrFemraU + CInt(gridaKlasat.Text)
         nrFemra = nrFemra + CInt(gridaKlasat.Text)
      End If
      If Klasa = "6" Or Klasa = "7" Or Klasa = "8" Or Klasa = "9" Then

         kaCikliTetevjecare = True
         nrT = nrT + 1

         

         gridaKlasat.col = 1
         nrMeshkujT = nrMeshkujT + CInt(gridaKlasat.Text)
         nrMeshkuj = nrMeshkuj + CInt(gridaKlasat.Text)
         gridaKlasat.col = 2
         nrFemraT = nrFemraT + CInt(gridaKlasat.Text)
         nrFemra = nrFemra + CInt(gridaKlasat.Text)
      End If
      If Klasa = "10" Or Klasa = "11" Or Klasa = "12" Then

         kaCikliMesem = True
         nrM = nrM + 1

         gridaKlasat.col = 1
         nrMeshkujM = nrMeshkujM + CInt(gridaKlasat.Text)
         nrMeshkuj = nrMeshkuj + CInt(gridaKlasat.Text)
         gridaKlasat.col = 2
         nrFemraM = nrFemraM + CInt(gridaKlasat.Text)
         nrFemra = nrFemra + CInt(gridaKlasat.Text)
      End If

   Next i

   rrU = 0
   rrT = 0
   rrM = 0
   iT = 0
   iM = 0
   If kaCikliUlet Then
      gridaKlasat.Rows = gridaKlasat.Rows + 1
      
      rrU = nrU + 1
      rrT = rrU
      rrM = rrU
      iT = iT + 1
      iM = iM + 1
   End If

   If kaCikliTetevjecare Then
      gridaKlasat.Rows = gridaKlasat.Rows + 1
      
      rrT = rrT + nrT + 1
      rrM = rrM + nrT + 1
      iM = iM + 1
   End If

   If kaCikliMesem Then
      gridaKlasat.Rows = gridaKlasat.Rows + 1
      
      rrM = rrM + nrM + 1
   End If
   gridaKlasat.Rows = gridaKlasat.Rows + 1

   If kaCikliMesem Then

      For i = nrU + nrM + nrT - 1 To nrU + nrT Step -1

         gridaKlasat.row = i
         gridaKlasat.col = 0
         kol0 = gridaKlasat.Text

         gridaKlasat.row = i
         gridaKlasat.col = 1
         kol1 = gridaKlasat.Text

         gridaKlasat.row = i
         gridaKlasat.col = 2
         kol2 = gridaKlasat.Text

         

         gridaKlasat.row = i + iM
         gridaKlasat.col = 0
         gridaKlasat.Text = kol0

         gridaKlasat.row = i + iM
         gridaKlasat.col = 1
         gridaKlasat.Text = kol1

         gridaKlasat.row = i + iM
         gridaKlasat.col = 2
         gridaKlasat.Text = kol2

         
      Next i

      gridaKlasat.row = rrM - 1

      gridaKlasat.col = 0
      gridaKlasat.Text = "E mesmja"
      gridaKlasat.CellForeColor = &HC0&

      gridaKlasat.col = 1
      gridaKlasat.Text = str(nrMeshkujM)
      gridaKlasat.CellForeColor = &HC0&

      gridaKlasat.col = 2
      gridaKlasat.Text = str(nrFemraM)
      gridaKlasat.CellForeColor = &HC0&

     

   End If

   If kaCikliTetevjecare Then

      For i = nrU + nrT - 1 To nrU Step -1

         gridaKlasat.row = i
         gridaKlasat.col = 0
         kol0 = gridaKlasat.Text

         gridaKlasat.row = i
         gridaKlasat.col = 1
         kol1 = gridaKlasat.Text

         gridaKlasat.row = i
         gridaKlasat.col = 2
         kol2 = gridaKlasat.Text

         

         gridaKlasat.row = i + iT
         gridaKlasat.col = 0
         gridaKlasat.Text = kol0

         gridaKlasat.row = i + iT
         gridaKlasat.col = 1
         gridaKlasat.Text = kol1

         gridaKlasat.row = i + iT
         gridaKlasat.col = 2
         gridaKlasat.Text = kol2

         

      Next i

      gridaKlasat.row = rrT - 1
      gridaKlasat.col = 0
      gridaKlasat.Text = "Nëntëvjeçarja"
      gridaKlasat.CellForeColor = &HC0&
      gridaKlasat.col = 1
      gridaKlasat.Text = str(nrMeshkujT)
      gridaKlasat.CellForeColor = &HC0&

      gridaKlasat.col = 2
      gridaKlasat.Text = str(nrFemraT)
      gridaKlasat.CellForeColor = &HC0&

     
   End If

   If kaCikliUlet Then

      For i = nrU - 1 To 0 Step -1

         gridaKlasat.row = i
         gridaKlasat.col = 0
         kol0 = gridaKlasat.Text

         gridaKlasat.row = i
         gridaKlasat.col = 1
         kol1 = gridaKlasat.Text

         gridaKlasat.row = i
         gridaKlasat.col = 2
         kol2 = gridaKlasat.Text

        

         gridaKlasat.row = i
         gridaKlasat.col = 0
         gridaKlasat.Text = kol0

         gridaKlasat.row = i
         gridaKlasat.col = 1
         gridaKlasat.Text = kol1

         gridaKlasat.row = i
         gridaKlasat.col = 2
         gridaKlasat.Text = kol2

        
      Next i

      gridaKlasat.row = rrU - 1
      gridaKlasat.col = 0
      gridaKlasat.Text = "Cikli i ulet"
      gridaKlasat.CellForeColor = &HC0&
      gridaKlasat.col = 1
      gridaKlasat.Text = str(nrMeshkujU)
      gridaKlasat.CellForeColor = &HC0&

      gridaKlasat.col = 2
      gridaKlasat.Text = str(nrFemraU)
      gridaKlasat.CellForeColor = &HC0&

     
   End If

   gridaKlasat.row = gridaKlasat.Rows - 1
   gridaKlasat.col = 0
   gridaKlasat.Text = "Totali"
   gridaKlasat.CellForeColor = &H8000&
   gridaKlasat.col = 1
   gridaKlasat.Text = str(nrMeshkuj)
   gridaKlasat.CellForeColor = &H8000&
   gridaKlasat.col = 2
   gridaKlasat.Text = str(nrFemra)
   gridaKlasat.CellForeColor = &H8000&
 

   For i = 0 To gridaKlasat.Rows - 1
      gridaKlasat.RowHeight(i) = 270
      For j = 0 To gridaKlasat.Cols - 1
         gridaKlasat.row = i
         gridaKlasat.col = j
         gridaKlasat.CellAlignment = 1
      Next j
   Next i
End Sub
Private Function ktheKlasenString(rreshti As String) As String
   Dim Klasa            As String
   s3 = Trim(rreshti)
   Dim s2               As String
   s2 = ""
   Dim ugjet            As Boolean
   Dim i, j             As Integer
   ugjet = False
   i = 1
   s2 = s3

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

   ktheKlasenString = Klasa
End Function



' Ben printimin e gjeneraliteteve te pergjithshme te shkolles
Private Sub PrintoFormatPergj()
    objPrintClass.PrintLeft 5, 17, lblShkolla.Caption, True, 12
    objPrintClass.PrintLeft 130, 17, "Adresa:", True, 12
    objPrintClass.PrintLeft 130, 23, adresaShkolla, True, 12
    objPrintClass.PrintLeft 20, 23, qytetiShkolla, True, 12
    objPrintClass.PrintLeft 130, 29, "Tel: " & telefoniShkolla, True, 12
    objPrintClass.PrintLeft 130, 35, "website: " & website, True, 12
    objPrintClass.PrintLeft 130, 41, "Email: " & email, True, 12
    If qytetiShkolla <> "" Then
        objPrintClass.PrintLeft 11, 261, qytetiShkolla & " me " & FormatDateTime(Now, vbShortDate), True, 12
    End If
    objPrintClass.PrintLeft 131, 261, "Drejtori  ___________________", True, 12
    If Image1 <> 0 Then
        objPrintClass.PrintPicture 81, 3, Image1, 35, 35
    End If
End Sub

