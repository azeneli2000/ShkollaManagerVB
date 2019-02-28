VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmStatistikaKlasa 
   Caption         =   "Statistika - Klasa"
   ClientHeight    =   10365
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
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
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10365
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtKlasaZgjedhur 
      Height          =   375
      Left            =   7200
      TabIndex        =   25
      Top             =   2880
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridaNr 
      Height          =   2535
      Left            =   12480
      TabIndex        =   22
      Top             =   1680
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   4471
      _Version        =   393216
      FixedRows       =   0
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridaLendetLabel 
      Height          =   375
      Left            =   7560
      TabIndex        =   21
      Top             =   1560
      Width           =   4575
      _ExtentX        =   8070
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridaLabelKlasat 
      Height          =   300
      Left            =   360
      TabIndex        =   20
      Top             =   1560
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   529
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H80000009&
      DisabledPicture =   "frmStatistikaKlasa.frx":0000
      DownPicture     =   "frmStatistikaKlasa.frx":353A
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
      Left            =   3960
      Picture         =   "frmStatistikaKlasa.frx":6A74
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   9120
      Width           =   2535
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
      Height          =   1455
      Left            =   12360
      TabIndex        =   12
      Top             =   0
      Width           =   2655
      Begin VB.CommandButton cmdWebsiste 
         BackColor       =   &H80000009&
         Caption         =   "Website"
         Height          =   375
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton cmdEmail 
         BackColor       =   &H80000009&
         Caption         =   "E-mail"
         Height          =   375
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   960
         Width           =   1095
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   120
         Picture         =   "frmStatistikaKlasa.frx":9FAE
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
         TabIndex        =   16
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
         TabIndex        =   15
         Top             =   120
         Width           =   2415
      End
   End
   Begin VB.ComboBox cboOptions 
      BackColor       =   &H80000003&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   360
      Left            =   960
      TabIndex        =   7
      Top             =   840
      Width           =   2535
   End
   Begin VB.CommandButton cmdNdertoGraf1 
      BackColor       =   &H80000009&
      Caption         =   "Nderto "
      Height          =   375
      Left            =   14280
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Ndertoni grafikun e mesatareve per lendet"
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton cmdNdertoGraf2 
      BackColor       =   &H80000009&
      Caption         =   "Nderto "
      Height          =   375
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Ndertoni grafikun per mesataren e cdo klase"
      Top             =   4680
      Width           =   975
   End
   Begin VB.CommandButton cmdDil 
      BackColor       =   &H80000009&
      DisabledPicture =   "frmStatistikaKlasa.frx":10F23
      DownPicture     =   "frmStatistikaKlasa.frx":17365
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
      Left            =   7680
      Picture         =   "frmStatistikaKlasa.frx":1D7A7
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9120
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000009&
      DisabledPicture =   "frmStatistikaKlasa.frx":23BE9
      DownPicture     =   "frmStatistikaKlasa.frx":27123
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
      Picture         =   "frmStatistikaKlasa.frx":2A65D
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9120
      Width           =   2535
   End
   Begin VB.ComboBox cboVitiShkollor 
      BackColor       =   &H80000003&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   360
      Left            =   3840
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   840
      Width           =   2535
   End
   Begin VB.CommandButton cmdShfaq 
      BackColor       =   &H80000009&
      DisabledPicture =   "frmStatistikaKlasa.frx":2DB97
      DownPicture     =   "frmStatistikaKlasa.frx":33FD9
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
      Left            =   7080
      Picture         =   "frmStatistikaKlasa.frx":3A41B
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   800
      Width           =   2415
   End
   Begin MSChart20Lib.MSChart grafLende 
      Height          =   4095
      Index           =   1
      Left            =   7200
      OleObjectBlob   =   "frmStatistikaKlasa.frx":4085D
      TabIndex        =   8
      Top             =   4680
      Width           =   7935
   End
   Begin VB.Frame fraZgjedhja 
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
      Left            =   360
      TabIndex        =   9
      Top             =   120
      Width           =   11775
      Begin VB.Label Label2 
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
         Left            =   3480
         TabIndex        =   11
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Semestri :"
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
         Left            =   600
         TabIndex        =   10
         Top             =   480
         Width           =   2535
      End
   End
   Begin MSChart20Lib.MSChart grafKlase 
      Height          =   4095
      Index           =   0
      Left            =   360
      OleObjectBlob   =   "frmStatistikaKlasa.frx":425FD
      TabIndex        =   6
      Top             =   4680
      Width           =   6855
   End
   Begin VB.PictureBox Picture1 
      Height          =   4095
      Left            =   0
      ScaleHeight     =   4035
      ScaleWidth      =   5955
      TabIndex        =   23
      Top             =   6600
      Visible         =   0   'False
      Width           =   6015
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   4095
      Left            =   120
      TabIndex        =   24
      Top             =   4440
      Width           =   12615
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridaKlasat 
      Height          =   2295
      Left            =   360
      TabIndex        =   18
      Top             =   1920
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   4048
      _Version        =   393216
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridaLendet 
      Height          =   2295
      Left            =   7560
      TabIndex        =   19
      Top             =   1920
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   4048
      _Version        =   393216
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frmStatistikaKlasa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private vektor1(30) As String
Private vektor2(30) As String
Private vektor3(30) As String
Dim objektGabimi As New clsErrorHandler
Dim objUIController As clsUIController
Dim objPrintForm As New frmPrintStatistika
Dim objPrintClass As Object
Dim PrintStat As Integer
Dim Klasa As String
Dim Adresa As String
Dim Qyteti As String
Dim Telefoni As String


Private Sub cmdDil_Click()
    Unload Me
    Set active_form = Nothing
End Sub

Private Sub cmdEmail_Click()
    If Not OpenEmailProgram(email) Then
        MsgBox "Ju nuk keni asnje program te instaluar per te derguar email", vbExclamation, "Gabim ne dergimin e emailit"
    End If
End Sub

Private Sub cmdNdertoGraf1_Click()
    Dim mesatarja As String
    Dim lenda As String
    Dim mes
    Dim j As Integer
    If (gridaLendet.Rows < 2) Then
        Exit Sub
    End If
    gridaLendet.row = 1
    gridaLendet.col = 0
    If gridaLendet.Text = "" Then
        MsgBox "Deri tani nuk eshte dhene asnje lende dhe note per te ndertuar grafikun", vbExclamation + vbOKOnly, "Ndertimi i grafikut"
    Else
        grafLende(1).Visible = True
        grafLende(1).rowCount = gridaLendet.Rows
    
        ' Rregullon shkallen e grafit
        With grafLende(1).Plot.Axis(VtChAxisIdY).ValueScale
            .Auto = False
            .Minimum = 0
            .Maximum = 10
            .MajorDivision = 10
            .MinorDivision = 1
        End With

    
        Dim i As Integer
        For i = 0 To gridaLendet.Rows - 1
          'lenda = Trim(lista2.List(i - 1))
          gridaLendet.col = 0
          gridaLendet.row = i
          lenda = gridaLendet.Text
          grafLende(1).row = i + 1
          grafLende(1).RowLabel = lenda
      
          'grafLende.data = Val(lista3.List(i - 1))
          gridaLendet.col = 1
          gridaLendet.row = i
          mesatarja = gridaLendet.Text
          grafLende(1).data = Val(mesatarja)
        Next i
    End If
End Sub

Private Sub cmdNdertoGraf2_Click()
    Dim mesatarja As String
    Dim Kl As String
    
    
    If (gridaKlasat.Rows < 2) Then
        Exit Sub
    End If

    gridaKlasat.row = 1
    gridaKlasat.col = 0
    
    If gridaKlasat.Text = "" Then
        MsgBox "Deri tani nuk eshte dhene asnje lende dhe note per te ndertuar grafikun", vbOKOnly, "Ndertimi i grafikut"
    Else
        grafKlase(0).rowCount = gridaKlasat.Rows
        grafKlase(0).Visible = True
        ' Rregullon shkallen e grafit
        With grafKlase(0).Plot.Axis(VtChAxisIdY).ValueScale
            .Auto = False
            .Minimum = 0
            .Maximum = 10
            .MajorDivision = 10
            .MinorDivision = 1
        End With

        Dim i As Integer
        For i = 0 To gridaKlasat.Rows - 1
          'lenda = Trim(lista2.List(i - 1))
          gridaKlasat.col = 0
          gridaKlasat.row = i
          Kl = gridaKlasat.Text
          grafKlase(0).row = i + 1
          
          grafKlase(0).RowLabel = Kl
          'grafLende.data = Val(lista3.List(i - 1))
          gridaKlasat.col = 3
          gridaKlasat.row = i
          mesatarja = gridaKlasat.Text
          grafKlase(0).data = Val(mesatarja)
        Next i
    End If
End Sub

Private Sub cmdOK_Click()
    lista1.Clear
    objectInitialization STATISTIKA_KLASA_SHFAQ
    ciklet
    lista1.Clear
  If Check1.Value = 1 Then
   '     mbushListen vektor1, 30, lista1
    End If
   If Check2.Value = 1 Then
   '     mbushListen vektor2, 30, lista1
    End If
    If Check3.Value = 1 Then
  '      mbushListen vektor3, 30, lista1
   End If
  '  renditListe lista1
End Sub

Private Sub cmdPrint_Click()
    objPrintForm.show vbModal
    PrintStat = objPrintForm.choice
    Unload objPrintForm
    Set objPrintClass = CreateObject("PrintimComponent.clsPrintim")
    If objPrintClass.PrinterIsInstalled = False Then
        Exit Sub
    End If
    Select Case PrintStat
        Case 1
            PrintStatistika
        Case 2
            PrintMesLendet
        Case 3
            PrintGrafKlase
        Case 4
            PrintGrafLende
    End Select
End Sub

Private Sub cmdShfaq_Click()
   
   gridaKlasat.Visible = False
   objektGabimi.kapGabimin
   objektGabimi.menazhim_gabimi
   grafKlase(0).rowCount = 0
   grafLende(1).rowCount = 0
   If objektGabimi.mvarGabimi = 0 Then
      
      pastroGriden gridaKlasat
      pastroGriden gridaLendet
      inicializoGridaLendet gridaLendet, gridaLendetLabel, 9
      ngjyrosGriden gridaLendet
      objectInitialization STATISTIKA_KLASA_SHFAQ
      
      renditGriden gridaKlasat
      Dim vitFillimi As Integer
      vitFillimi = CInt(Mid(Me.cboVitiShkollor.Text, 1, 4))
      If (vitFillimi <= 2007) Then
        shtoStatistikat
      Else
        shtoStatistikat1
      End If
      formatoGrida gridaKlasat, 2430
      ngjyrosGriden gridaKlasat
      cmdPrint.Enabled = True
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



Private Sub Form_Load()
   loadForm Me
   mbushKomboBox
   viti
   'txtVitiShkollor.Text = cboVitiShkollor.Text
   Me.grafKlase(0).rowCount = 0
   Me.grafLende(1).rowCount = 0
   Image1.Picture = LoadPicture(adresaLogo)
   lblShkolla.Caption = emerShkolla
   Adresa = adresaShkolla
   Qyteti = qytetiShkolla
   Telefoni = telefoniShkolla
   cmdPrint.Enabled = False
   'inicializoGridat
   inicializoGridaKlasat gridaKlasat, gridaLabelKlasat, 9
   inicializoGridaLendet gridaLendet, gridaLendetLabel, 9
   loadComboItems
   ngjyrosGriden gridaKlasat
   ngjyrosGriden gridaLendet
 
   
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


Private Sub gridaKlasat_Click()
   Dim i                As Integer
   'Dim Klasa            As String
   i = gridaKlasat.RowSel
   gridaKlasat.row = i
   gridaKlasat.col = 0
   ' gridaLendet.Height = 2700
   Klasa = gridaKlasat.Text
   Me.txtKlasaZgjedhur.Text = Klasa
   If Klasa <> "" Then
      If Not (Klasa = "Totali" Or Klasa = "Cikli i ulet" Or Klasa = "E mesmja" Or Klasa = "Tetevjecarja" Or Klasa = "Nëntëvjeçarja") Then
         gridaLendet.Visible = False
         pastroGriden gridaLendet
         grafLende(1).rowCount = 0
         objectInitialization STATISTIKA_LENDET_MESATARET
         formatoGrida gridaLendet, 2430
         ngjyrosGriden gridaLendet
         gridaLendet.Visible = True
         gridaLendetLabel.Visible = True
      Else
         gridaLendet.Visible = False
         gridaLendetLabel.Visible = False
      End If
   End If

End Sub

Private Sub lista1_Click()
   lista2.Clear
   lista3.Clear
   objectInitialization STATISTIKA_LENDET_MESATARET
End Sub


Private Sub ciklet()
    Dim k1 As Integer
    Dim k2 As Integer
    Dim k3 As Integer
    k1 = 1
    k2 = 1
    k3 = 1
    
    Dim i As Integer
    Dim rreshti As String
    For i = 0 To lista1.ListCount - 1
        rreshti = lista1.List(i)
        If ktheKlasenString(rreshti) = "1" Or ktheKlasenString(rreshti) = "2" Or ktheKlasenString(rreshti) = "3" Or ktheKlasenString(rreshti) = "4" Then
            vektor1(k1) = rreshti
            k1 = k1 + 1
        End If
        If ktheKlasenString(rreshti) = "5" Or ktheKlasenString(rreshti) = "6" Or ktheKlasenString(rreshti) = "7" Or ktheKlasenString(rreshti) = "8" Then
            vektor2(k2) = rreshti
            k2 = k2 + 1
        End If
        If ktheKlasenString(rreshti) = "9" Or ktheKlasenString(rreshti) = "10" Or ktheKlasenString(rreshti) = "11" Or ktheKlasenString(rreshti) = "12" Then
            vektor3(k3) = rreshti
            k3 = k3 + 1
        End If
    Next i
End Sub

Private Sub mbushListen(vektori() As String, M As Integer, lista As ListBox)
    
   Dim i As Integer
   For i = 1 To M
    If vektori(i) <> "" Then
        lista.AddItem vektori(i)
    End If
    i = i + 1
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

Private Sub inicializoGridat()

   Dim i                As Integer
   Dim j                As Integer
   gridaKlasat.Cols = 4
   gridaKlasat.FixedCols = 0
   gridaKlasat.row = 0
   gridaKlasat.col = 0
   gridaKlasat.ColWidth(0) = 1000
   gridaKlasat.ColWidth(1) = 2000
   gridaKlasat.ColWidth(2) = 2000
   gridaKlasat.ColWidth(3) = 1590
   gridaKlasat.CellFontSize = 10
   gridaKlasat.CellAlignment = vbCenter
   gridaKlasat.Text = "Klasa"
   gridaKlasat.col = 1

   gridaKlasat.CellFontSize = 10
   gridaKlasat.CellAlignment = vbCenter
   gridaKlasat.Text = "Mungesat(pa arsye)"
   gridaKlasat.col = 2

   gridaKlasat.CellFontSize = 10
   gridaKlasat.CellAlignment = vbCenter
   gridaKlasat.Text = "Mungesat(me arsye)"
   gridaKlasat.col = 3

   gridaKlasat.CellFontSize = 10
   gridaKlasat.CellAlignment = vbCenter
   gridaKlasat.Text = "Mesatarja"

   gridaLendet.Cols = 3
   gridaLendet.FixedCols = 0
   gridaLendet.row = 0
   gridaLendet.col = 0
   gridaLendet.CellFontSize = 10
   gridaLendet.CellAlignment = vbCenter
   gridaLendet.Text = "Lenda"
   gridaLendet.col = 1
   gridaLendet.CellFontSize = 10
   gridaLendet.CellAlignment = vbCenter
   gridaLendet.Text = "Mesatarja"
   gridaLendet.ColWidth(0) = 2300
   gridaLendet.ColWidth(1) = 2300
   gridaLendet.ScrollBars = flexScrollBarNone
   gridaKlasat.ScrollBars = flexScrollBarNone
   gridaKlasat.Rows = 12
   gridaLendet.Rows = 12

   For i = 0 To gridaKlasat.Rows - 1
      For j = 0 To gridaKlasat.Cols - 1
         gridaKlasat.row = i
         gridaKlasat.col = j
         gridaKlasat.CellAlignment = vbCenter
      Next j
   Next i
   
   For i = 0 To gridaLendet.Rows - 1
      For j = 0 To gridaLendet.Cols - 1
         gridaLendet.row = i
         gridaLendet.col = j
         gridaLendet.CellAlignment = vbCenter
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
   Dim mungesa_pa_arsyeI As String
   Dim mungesa_me_arsyeI As String
   Dim mesatarjaI       As String
   Dim mungesa_pa_arsyeJ As String
   Dim mungesa_me_arsyeJ As String
   Dim mesatarjaJ       As String
   Dim ruajMesataren    As String
   Dim ruajMungesa_pa_arsye As String
   Dim ruajMungesa_me_arsye As String
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
            mungesa_pa_arsyeI = gridaKlasat.Text
            grida.row = j
            mungesa_pa_arsyeJ = gridaKlasat.Text
            gridaKlasat.row = i
            gridaKlasat.Text = mungesa_pa_arsyeJ
            grida.row = j
            gridaKlasat.Text = mungesa_pa_arsyeI

            gridaKlasat.col = 2
            gridaKlasat.row = i
            mungesa_me_arsyeI = gridaKlasat.Text
            grida.row = j
            mungesa_me_arsyeJ = gridaKlasat.Text
            gridaKlasat.row = i
            gridaKlasat.Text = mungesa_me_arsyeJ
            grida.row = j
            gridaKlasat.Text = mungesa_me_arsyeI

            gridaKlasat.col = 3
            gridaKlasat.row = i
            mesatarjaI = gridaKlasat.Text
            grida.row = j
            mesatarjaJ = gridaKlasat.Text
            gridaKlasat.row = i
            gridaKlasat.Text = mesatarjaJ
            grida.row = j
            gridaKlasat.Text = mesatarjaI

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
               mungesa_pa_arsyeI = gridaKlasat.Text
               grida.row = j
               mungesa_pa_arsyeJ = gridaKlasat.Text
               gridaKlasat.row = i
               gridaKlasat.Text = mungesa_pa_arsyeJ
               grida.row = j
               gridaKlasat.Text = mungesa_pa_arsyeI

               gridaKlasat.col = 2
               gridaKlasat.row = i
               mungesa_me_arsyeI = gridaKlasat.Text
               grida.row = j
               mungesa_me_arsyeJ = gridaKlasat.Text
               gridaKlasat.row = i
               gridaKlasat.Text = mungesa_me_arsyeJ
               grida.row = j
               gridaKlasat.Text = mungesa_me_arsyeI

               gridaKlasat.col = 3
               gridaKlasat.row = i
               mesatarjaI = gridaKlasat.Text
               grida.row = j
               mesatarjaJ = gridaKlasat.Text
               gridaKlasat.row = i
               gridaKlasat.Text = mesatarjaJ
               grida.row = j
               gridaKlasat.Text = mesatarjaI
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
   If l <= lartesia Then
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

' Printon statistikat per klasat duke u nisur nga grida e statistikave
Private Sub PrintStatistika()
   Dim i, j, p          As Integer
   Dim tekst            As String
   Dim maxRows          As Integer
    Dim currPage As Integer, totalPages As Integer
    maxRows = 0
    If gridaKlasat.Rows Mod 27 = 0 Then
        totalPages = gridaKlasat.Rows / 27
    Else
        totalPages = Fix(gridaKlasat.Rows / 27) + 1
    End If
    currPage = 1
   p = 0
   Do While p <= gridaKlasat.Rows - 1
      objPrintClass.PrintFont "Times New Roman"
      objPrintClass.OrientimFaqe True
      objPrintClass.PrintLeft 55, 65, "Statistikat e klasave për vitin shkollor ", True, 12
      If objPrintClass.Gabim = True Then
        Exit Sub
      End If
      objPrintClass.PrintLeft 130, 65, cboVitiShkollor.Text, True, 12, False, True
      If gridaKlasat.Rows - 1 - p > 26 Then
         maxRows = 26
      Else
         maxRows = gridaKlasat.Rows - p
      End If
      ' Printohen te dhenat e shkolles qe sherbejne ne cdo printim
      PrintoFormatPergj
      objPrintClass.PrintTabele 30, 75, 40, 14, 1, 1
      objPrintClass.PrintTabele 70, 75, 50, 7, 1, 1
      objPrintClass.PrintTabele 70, 82, 25, 7, 1, 2
      objPrintClass.PrintTabele 120, 75, 40, 14, 1, 1
      objPrintClass.PrintTabele 30, 89, 40, 6, maxRows, 1
      objPrintClass.PrintTabele 70, 89, 25, 6, maxRows, 2
      objPrintClass.PrintTabele 120, 89, 40, 6, maxRows, 1

      objPrintClass.PrintLeft 40, 80, "Klasa", True, 12, False
      objPrintClass.PrintLeft 86, 76.5, "Mungesa", True, 12, False
      objPrintClass.PrintLeft 72, 83.5, "pa arsye", True, 12, False
      objPrintClass.PrintLeft 97, 83.5, "me arsye", True, 12, False
      objPrintClass.PrintLeft 126, 80, "Mesatarja", True, 12, False
      j = 0
      For i = p To maxRows - 1 + p
         gridaKlasat.col = 0
         gridaKlasat.row = i
         tekst = gridaKlasat.Text
         objPrintClass.PrintLeft 34, 90 + 6 * j, tekst, False, 12, False
         gridaKlasat.col = 1
         gridaKlasat.row = i
         tekst = gridaKlasat.Text
         objPrintClass.PrintLeft 80, 90 + 6 * j, tekst, False, 12, False
         gridaKlasat.col = 2
         gridaKlasat.row = i
         tekst = gridaKlasat.Text
         objPrintClass.PrintLeft 105, 90 + 6 * j, tekst, False, 12, False
         gridaKlasat.col = 3
         gridaKlasat.row = i
         tekst = gridaKlasat.Text
         objPrintClass.PrintLeft 125, 90 + 6 * j, tekst, False, 12, False
         j = j + 1
      Next
      p = p + maxRows
      If p <= gridaKlasat.Rows - 1 Then
         'maxRows = maxRows + 26
         objPrintClass.PrintLeft 191, 275, CStr(currPage) & "/" & CStr(totalPages)
         currPage = currPage + 1
         objPrintClass.NewPage
      End If
   Loop
   objPrintClass.PrintLeft 191, 275, CStr(currPage) & "/" & CStr(totalPages)
   objPrintClass.EndDoc
End Sub

Private Sub PrintMesLendet()
    Dim i, j As Integer
    Dim tekst As String
    gridaLendet.row = 1
    gridaLendet.col = 1
    If gridaLendet.Text <> "" Then
        objPrintClass.PrintFont "Times New Roman"
        objPrintClass.OrientimFaqe True
        objPrintClass.PrintLeft 57, 70, "Klasa:", True, 12, False
        ' Nese perdoruesi anullon printimin
        If objPrintClass.Gabim = True Then
            Exit Sub
        End If
        objPrintClass.PrintLeft 74, 70, Klasa, True, 12, False, True
        objPrintClass.PrintLeft 107, 70, "Viti shkollor:", True, 12, False, False
        objPrintClass.PrintLeft 134, 70, cboVitiShkollor.Text, True, 12, False, True
        PrintoFormatPergj
        objPrintClass.PrintTabele 51, 80, 11, 6, 21, 1
        objPrintClass.PrintTabele 62, 80, 57, 6, 21, 1
        objPrintClass.PrintTabele 119, 80, 40, 6, 21, 1
        
        objPrintClass.PrintLeft 53, 81, "Nr", True, 12, False
        objPrintClass.PrintLeft 66, 81, "Lenda", True, 12, False
        objPrintClass.PrintLeft 126, 81, "Mesatarja", True, 12, False
        
        For i = 1 To 20
            objPrintClass.PrintLeft 53, 87 + 6 * (i - 1), CStr(i), False, 12, False
        Next
        
        For i = 0 To gridaLendet.Rows - 1
            gridaLendet.col = 0
            gridaLendet.row = i
            tekst = gridaLendet.Text
            objPrintClass.PrintLeft 63, 87 + 6 * i, tekst, False, 12, False
            gridaLendet.col = 1
            gridaLendet.row = i
            tekst = gridaLendet.Text
            objPrintClass.PrintLeft 120, 87 + 6 * i, tekst, False, 12, False
           
        Next
        objPrintClass.EndDoc
    Else
        MsgBox "Ju nuk keni perzgjedhur asnje klase per te pare mesataret e lendeve", vbOKOnly + vbInformation, "Zgjidhni klasen"
    End If
End Sub


Private Sub PrintGrafKlase()
    If grafKlase(0).rowCount <> 0 Then
        ' Meqe eshte grafik printimi do behet ne Landscape dhe jo ne Portrait
        On Error GoTo fund
        objPrintClass.PrintFont "Times New Roman"
        objPrintClass.OrientimFaqe True
        objPrintClass.PrintLeft 0, 0, ""
        If objPrintClass.Gabim = True Then
            Exit Sub
        End If
        PrintoFormatPergj
        grafKlase(0).Plot.DataSeriesInRow = False
        'With grafKlase(0).Plot.SeriesCollection(1).DataPoints(-1)
        '    .Brush.Style = VtBrushStyleSolid
        '    .Brush.FillColor.Set 0, 30, 30
        'End With
        grafKlase(0).Width = 9000
        grafKlase(0).Height = 5500

        'grafKlase(0).Plot.Backdrop.Fill.Brush.FillColor.Set 10, 10, 10
        grafKlase(0).EditCopy
        Picture1.Picture = Clipboard.GetData(Picture)
        'With grafKlase(0).Plot.SeriesCollection(1).DataPoints(-1)
        '    .Brush.Style = VtBrushStyleSolid
        '    .Brush.FillColor.Set 0, 30, 30
        'End With
        objPrintClass.PrintPicture 40, 65, Picture1 ', 130, 50
        'grafKlase(0).Plot.DataSeriesInRow = True
        'objPrintClass.PrinterColor = "vbPRCMColor"
        objPrintClass.EndDoc
        grafKlase(0).Width = 6855
        grafKlase(0).Height = 4095
    Else
        MsgBox "Ju nuk keni asgje te paraqitur ne grafikun per mesataret e klasave", vbOKOnly + vbInformation, "Ndertoni me pare grafikun"
    End If
fund:
End Sub


Private Sub PrintGrafLende()
    If grafLende(1).rowCount <> 0 Then
        ' Meqe eshte grafik printimi do behet ne Landscape dhe jo ne Portrait
        On Error GoTo fund
        objPrintClass.OrientimFaqe True
        'objPrintClass.PrintLeft 40, 10, "Grafiku për mesataret e klasave është:", True, 12, False
        PrintoFormatPergj
        If objPrintClass.Gabim = True Then
            Exit Sub
        End If
        grafLende(1).Plot.DataSeriesInRow = False
        'With grafLende(1).Plot.SeriesCollection(1).DataPoints(-1)
        '    .Brush.Style = VtBrushStyleSolid
        '    .Brush.FillColor.Set 0, 30, 30
        'End With
        grafLende(1).EditCopy
        Picture1.Picture = Clipboard.GetData(Picture)
        objPrintClass.PrintPicture 40, 75, Picture1 ', 130, 50
        'grafLende(1).Plot.DataSeriesInRow = True
        'With grafLende(1).Plot.SeriesCollection(1).DataPoints(-1)
        '    .Brush.Style = VtBrushStyleSolid
        '    .Brush.FillColor.Set 255, 0, 0
        'End With
        objPrintClass.PrintLeft 80, 150, "Klasa: ", True, 12, False
        objPrintClass.PrintLeft 93, 150, Klasa, True, 12, False, True
        'objPrintClass.ColorMode = "vbPRCMColor"
        objPrintClass.EndDoc
    Else
        MsgBox "Ju nuk keni asgje te paraqitur ne grafikun per mesataren e lendeve", vbOKOnly + vbInformation, "Ndertoni me pare grafikun"
    End If
fund:
End Sub


Private Sub ngjyrosGriden(grida As MSHFlexGrid)

   Dim i                As Integer
   Dim j                As Integer
   For i = 1 To grida.Rows - 1
      If i Mod 2 = 1 Then
         For j = 0 To grida.Cols - 1
            grida.col = j
            grida.row = i
            grida.CellBackColor = &HE0E0E0
         Next j
      End If
   Next i
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
   Dim mungMeM          As Integer, mungPaM As Integer
   Dim mungMeT          As Integer, mungPaT As Integer
   Dim mungMeu          As Integer, mungPaU As Integer
   Dim mungMe           As Integer, mungPa As Integer
   Dim mesatarjaMesme   As Double
   Dim mesatarjaUlet    As Double
   Dim mesatarjaTetevjecare As Double
   Dim mesatarja        As Double
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
         gridaNr.col = 1
         gridaNr.row = i

         gridaKlasat.col = 3
         gridaKlasat.row = i
         If IsNumeric(gridaKlasat.Text) Then
            mesatarjaUlet = CInt(gridaNr.Text) * CDbl(gridaKlasat.Text) + mesatarjaUlet
            nrSu = nrSu + CInt(gridaNr.Text)
            mesatarja = mesatarja + CInt(gridaNr.Text) * CDbl(gridaKlasat.Text)
            nrs = nrs + CInt(gridaNr.Text)
         End If

         gridaKlasat.col = 1
         mungMeu = mungMeu + CInt(gridaKlasat.Text)
         mungMe = mungMe + CInt(gridaKlasat.Text)
         gridaKlasat.col = 2
         mungPaU = mungPaU + CInt(gridaKlasat.Text)
         mungPa = mungPa + CInt(gridaKlasat.Text)
      End If
      If Klasa = "5" Or Klasa = "6" Or Klasa = "7" Or Klasa = "8" Then

         kaCikliTetevjecare = True
         nrT = nrT + 1

         gridaNr.col = 1
         gridaNr.row = i

         gridaKlasat.col = 3
         gridaKlasat.row = i
         If IsNumeric(gridaKlasat.Text) Then
            mesatarjaTetevjecare = CInt(gridaNr.Text) * CDbl(gridaKlasat.Text) + mesatarjaTetevjecare
            nrSt = nrSt + CInt(gridaNr.Text)
            mesatarja = mesatarja + CInt(gridaNr.Text) * CDbl(gridaKlasat.Text)
            nrs = nrs + CInt(gridaNr.Text)
         End If

         gridaKlasat.col = 1
         mungMeT = mungMeT + CInt(gridaKlasat.Text)
         mungMe = mungMe + CInt(gridaKlasat.Text)
         gridaKlasat.col = 2
         mungPaT = mungPaT + CInt(gridaKlasat.Text)
         mungPa = mungPa + CInt(gridaKlasat.Text)
      End If
      If Klasa = "9" Or Klasa = "10" Or Klasa = "11" Or Klasa = "12" Then

         kaCikliMesem = True
         nrM = nrM + 1

         gridaNr.col = 1
         gridaNr.row = i

         gridaKlasat.col = 3
         gridaKlasat.row = i
         If IsNumeric(gridaKlasat.Text) Then
            mesatarjaMesme = CInt(gridaNr.Text) * CDbl(gridaKlasat.Text) + mesatarjaMesme
            nrsm = nrsm + CInt(gridaNr.Text)
            mesatarja = mesatarja + CInt(gridaNr.Text) * CDbl(gridaKlasat.Text)
            nrs = nrs + CInt(gridaNr.Text)
         End If

         gridaKlasat.col = 1
         mungMeM = mungMeM + CInt(gridaKlasat.Text)
         mungMe = mungMe + CInt(gridaKlasat.Text)
         gridaKlasat.col = 2
         mungPaM = mungPaM + CInt(gridaKlasat.Text)
         mungPa = mungPa + CInt(gridaKlasat.Text)
      End If

   Next i

   rrU = 0
   rrT = 0
   rrM = 0
   iT = 0
   iM = 0
   If kaCikliUlet Then
      gridaKlasat.Rows = gridaKlasat.Rows + 1
      gridaNr.Rows = gridaNr.Rows + 1
      rrU = nrU + 1
      rrT = rrU
      rrM = rrU
      iT = iT + 1
      iM = iM + 1
   End If

   If kaCikliTetevjecare Then
      gridaKlasat.Rows = gridaKlasat.Rows + 1
      gridaNr.Rows = gridaNr.Rows + 1
      rrT = rrT + nrT + 1
      rrM = rrM + nrT + 1
      iM = iM + 1
   End If

   If kaCikliMesem Then
      gridaKlasat.Rows = gridaKlasat.Rows + 1
      gridaNr.Rows = gridaNr.Rows + 1
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

         gridaKlasat.row = i
         gridaKlasat.col = 3
         kol3 = gridaKlasat.Text

         gridaKlasat.row = i + iM
         gridaKlasat.col = 0
         gridaKlasat.Text = kol0

         gridaKlasat.row = i + iM
         gridaKlasat.col = 1
         gridaKlasat.Text = kol1

         gridaKlasat.row = i + iM
         gridaKlasat.col = 2
         gridaKlasat.Text = kol2

         gridaKlasat.row = i + iM
         gridaKlasat.col = 3
         gridaKlasat.Text = kol3

      Next i

      gridaKlasat.row = rrM - 1

      gridaKlasat.col = 0
      gridaKlasat.Text = "E mesmja"
      gridaKlasat.CellForeColor = &HC0&

      gridaKlasat.col = 1
      gridaKlasat.Text = Space(10) & str(mungMeM)
      gridaKlasat.CellForeColor = &HC0&

      gridaKlasat.col = 2
      gridaKlasat.Text = Space(10) & str(mungPaM)
      gridaKlasat.CellForeColor = &HC0&

      gridaKlasat.col = 3
      If nrsm > 0 Then
         gridaKlasat.Text = Space(10) & Format(mesatarjaMesme / nrsm, ".0")
      Else
         gridaKlasat.Text = Space(10) & "---"
      End If
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

         gridaKlasat.row = i
         gridaKlasat.col = 3
         kol3 = gridaKlasat.Text

         gridaKlasat.row = i + iT
         gridaKlasat.col = 0
         gridaKlasat.Text = kol0

         gridaKlasat.row = i + iT
         gridaKlasat.col = 1
         gridaKlasat.Text = kol1

         gridaKlasat.row = i + iT
         gridaKlasat.col = 2
         gridaKlasat.Text = kol2

         gridaKlasat.row = i + iT
         gridaKlasat.col = 3
         gridaKlasat.Text = kol3

      Next i

      gridaKlasat.row = rrT - 1
      gridaKlasat.col = 0
      gridaKlasat.Text = "Tetevjecarja"
      gridaKlasat.CellForeColor = &HC0&
      gridaKlasat.col = 1
      gridaKlasat.Text = Space(10) & str(mungMeT)
      gridaKlasat.CellForeColor = &HC0&

      gridaKlasat.col = 2
      gridaKlasat.Text = Space(10) & str(mungPaT)
      gridaKlasat.CellForeColor = &HC0&

      gridaKlasat.col = 3
      If nrSt > 0 Then
         gridaKlasat.Text = Space(10) & Format(mesatarjaTetevjecare / nrSt, ".0")
      Else
         gridaKlasat.Text = Space(10) & "---"
      End If
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
         gridaKlasat.col = 3
         kol3 = gridaKlasat.Text

         gridaKlasat.row = i
         gridaKlasat.col = 0
         gridaKlasat.Text = kol0

         gridaKlasat.row = i
         gridaKlasat.col = 1
         gridaKlasat.Text = kol1

         gridaKlasat.row = i
         gridaKlasat.col = 2
         gridaKlasat.Text = kol2

         gridaKlasat.row = i
         gridaKlasat.col = 3
         gridaKlasat.Text = kol3

      Next i

      gridaKlasat.row = rrU - 1
      gridaKlasat.col = 0
      gridaKlasat.Text = "Cikli i ulet"
      gridaKlasat.CellForeColor = &HC0&
      gridaKlasat.col = 1
      gridaKlasat.Text = Space(10) & str(mungMeu)
      gridaKlasat.CellForeColor = &HC0&

      gridaKlasat.col = 2
      gridaKlasat.Text = Space(10) & str(mungPaU)
      gridaKlasat.CellForeColor = &HC0&

      gridaKlasat.col = 3
      If nrSu > 0 Then
         gridaKlasat.Text = Space(10) & Format(mesatarjaUlet / nrSu, ".0")
      Else
         gridaKlasat.Text = Space(10) & "---"
      End If
      gridaKlasat.CellForeColor = &HC0&

   End If

   gridaKlasat.row = gridaKlasat.Rows - 1
   gridaKlasat.col = 0
   gridaKlasat.Text = "Totali"
   gridaKlasat.CellForeColor = &H8000&
   gridaKlasat.col = 1
   gridaKlasat.Text = Space(10) & str(mungMe)
   gridaKlasat.CellForeColor = &H8000&
   gridaKlasat.col = 2
   gridaKlasat.Text = Space(10) & str(mungPa)
   gridaKlasat.CellForeColor = &H8000&
   gridaKlasat.col = 3
   If nrs > 0 Then
      gridaKlasat.Text = Space(10) & Format(mesatarja / nrs, ".0")
   Else
      gridaKlasat.Text = Space(10) & "---"
   End If
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
   Dim mungMeM          As Integer, mungPaM As Integer
   Dim mungMeT          As Integer, mungPaT As Integer
   Dim mungMeu          As Integer, mungPaU As Integer
   Dim mungMe           As Integer, mungPa As Integer
   Dim mesatarjaMesme   As Double
   Dim mesatarjaUlet    As Double
   Dim mesatarjaTetevjecare As Double
   Dim mesatarja        As Double
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
         gridaNr.col = 1
         gridaNr.row = i

         gridaKlasat.col = 3
         gridaKlasat.row = i
         If IsNumeric(gridaKlasat.Text) Then
            mesatarjaUlet = CInt(gridaNr.Text) * CDbl(gridaKlasat.Text) + mesatarjaUlet
            nrSu = nrSu + CInt(gridaNr.Text)
            mesatarja = mesatarja + CInt(gridaNr.Text) * CDbl(gridaKlasat.Text)
            nrs = nrs + CInt(gridaNr.Text)
         End If

         gridaKlasat.col = 1
         mungMeu = mungMeu + CInt(gridaKlasat.Text)
         mungMe = mungMe + CInt(gridaKlasat.Text)
         gridaKlasat.col = 2
         mungPaU = mungPaU + CInt(gridaKlasat.Text)
         mungPa = mungPa + CInt(gridaKlasat.Text)
      End If
      If Klasa = "6" Or Klasa = "7" Or Klasa = "8" Or Klasa = "9" Then

         kaCikliTetevjecare = True
         nrT = nrT + 1

         gridaNr.col = 1
         gridaNr.row = i

         gridaKlasat.col = 3
         gridaKlasat.row = i
         If IsNumeric(gridaKlasat.Text) Then
            mesatarjaTetevjecare = CInt(gridaNr.Text) * CDbl(gridaKlasat.Text) + mesatarjaTetevjecare
            nrSt = nrSt + CInt(gridaNr.Text)
            mesatarja = mesatarja + CInt(gridaNr.Text) * CDbl(gridaKlasat.Text)
            nrs = nrs + CInt(gridaNr.Text)
         End If

         gridaKlasat.col = 1
         mungMeT = mungMeT + CInt(gridaKlasat.Text)
         mungMe = mungMe + CInt(gridaKlasat.Text)
         gridaKlasat.col = 2
         mungPaT = mungPaT + CInt(gridaKlasat.Text)
         mungPa = mungPa + CInt(gridaKlasat.Text)
      End If
      If Klasa = "10" Or Klasa = "11" Or Klasa = "12" Then

         kaCikliMesem = True
         nrM = nrM + 1

         gridaNr.col = 1
         gridaNr.row = i

         gridaKlasat.col = 3
         gridaKlasat.row = i
         If IsNumeric(gridaKlasat.Text) Then
            mesatarjaMesme = CInt(gridaNr.Text) * CDbl(gridaKlasat.Text) + mesatarjaMesme
            nrsm = nrsm + CInt(gridaNr.Text)
            mesatarja = mesatarja + CInt(gridaNr.Text) * CDbl(gridaKlasat.Text)
            nrs = nrs + CInt(gridaNr.Text)
         End If

         gridaKlasat.col = 1
         mungMeM = mungMeM + CInt(gridaKlasat.Text)
         mungMe = mungMe + CInt(gridaKlasat.Text)
         gridaKlasat.col = 2
         mungPaM = mungPaM + CInt(gridaKlasat.Text)
         mungPa = mungPa + CInt(gridaKlasat.Text)
      End If

   Next i

   rrU = 0
   rrT = 0
   rrM = 0
   iT = 0
   iM = 0
   If kaCikliUlet Then
      gridaKlasat.Rows = gridaKlasat.Rows + 1
      gridaNr.Rows = gridaNr.Rows + 1
      rrU = nrU + 1
      rrT = rrU
      rrM = rrU
      iT = iT + 1
      iM = iM + 1
   End If

   If kaCikliTetevjecare Then
      gridaKlasat.Rows = gridaKlasat.Rows + 1
      gridaNr.Rows = gridaNr.Rows + 1
      rrT = rrT + nrT + 1
      rrM = rrM + nrT + 1
      iM = iM + 1
   End If

   If kaCikliMesem Then
      gridaKlasat.Rows = gridaKlasat.Rows + 1
      gridaNr.Rows = gridaNr.Rows + 1
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

         gridaKlasat.row = i
         gridaKlasat.col = 3
         kol3 = gridaKlasat.Text

         gridaKlasat.row = i + iM
         gridaKlasat.col = 0
         gridaKlasat.Text = kol0

         gridaKlasat.row = i + iM
         gridaKlasat.col = 1
         gridaKlasat.Text = kol1

         gridaKlasat.row = i + iM
         gridaKlasat.col = 2
         gridaKlasat.Text = kol2

         gridaKlasat.row = i + iM
         gridaKlasat.col = 3
         gridaKlasat.Text = kol3

      Next i

      gridaKlasat.row = rrM - 1

      gridaKlasat.col = 0
      gridaKlasat.Text = "E mesmja"
      gridaKlasat.CellForeColor = &HC0&

      gridaKlasat.col = 1
      gridaKlasat.Text = Space(10) & str(mungMeM)
      gridaKlasat.CellForeColor = &HC0&

      gridaKlasat.col = 2
      gridaKlasat.Text = Space(10) & str(mungPaM)
      gridaKlasat.CellForeColor = &HC0&

      gridaKlasat.col = 3
      If nrsm > 0 Then
         gridaKlasat.Text = Space(10) & Format(mesatarjaMesme / nrsm, ".0")
      Else
         gridaKlasat.Text = Space(10) & "---"
      End If
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

         gridaKlasat.row = i
         gridaKlasat.col = 3
         kol3 = gridaKlasat.Text

         gridaKlasat.row = i + iT
         gridaKlasat.col = 0
         gridaKlasat.Text = kol0

         gridaKlasat.row = i + iT
         gridaKlasat.col = 1
         gridaKlasat.Text = kol1

         gridaKlasat.row = i + iT
         gridaKlasat.col = 2
         gridaKlasat.Text = kol2

         gridaKlasat.row = i + iT
         gridaKlasat.col = 3
         gridaKlasat.Text = kol3

      Next i

      gridaKlasat.row = rrT - 1
      gridaKlasat.col = 0
      gridaKlasat.Text = "Nëntëvjeçarja"
      gridaKlasat.CellForeColor = &HC0&
      gridaKlasat.col = 1
      gridaKlasat.Text = Space(10) & str(mungMeT)
      gridaKlasat.CellForeColor = &HC0&

      gridaKlasat.col = 2
      gridaKlasat.Text = Space(10) & str(mungPaT)
      gridaKlasat.CellForeColor = &HC0&

      gridaKlasat.col = 3
      If nrSt > 0 Then
         gridaKlasat.Text = Space(10) & Format(mesatarjaTetevjecare / nrSt, ".0")
      Else
         gridaKlasat.Text = Space(10) & "---"
      End If
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
         gridaKlasat.col = 3
         kol3 = gridaKlasat.Text

         gridaKlasat.row = i
         gridaKlasat.col = 0
         gridaKlasat.Text = kol0

         gridaKlasat.row = i
         gridaKlasat.col = 1
         gridaKlasat.Text = kol1

         gridaKlasat.row = i
         gridaKlasat.col = 2
         gridaKlasat.Text = kol2

         gridaKlasat.row = i
         gridaKlasat.col = 3
         gridaKlasat.Text = kol3

      Next i

      gridaKlasat.row = rrU - 1
      gridaKlasat.col = 0
      gridaKlasat.Text = "Cikli i ulet"
      gridaKlasat.CellForeColor = &HC0&
      gridaKlasat.col = 1
      gridaKlasat.Text = Space(10) & str(mungMeu)
      gridaKlasat.CellForeColor = &HC0&

      gridaKlasat.col = 2
      gridaKlasat.Text = Space(10) & str(mungPaU)
      gridaKlasat.CellForeColor = &HC0&

      gridaKlasat.col = 3
      If nrSu > 0 Then
         gridaKlasat.Text = Space(10) & Format(mesatarjaUlet / nrSu, ".0")
      Else
         gridaKlasat.Text = Space(10) & "---"
      End If
      gridaKlasat.CellForeColor = &HC0&

   End If

   gridaKlasat.row = gridaKlasat.Rows - 1
   gridaKlasat.col = 0
   gridaKlasat.Text = "Totali"
   gridaKlasat.CellForeColor = &H8000&
   gridaKlasat.col = 1
   gridaKlasat.Text = Space(10) & str(mungMe)
   gridaKlasat.CellForeColor = &H8000&
   gridaKlasat.col = 2
   gridaKlasat.Text = Space(10) & str(mungPa)
   gridaKlasat.CellForeColor = &H8000&
   gridaKlasat.col = 3
   If nrs > 0 Then
      gridaKlasat.Text = Space(10) & Format(mesatarja / nrs, ".0")
   Else
      gridaKlasat.Text = Space(10) & "---"
   End If
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

