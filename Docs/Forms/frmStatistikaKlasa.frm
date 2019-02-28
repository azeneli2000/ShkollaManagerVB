VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmStatistikaKlasa 
   Caption         =   "Statistika - Klasa"
   ClientHeight    =   10365
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
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
      TabIndex        =   14
      Top             =   0
      Width           =   2655
      Begin VB.CommandButton cmdWebsiste 
         BackColor       =   &H80000009&
         Caption         =   "Website"
         Height          =   375
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton cmdEmail 
         BackColor       =   &H80000009&
         Caption         =   "E-mail"
         Height          =   375
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   960
         Width           =   1095
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   120
         Picture         =   "frmStatistikaKlasa.frx":0000
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
         TabIndex        =   18
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
         TabIndex        =   17
         Top             =   120
         Width           =   2415
      End
   End
   Begin MSFlexGridLib.MSFlexGrid gridaLendet 
      Height          =   2775
      Left            =   7560
      TabIndex        =   9
      Top             =   1560
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   4895
      _Version        =   393216
      Rows            =   9
      FixedCols       =   0
      ScrollBars      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid gridaKlasat 
      Height          =   2775
      Left            =   360
      TabIndex        =   8
      Top             =   1560
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   4895
      _Version        =   393216
      Rows            =   11
      Cols            =   4
      FixedCols       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton cmdNdertoGraf2 
      BackColor       =   &H80000009&
      Caption         =   "Nderto "
      Height          =   375
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4680
      Width           =   975
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
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9000
      Width           =   2295
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
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9000
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
      Left            =   3840
      TabIndex        =   1
      Top             =   840
      Width           =   2535
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
      Height          =   420
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   720
      Width           =   2415
   End
   Begin MSChart20Lib.MSChart grafKlase 
      Height          =   4095
      Index           =   0
      Left            =   120
      OleObjectBlob   =   "frmStatistikaKlasa.frx":6F75
      TabIndex        =   6
      Top             =   4680
      Width           =   6855
   End
   Begin MSChart20Lib.MSChart grafLende 
      Height          =   4095
      Index           =   1
      Left            =   7200
      OleObjectBlob   =   "frmStatistikaKlasa.frx":8869
      TabIndex        =   10
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
      TabIndex        =   11
      Top             =   120
      Width           =   9375
      Begin VB.Label Label2 
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
         Left            =   3480
         TabIndex        =   13
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label1 
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
         Left            =   600
         TabIndex        =   12
         Top             =   480
         Width           =   2535
      End
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   15240
      Y1              =   4560
      Y2              =   4560
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


Private Sub cmdDil_Click()
    Unload Me
    Set active_form = Nothing
End Sub

Private Sub cmdNdertoGraf1_Click()
    Dim mesatarja As String
    Dim lenda As String
    grafLende(1).RowCount = gridaLendet.Rows - 1
    
    
    
    
    Dim I As Integer
    For I = 1 To gridaLendet.Rows - 1
      'lenda = Trim(lista2.List(i - 1))
      gridaLendet.col = 0
      gridaLendet.row = I
      lenda = gridaLendet.Text
      grafLende(1).row = I
      grafLende(1).RowLabel = lenda
      'grafLende.data = Val(lista3.List(i - 1))
      gridaLendet.col = 1
      gridaLendet.row = I
      mesatarja = gridaLendet.Text
      grafLende(1).data = Val(mesatarja)
    Next I
    
    
End Sub

Private Sub cmdNdertoGraf2_Click()
    Dim mesatarja As String
    Dim Klasa As String
    grafKlase(0).RowCount = gridaKlasat.Rows - 1
    
    
    Dim I As Integer
    For I = 1 To gridaKlasat.Rows - 1
      'lenda = Trim(lista2.List(i - 1))
      gridaKlasat.col = 0
      gridaKlasat.row = I
      Klasa = gridaKlasat.Text
      grafKlase(0).row = I
      grafKlase(0).RowLabel = Klasa
      'grafLende.data = Val(lista3.List(i - 1))
      gridaKlasat.col = 3
      gridaKlasat.row = I
      mesatarja = gridaKlasat.Text
      grafKlase(0).data = Val(mesatarja)
    Next I
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

Private Sub cmdShfaq_Click()
   'lista1.Clear
   objektGabimi.kapGabimin
   objektGabimi.menazhim_gabimi
   If objektGabimi.mvarGabimi = 0 Then
      
      pastroGriden gridaKlasat
      pastroGriden gridaLendet
      objectInitialization STATISTIKA_KLASA_SHFAQ
      renditGriden gridaKlasat
      formatoGrida gridaKlasat, 2775
   End If
   
End Sub

Private Sub Command1_Click()
    CallHelp indeksHelp
End Sub

Private Sub Form_Load()
   loadForm Me
   viti
   'txtVitiShkollor.Text = cboVitiShkollor.Text
   Image1.Picture = LoadPicture(adresaLogo)
   lblShkolla.Caption = emerShkolla
   mbushKomboBox
   inicializoGridat
   loadComboItems
   
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
   Dim I, j             As Integer

   For I = 0 To liste.ListCount - 1
      vektorStr(I + 1) = liste.List(I)
      ' vektorKlasa(i + 1) = ktheKlase(liste.List(i))
      ' vektorIndeksi(i + 1) = ktheIndeksi(liste.List(i))
   Next I
   Dim nr               As Integer
   nr = liste.ListCount
   For I = 1 To nr - 1
      For j = I + 1 To nr
         klasaI = ktheKlase(vektorStr(I))
         klasaJ = ktheKlase(vektorStr(j))
         If klasaI > klasaJ Then
            ruajRreshtin = vektorStr(I)
            vektorStr(I) = vektorStr(j)
            vektorStr(j) = ruajRreshtin
         Else
            If klasaI = klasaJ Then
               indeksiI = ktheIndeksi(vektorStr(I))
               indeksiJ = ktheIndeksi(vektorStr(j))
               If indeksiI > indeksiJ Then
                  ruajRreshtin = vektorStr(I)
                  vektorStr(I) = vektorStr(j)
                  vektorStr(j) = ruajRreshtin
               End If
            End If
         End If
      Next j
   Next I
   liste.Clear
   For I = 1 To nr
      liste.AddItem vektorStr(I)
   Next I



End Sub

Private Function ktheKlase(rreshti As String) As Double
   Dim Klasa            As String
   s3 = Trim(rreshti)
   Dim s2               As String
   s2 = ""
   Dim ugjet            As Boolean
   Dim I, j             As Integer
   ugjet = False
   I = 1
   Do While Not ugjet
      If Mid(s3, I, 1) <> " " Then
         s2 = s2 & Mid(s3, I, 1)
         I = I + 1
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
   Dim I, j             As Integer
   s2 = ""
   ugjet = False
   I = 1
   Do While Not ugjet
      If Mid(s3, I, 1) <> " " Then
         s2 = s2 & Mid(s3, I, 1)
         I = I + 1
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
   Dim I                As Integer
   Dim Klasa            As String
   I = gridaKlasat.RowSel
   gridaKlasat.row = I
   gridaKlasat.col = 0
   Klasa = gridaKlasat.Text
   If I <> 0 Then
      If Klasa <> "" Then
         pastroGriden gridaLendet
         objectInitialization STATISTIKA_LENDET_MESATARET
         formatoGrida gridaLendet, 2775
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
    
    Dim I As Integer
    Dim rreshti As String
    For I = 0 To lista1.ListCount - 1
        rreshti = lista1.List(I)
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
    Next I
End Sub

Private Sub mbushListen(vektori() As String, M As Integer, lista As ListBox)
    
   Dim I As Integer
   For I = 1 To M
    If vektori(I) <> "" Then
        lista.AddItem vektori(I)
    End If
    I = I + 1
   Next I
   
    
End Sub

Private Function ktheKlasenString(rreshti As String) As String
   Dim Klasa            As String
   s3 = Trim(rreshti)
   Dim s2               As String
   s2 = ""
   Dim ugjet            As Boolean
   Dim I, j             As Integer
   ugjet = False
   I = 1
   Do While Not ugjet
      If Mid(s3, I, 1) <> " " Then
         s2 = s2 & Mid(s3, I, 1)
         I = I + 1
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

   ktheKlasenString = Klasa
End Function

Private Sub inicializoGridat()
    
    
    gridaKlasat.row = 0
    gridaKlasat.col = 0
    gridaKlasat.ColWidth(0) = 1000
    gridaKlasat.ColWidth(1) = 2000
    gridaKlasat.ColWidth(2) = 2000
    gridaKlasat.ColWidth(3) = 1550
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
End Sub

Private Sub renditGriden(grida As MSFlexGrid)

   Dim gjatesia         As Integer
   Dim I                As Integer
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
   For I = 1 To gjatesia - 2
      For j = I + 1 To gjatesia - 1

         gridaKlasat.col = 0
         gridaKlasat.row = I
         klasaI = gridaKlasat.Text
         gridaKlasat.row = j
         klasaJ = gridaKlasat.Text

         If Val(klasaI) > Val(klasaJ) Then
            ruajKlasa = klasaI
            klasaI = klasaJ
            klasaJ = ruajKlasa
            gridaKlasat.col = 0
            gridaKlasat.row = I
            gridaKlasat.Text = klasaI
            gridaKlasat.row = j
            gridaKlasat.Text = klasaJ

            gridaKlasat.col = 1
            gridaKlasat.row = I
            mungesa_pa_arsyeI = gridaKlasat.Text
            grida.row = j
            mungesa_pa_arsyeJ = gridaKlasat.Text
            gridaKlasat.row = I
            gridaKlasat.Text = mungesa_pa_arsyeJ
            grida.row = j
            gridaKlasat.Text = mungesa_pa_arsyeI

            gridaKlasat.col = 2
            gridaKlasat.row = I
            mungesa_me_arsyeI = gridaKlasat.Text
            grida.row = j
            mungesa_me_arsyeJ = gridaKlasat.Text
            gridaKlasat.row = I
            gridaKlasat.Text = mungesa_me_arsyeJ
            grida.row = j
            gridaKlasat.Text = mungesa_me_arsyeI

            gridaKlasat.col = 3
            gridaKlasat.row = I
            mesatarjaI = gridaKlasat.Text
            grida.row = j
            mesatarjaJ = gridaKlasat.Text
            gridaKlasat.row = I
            gridaKlasat.Text = mesatarjaJ
            grida.row = j
            gridaKlasat.Text = mesatarjaI

         ElseIf Val(klasaI) = Val(klasaJ) Then
            If indeksiI > indeksiJ Then
               ruajKlasa = klasaI
               klasaI = klasaJ
               klasaJ = ruajKlasa
               gridaKlasat.col = 0
               gridaKlasat.row = I
               gridaKlasat.Text = klasaI
               gridaKlasat.row = j
               gridaKlasat.Text = klasaJ

               gridaKlasat.col = 1
               gridaKlasat.row = I
               mungesa_pa_arsyeI = gridaKlasat.Text
               grida.row = j
               mungesa_pa_arsyeJ = gridaKlasat.Text
               gridaKlasat.row = I
               gridaKlasat.Text = mungesa_pa_arsyeJ
               grida.row = j
               gridaKlasat.Text = mungesa_pa_arsyeI

               gridaKlasat.col = 2
               gridaKlasat.row = I
               mungesa_me_arsyeI = gridaKlasat.Text
               grida.row = j
               mungesa_me_arsyeJ = gridaKlasat.Text
               gridaKlasat.row = I
               gridaKlasat.Text = mungesa_me_arsyeJ
               grida.row = j
               gridaKlasat.Text = mungesa_me_arsyeI

               gridaKlasat.col = 3
               gridaKlasat.row = I
               mesatarjaI = gridaKlasat.Text
               grida.row = j
               mesatarjaJ = gridaKlasat.Text
               gridaKlasat.row = I
               gridaKlasat.Text = mesatarjaJ
               grida.row = j
               gridaKlasat.Text = mesatarjaI
            End If
         Else
         End If
      Next j
   Next I

End Sub

Private Sub formatoGrida(grida As MSFlexGrid, lartesia As Long)
    
    Dim l As Long
    Dim I As Integer
    I = grida.Rows
    l = I * grida.RowHeight(0) + 100
    If l < lartesia Then
        grida.Height = l
    End If
    
    
End Sub
Private Sub viti()
   Dim d, muai, viti    As String
   d = Now
   Dim M, v             As Integer
   muai = Mid(d, 4, 2)
   viti = Mid(d, 7, 4)
   M = Val(muai)
   v = Val(viti)
   cboVitiShkollor.Text = gjej_vitin(M, v)

End Sub

Private Sub pastroGriden(grida As MSFlexGrid)
    
    Dim I As Integer
    Dim j As Integer
    Dim nr_rreshta As Integer
    Dim nr_shtylla As Integer
    nr_rreshta = grida.Rows
    nr_shtylla = grida.Cols
    For I = 1 To nr_rreshta - 1
        For j = 0 To nr_shtylla - 1
            grida.row = I
            grida.col = j
            grida.Text = ""
        Next j
    Next I
    
End Sub
