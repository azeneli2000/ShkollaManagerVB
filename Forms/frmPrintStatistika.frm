VERSION 5.00
Begin VB.Form frmPrintStatistika 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Printo "
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5835
   Icon            =   "frmPrintStatistika.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   5835
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00FFFFFF&
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmPrintStatistika.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1200
      Width           =   1815
   End
   Begin VB.CommandButton cmdDil 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3840
      Picture         =   "frmPrintStatistika.frx":3844
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1800
      Width           =   1815
   End
   Begin VB.OptionButton optStat 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Printo statistikat e klasave"
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   360
      Value           =   -1  'True
      Width           =   3015
   End
   Begin VB.OptionButton optMesLende 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Printo mesataret ne lendet e klases se zgjedhur"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   3255
   End
   Begin VB.OptionButton optGrafMesLende 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Printo grafikun e mesatares per cdo lende"
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   2280
      Width           =   3255
   End
   Begin VB.OptionButton optGrafMesKlase 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Printo grafikun e mesatares se klasave"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   1680
      Width           =   3135
   End
End
Attribute VB_Name = "frmPrintStatistika"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public choice As Integer


Private Sub cmdDil_Click()
    choice = 0
    Me.Hide
End Sub

Private Sub cmdPrint_Click()
    If Me.optStat.Value = True Then
        choice = 1
    ElseIf Me.optMesLende.Value = True Then
        choice = 2
    ElseIf Me.optGrafMesKlase.Value = True Then
        choice = 3
    ElseIf Me.optGrafMesLende.Value = True Then
        choice = 4
    Else: choice = 0
    End If
    Me.Hide
    
End Sub

Private Sub Form_Load()
    ' Behet pozicionimi i formes ne mes te ekranit
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
End Sub

