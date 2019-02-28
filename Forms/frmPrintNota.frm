VERSION 5.00
Begin VB.Form frmPrintNota 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Printo"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5940
   Icon            =   "frmPrintNota.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   5940
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optPerfSem2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Printo perfundimet e semestrit te dyte"
      Height          =   495
      Left            =   2640
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.OptionButton optPerfSem1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Printo perfundimet e semestrit te pare"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   3255
   End
   Begin VB.OptionButton optDeftese 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Printo deftesen"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   3255
   End
   Begin VB.OptionButton optNotePerf 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Printo vertetimin per perfundimet vjetore"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Value           =   -1  'True
      Width           =   3255
   End
   Begin VB.CommandButton cmdDil 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3840
      Picture         =   "frmPrintNota.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00FFFFFF&
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmPrintNota.frx":674C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      Width           =   1695
   End
End
Attribute VB_Name = "frmPrintNota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public nota As Integer

Private Sub cmdDil_Click()
    nota = 0
    Me.Hide
End Sub

Private Sub cmdPrint_Click()
    'Set objKonNota = CreateObject(frmKonsultimeNota)
    If optNotePerf.Value = True Then
        nota = 1
    ElseIf optPerfSem1.Value = True Then
        nota = 2
    ElseIf optPerfSem2.Value = True Then
        nota = 3
    ElseIf optDeftese.Value = True Then
        nota = 4
    End If
    Me.Hide
End Sub

Private Sub Form_Load()
    ' Behet pozicionimi i formes ne mes te ekranit
    Me.left = (Screen.Width - Me.Width) / 2
    Me.top = (Screen.Height - Me.Height) / 2
End Sub

