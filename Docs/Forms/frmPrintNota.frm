VERSION 5.00
Begin VB.Form frmPrintNota 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Printimi i notave"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6540
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   6540
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optNoteSem2 
      Caption         =   "Printo notat e semestrit te dyte"
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   1920
      Width           =   2895
   End
   Begin VB.OptionButton optDeftesa 
      Caption         =   "Printo Deftesen"
      Height          =   615
      Left            =   480
      TabIndex        =   4
      Top             =   2520
      Width           =   2895
   End
   Begin VB.OptionButton optNoteSem1 
      Caption         =   "Printo notat e semestrit te pare"
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   1320
      Width           =   2895
   End
   Begin VB.OptionButton optNoteMom 
      Caption         =   "Printo notat momentale"
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   720
      Value           =   -1  'True
      Width           =   2895
   End
   Begin VB.CommandButton cmdDil 
      Caption         =   "Dil"
      Height          =   495
      Left            =   4560
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Printo"
      Height          =   495
      Left            =   4560
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "frmPrintNota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objKonNota As New frmKonsultimeNota
Public Nota As Integer

Private Sub cmdDil_Click()
    frmKonsultimeNota.PrintNota = 0
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    'Set objKonNota = CreateObject(frmKonsultimeNota)
    If optNoteMom.Value = True Then
        Nota = 1
    ElseIf optNoteSem1.Value = True Then
        Nota = 2
    ElseIf optNoteSem2.Value = True Then
        Nota = 3
    ElseIf optDeftesa.Value = True Then
        Nota = 4
    End If
    Me.Hide
End Sub
