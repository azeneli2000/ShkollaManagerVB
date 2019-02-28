VERSION 5.00
Begin VB.Form frmAbout 
   Caption         =   "Rreth Programit"
   ClientHeight    =   4350
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   ScaleHeight     =   4350
   ScaleWidth      =   6420
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   600
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   585
      ScaleWidth      =   465
      TabIndex        =   1
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Tel : 04 251972"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Website : www.visioninfosolution.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   3120
      Width           =   3735
   End
   Begin VB.Label Label2 
      Caption         =   "Adresa Elektronike : info@visioninfosolution.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2760
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "                  Te githa te drejtat e rezervuara"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   2040
      Width           =   6015
   End
   Begin VB.Label lblShenime 
      Caption         =   "  Copyright  ©  2004      Vision Info Solution "
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
      Left            =   0
      TabIndex        =   3
      Top             =   1560
      Width           =   6375
   End
   Begin VB.Label lblVersioni 
      Caption         =   "Versioni: 1.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblTitle 
      Caption         =   "Programi ""Shkolla Menaxher"""
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   0
      Top             =   360
      Width           =   3855
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  loadForm Me
  
  'lblVersionNo.Caption = App.Major & App.Minor
  Me.Width = 6540
  Me.Height = 4860
End Sub

