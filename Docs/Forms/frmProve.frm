VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmProve 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6630
   ClientLeft      =   2655
   ClientTop       =   4260
   ClientWidth     =   9285
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6630
   ScaleWidth      =   9285
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   3840
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      Format          =   19791873
      CurrentDate     =   38302
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1575
      Left            =   480
      TabIndex        =   2
      Top             =   1320
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   2778
      _Version        =   393216
      Cols            =   7
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   4680
      TabIndex        =   1
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "frmProve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    
  loadForm Me
  
  With MSFlexGrid1
  
       .Cols = 8
       
       .Row = 0
       
       .col = 1
       .Text = "Matematike"
       
       .col = 2
       .Text = "Matematike"
       
       .col = 3
       .Text = "Matematike"
       
       .col = 4
       .Text = "Matematike"
       
       .col = 5
       .Text = "Matematike"
       
       .col = 6
       .Text = "Matematike"
       
       .col = 7
       .Text = "Matematike"
       
      
       
       .MergeRow(0) = True
       
       .MergeCells = flexMergeRestrictColumns
       
       
       For i = 1 To 7
           .ColWidth(i) = 200
       Next i
       
  End With
  
  
  
End Sub
