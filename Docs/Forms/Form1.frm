VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmFlexGrid 
   Caption         =   "flexgrid"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5580
   ScaleWidth      =   7665
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Height          =   375
      Left            =   5160
      TabIndex        =   3
      Top             =   4800
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   5160
      Width           =   2535
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      Format          =   66715649
      CurrentDate     =   38304
   End
End
Attribute VB_Name = "frmFlexGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
  With ucFlexGrid1
     .Rows = 2
     .Cols = 5
     
     .TextMatrix(0, 1) = "Emri Nxenesit"
     .TextMatrix(0, 2) = "Matematike"
     .TextMatrix(0, 3) = "Matematike"
     .TextMatrix(0, 4) = "Letersi"
     
     .MergeRow(0) = True
     .MergeCells = flexMergeRestrictRows
     
     For i = 1 To 4
         .RowHeight(0) = 400
         .ColAlignment(i) = flexAlignCenterCenter
     Next i
     .ColWidth(1) = 500
     .ColWidth(2) = 100
     .ColWidth(3) = 100
     
  End With
  
  
  With MSFlexGrid1
  End With
  
End Sub

'Private Sub ucFlexGrid1_LeaveCell()
'   Dim x As String
'
'   x = ucFlexGrid1.Text
'
'   MsgBox x
'
'
'End Sub


