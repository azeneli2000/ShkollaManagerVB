VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrintMomentale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim objPrintim As New clsPrintim
Public Sub PrintNotaMomentale()
    Dim fontib As Boolean
    Dim i, j As Integer
    objPrintim.PrintLeft 20, 20, "LENDET", True
    objPrintim.PrintLeft 70, 20, "NOTAT DHE MUNGESAT", True
    For i = 1 To Cgrid1.Height / 255
        ' Printon emrin e lendes korrente
        objPrintim.PrintLeft 20, 30 + (10 * i), Lendet.Text(i, 1), True
        ' Printon notat per lenden korrente
        For j = 1 To Cgrid1.Width / 300
            ' Percaktohet nga ngjyra e cellit nese do te printohet nota si bold
            If Cgrid1.CellForeColor(i, j) = 16711680 Then
                fontib = True
            Else: fontib = False
            End If
                objPrintim.PrintLeft 70 + (5 * j), 30 + (10 * i), Cgrid1.Text(i, j), fontib
        Next
    Next
    Printer.EndDoc

End Sub
