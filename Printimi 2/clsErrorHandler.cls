VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsErrorHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim titull As String
Dim objPrintim As New clsPrintim
Public Sub KontrolloGabimin(nrGabimi As Integer)
    titull = "Gabim ne printim!"
    If nrGabimi = 380 Then
        'MsgBox "Printeri juaj nuk ju lejon qe ta printoni deftesen ne formatin e saj te plote." & _
            'vbCrLf & "Ju lutemi paloseni fleten me dysh!", vbOKOnly, titull
    'ElseIf nrGabimi = 482 Then
    '   MsgBox "Nuk keni asnje printer te instaluar!", vbOKOnly, titull
    End If
End Sub
