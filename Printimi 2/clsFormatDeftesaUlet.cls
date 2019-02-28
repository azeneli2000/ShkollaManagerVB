VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFormatDeftesaUlet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim Lendet(10) As String
Dim objPrintim As New clsPrintim
Dim objErrHandler As New clsErrorHandler



'Do te beje formatimin e te dhenave per deftesen e ciklit te ulet
Public Sub FormatDeftesaUlet(vekTeDhenash() As String, tabNotash() As String)
    If objPrintim.PrinterIsInstalled Then
    On Error GoTo Gabimi
        Printer.PaperSize = vbPRPSA4
        Printer.Orientation = vbPRORPortrait
        Printer.ScaleMode = 6
        ' Printo faqen e pare te defteses
        objPrintim.PrintLeft 63, 73, vekTeDhenash(6)
        
        ' Mesazh qe i jepet perdoruesit per te printuar gjeneralitet
        MsgBox "Futeni deftesen ne printer ne menyre qe te printohen gjeneralitet.", vbOKOnly, "Printimi i defteses"
        objPrintim.PrintLeft 234, 70, vekTeDhenash(6)
        
        ' Nese perdoruesi anullon printimin e dokumentit atehere
        ' dil nga procedura
        If objPrintim.Gabim = True Then
            Exit Sub
        End If
        
        objPrintim.PrintLeft 28, 89, vekTeDhenash(5)
        objPrintim.PrintLeft 34, 104, vekTeDhenash(4)
        objPrintim.PrintLeft 32, 179, vekTeDhenash(0) + " " + vekTeDhenash(2) + " " + vekTeDhenash(1)
        objPrintim.PrintLeft 30, 187, vekTeDhenash(6)
        objPrintim.PrintLeft 73, 192, vekTeDhenash(7)
        objPrintim.PrintLeft 117, 192, vekTeDhenash(8)
        objPrintim.PrintLeft 116, 102, vekTeDhenash(9)
        objPrintim.PrintLeft 131, 200, vekTeDhenash(10)
        Printer.EndDoc
        ' Afisho nje mesazh qe deftesa te kthehet nga ana tjeter
        MsgBox "Ju lutemi kthejeni deftesen nga ana tjeter ne printer", vbOKOnly, "Duke printuar..."
        
        ' Printo faqen prapa te defteses
        Dim i, j, k As Integer
        Dim nukEkziston As Boolean
        k = 0
        j = 0
        ' Kontrollo korrespondencat e emrave te lendeve
        Do While tabNotash(j, 0) <> ""
        'For i = 0 To 9
            nukEkziston = True
            For i = 0 To 9
            'Do While tabNotash(j, 0) <> ""
                ' Nese emri i lendes ne gride gjendet ne vektorin Lendet
                If Lendet(i) = tabNotash(j, 0) Then
                    objPrintim.PrintLeft 59, 31 + (7 * i), objPrintim.MerrNote(tabNotash(j, 1))
                    objPrintim.PrintLeft 78, 31 + (7 * i), objPrintim.MerrNote(tabNotash(j, 2))
                    objPrintim.PrintLeft 97, 31 + (7 * i), objPrintim.MerrNote(tabNotash(j, 3))
                    nukEkziston = False
                End If
            Next
            If nukEkziston Then
                If k <= 7 Then
                    objPrintim.PrintLeft 29, 100 + k, tabNotash(j, 0)
                    objPrintim.PrintLeft 59, 100 + k, objPrintim.MerrNote(tabNotash(j, 1))
                    objPrintim.PrintLeft 78, 100 + k, objPrintim.MerrNote(tabNotash(j, 2))
                    objPrintim.PrintLeft 97, 100 + k, objPrintim.MerrNote(tabNotash(j, 3))
                    k = k + 7
                End If
            End If
            j = j + 1
        Loop
        Printer.EndDoc
    End If
    'frmPrint1.lblEmri.Caption = vekTeDhenash(0) + " " + vekTeDhenash(1)
    'frmPrint1.lblAmza.Caption = vekTeDhenash(2)
    'PrintoDeftesen
Gabimi:     objErrHandler.KontrolloGabimin Err.Number
End Sub


Private Sub Class_Initialize()
    Lendet(0) = "Gjuhe shqipe"
    Lendet(1) = "Lexim letrar"
    Lendet(2) = "Matematike"
    Lendet(3) = "Histori"
    Lendet(4) = "Dituri natyre"
    Lendet(5) = "Edukate shoqerore"
    Lendet(6) = "Edukim fizik"
    Lendet(7) = "Vizatim"
    Lendet(8) = "Edukim muzikor"
    Lendet(9) = "Mesim pune"
End Sub

Public Sub PrintoNotatUlet(tabNotash() As String)
    ' Printo faqen prapa te defteses
    Dim i, j, k As Integer
    Dim nukEkziston As Boolean
    k = 0
    j = 0
    ' Kontrollo korrespondencat e emrave te lendeve
    Do While tabNotash(j, 0) <> ""
    'For i = 0 To 9
        nukEkziston = True
        For i = 0 To 9
        'Do While tabNotash(j, 0) <> ""
            ' Nese emri i lendes ne gride gjendet ne vektorin Lendet
            If Lendet(i) = tabNotash(j, 0) Then
                objPrintim.PrintLeft 59, 31 + (7 * i), objPrintim.MerrNote(tabNotash(j, 1))
                objPrintim.PrintLeft 78, 31 + (7 * i), objPrintim.MerrNote(tabNotash(j, 2))
                objPrintim.PrintLeft 97, 31 + (7 * i), objPrintim.MerrNote(tabNotash(j, 3))
                nukEkziston = False
            End If
        Next
        If nukEkziston Then
            If k <= 7 Then
                objPrintim.PrintLeft 29, 100 + k, tabNotash(j, 0)
                objPrintim.PrintLeft 59, 100 + k, objPrintim.MerrNote(tabNotash(j, 1))
                objPrintim.PrintLeft 78, 100 + k, objPrintim.MerrNote(tabNotash(j, 2))
                objPrintim.PrintLeft 97, 100 + k, objPrintim.MerrNote(tabNotash(j, 3))
                k = k + 7
            End If
        End If
        j = j + 1
    Loop
End Sub
