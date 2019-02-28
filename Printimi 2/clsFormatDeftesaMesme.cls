VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFormatDeftesaMesme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim objPrintim As New clsPrintim
Dim objErrHandler As New clsErrorHandler
Public Sub FormatDeftesaMesme(vekTeDhenash() As String, tabNotash() As String)
    Dim lart, gjere As Single
    On Error GoTo Gabimi
    Printer.PaperSize = vbPRPSA3
    Printer.Orientation = 2
    MsgBox "Futeni deftesen ne printer ne menyre qe te printohen gjeneralitet.", vbOKOnly, "Printimi i defteses"
    objPrintim.PrintLeft 234, 70, vekTeDhenash(6)
        
    ' Nese perdoruesi anullon printimin e dokumentit atehere
    ' dil nga procedura
    If objPrintim.Gabim = True Then
        Exit Sub
    End If
        
    objPrintim.PrintLeft 240, 69, vekTeDhenash(7)
    objPrintim.PrintLeft 195, 83, vekTeDhenash(5)
    objPrintim.PrintLeft 200, 95, vekTeDhenash(2)
    objPrintim.PrintLeft 196, 159, vekTeDhenash(0) + " " + vekTeDhenash(4) + " " + vekTeDhenash(1)
    objPrintim.PrintLeft 200, 169, vekTeDhenash(6)
    objPrintim.PrintLeft 279, 169, vekTeDhenash(6)
    objPrintim.PrintLeft 200, 180, vekTeDhenash(8)
    objPrintim.PrintLeft 206, 189, objPrintim.MerrKlase(vekTeDhenash(10))
    objPrintim.PrintLeft 240, 200, objPrintim.MerrKlasePara(vekTeDhenash(10))
    Printer.EndDoc
    If faqePrapa = True Then
    ' Afisho nje mesazh qe deftesa te kthehet nga ana tjeter
        MsgBox "Kthejeni deftesen nga ana tjeter ne printer", vbOKOnly, "Printimi i defteses"
        ' Printo faqen prapa te defteses
        Dim i, j As Integer
        i = 0
        Do While tabNotash(i, 0) = ""
            objPrintim.PrintLeft 40, 35 + (5.9 * i), tabNotash(i, 0)
            ' Nese perdoruesi anullon printimin e dokumentit atehere
            ' dil nga procedura
            If objPrintim.Gabim = True Then
                Exit Sub
            End If
            objPrintim.PrintLeft 130, 35 + (5.9 * i), objPrintim.MerrNote(tabNotash(i, 1))
            objPrintim.PrintLeft 169, 35 + (5.9 * i), objPrintim.MerrNote(tabNotash(i, 2))
            objPrintim.PrintLeft 206, 35 + (5.9 * i), objPrintim.MerrNote(tabNotash(i, 3))
        Loop
        Printer.EndDoc
    End If
Gabimi:     objErrHandler.KontrolloGabimin (Err.Number)
    If Err.Number = 380 Then
        MsgBox "Printeri juaj nuk mund te printoje ne formatin A3, qe duhet per printimin" _
        & " e defteses. Ju lutemi nderroni printerin!", vbInformation + vbOKOnly, "Gabim ne printim!"
    End If
End Sub

Public Sub FormatDeftesaMesme1(vekTeDhenash() As String, tabNotash() As String)
    On Error GoTo Gabimi
    If objPrintim.PrinterIsInstalled Then
        MsgBox "Futeni deftesen ne printer, te palosur ne menyre qe te printohen gjeneralitetet", vbOKOnly, "Printimi i defteses"
        objPrintim.PrintLeft 63, 69, vekTeDhenash(6)
        ' Nese perdoruesi anullon printimin e dokumentit atehere
        ' dil nga procedura
        If objPrintim.Gabim = True Then
            Exit Sub
        End If
        objPrintim.PrintLeft 30, 82, vekTeDhenash(5)
        objPrintim.PrintLeft 33, 95, vekTeDhenash(3)
        objPrintim.PrintLeft 110, 95, vekTeDhenash(9)
        objPrintim.PrintLeft 30, 159, vekTeDhenash(0) & " " & vekTeDhenash(2) & " " & vekTeDhenash(1)
        objPrintim.PrintLeft 33, 169, vekTeDhenash(8)
        objPrintim.PrintLeft 110, 169, vekTeDhenash(6)
        objPrintim.PrintLeft 35, 190, objPrintim.MerrKlase(vekTeDhenash(10))
        objPrintim.PrintLeft 70, 200, objPrintim.MerrKlasePara(vekTeDhenash(10))
        Printer.EndDoc
        MsgBox "Kthejeni deftesen nga ana tjeter ne printer, ne menyre qe te printohen " & _
            "emertimet e notave dhe notat e semestrit te pare", vbOKOnly, "Printimi i defteses"
        Dim i As Integer
        i = 0
        Do While tabNotash(i, 0) <> ""
            objPrintim.PrintLeft 40, 35 + (7.7 * i), tabNotash(i, 0)
        
            ' Nese perdoruesi anullon printimin e dokumentit atehere
            ' dil nga procedura
            If objPrintim.Gabim = True Then
                Exit Sub
            End If
            objPrintim.PrintLeft 130, 35 + (7.7 * i), objPrintim.MerrNote(tabNotash(i, 1))
            i = i + 1
        Loop
        Printer.EndDoc
        MsgBox "Kthejeni deftesen, ne menyre qe te printohen notat" & _
            "e semestrit te dyte dhe notat vjetore"
        i = 0
        Do While tabNotash(i, 0) <> ""
            objPrintim.PrintLeft 5, 35 + (7.7 * i), objPrintim.MerrNote(tabNotash(i, 2))
        
            ' Nese perdoruesi anullon printimin e dokumentit atehere
            ' dil nga procedura
            If objPrintim.Gabim = True Then
                Exit Sub
            End If
            objPrintim.PrintLeft 35, 35 + (7.7 * i), objPrintim.MerrNote(tabNotash(i, 3))
            i = i + 1
        Loop
        Printer.EndDoc
    End If
Gabimi:     objErrHandler.KontrolloGabimin Err.Number
End Sub

Public Sub PrintoNotatMesme(tabNotash() As String)
    Dim i, j As Integer
    i = 0
    Do While tabNotash(i, 0) = ""
        objPrintim.PrintLeft 40, 35 + (5.9 * i), tabNotash(i, 0)
        ' Nese perdoruesi anullon printimin e dokumentit atehere dil nga metoda
        If objPrintim.Gabim = True Then
            Exit Sub
        End If
        objPrintim.PrintLeft 130, 35 + (5.9 * i), objPrintim.MerrNote(tabNotash(i, 1))
        objPrintim.PrintLeft 169, 35 + (5.9 * i), objPrintim.MerrNote(tabNotash(i, 2))
        objPrintim.PrintLeft 206, 35 + (5.9 * i), objPrintim.MerrNote(tabNotash(i, 3))
    Loop
End Sub

Public Sub EndDoc()
    Printer.EndDoc
End Sub

