VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrintim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Const LOGPIXELSX = 88
Private Const LOGPIXELSY = 90
Private Const PHYSICALOFFSETX = 112
Private Const PHYSICALOFFSETY = 113
Option Explicit

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long

Public Gabim As Boolean
Dim objErrHandler As New clsErrorHandler
Dim LeftMargin As Long
Dim TopMargin As Long

' Metode qe percakton marxhinat specifike te printerit ne milimetra.
Private Sub GetPrinterOffsets()
    Dim dpiX As Long, dpiY As Long
    Dim marginLeftPix As Long, marginTopPix As Long
    Dim lart, gjere As Single
    dpiX = GetDeviceCaps(Printer.hDC, LOGPIXELSX)
    dpiY = GetDeviceCaps(Printer.hDC, LOGPIXELSY)
    
    lart = Printer.Height / 567
    gjere = Printer.Width / 567
    ' Marxhina majtas e printerit ne pixel
    marginLeftPix = GetDeviceCaps(Printer.hDC, PHYSICALOFFSETX)
    ' Marxhina siper e printerit ne pixel
    marginTopPix = GetDeviceCaps(Printer.hDC, PHYSICALOFFSETY)
    ' I kalojme keto marxhina ne milimetra
    LeftMargin = (marginLeftPix / dpiX) * 25.4
    TopMargin = (marginTopPix / dpiY) * 25.4
End Sub

Private Sub Class_Initialize()
    On Error GoTo Gabimi
    ' Vendoset ScaleMode ne milimetra
    Printer.ScaleMode = vbMillimeters
    'lexon pjesen qe nuk mund te printohet nga printeri
    GetPrinterOffsets
Gabimi:     objErrHandler.KontrolloGabimin Err.Number
End Sub

' Funksion qe kontrollon nese ka printer te instaluar
Public Function PrinterIsInstalled() As Boolean
    Dim dummy As String
    On Error Resume Next
    dummy = Printer.DeviceName
    If Err.Number Then
        MsgBox "Nuk ka asnje printer te instaluar", vbExclamation + vbOKOnly, _
        "Gabim ne printim"
        PrinterIsInstalled = False
    Else
        PrinterIsInstalled = True
    End If
End Function


' Metode qe ben te mundur printimin e nje teksti Text ne pozicionin e
' percaktuar nga parametrat XPos dhe YPos. Ne varesi te situates, ky printim
' mund te behet ne bold ose jo dhe gjithashtu mund te ndryshoje edhe madhesia
' e shkrimit dhe gjithashtu nese shkronjat do jene korsive
Public Sub PrintLeft(XPos As Long, YPos As Long, Tekst As String, Optional Fontibold As Boolean, Optional Madhesi As Integer, Optional Italike As Boolean, Optional nenvizuar As Boolean)
    Gabim = False
    On Error GoTo Gabimi
    If XPos >= 0 Then
        Printer.CurrentX = XPos - LeftMargin
    End If
    If YPos >= 0 Then
        Printer.CurrentY = YPos - TopMargin
    End If
    
    If Fontibold = True Then
        Printer.FontBold = True
    Else:
        Printer.FontBold = Fontibold
    End If
    If Madhesi > 0 Then
        Printer.Font.Size = Madhesi
    End If
    If Italike = True Then
        Printer.FontItalic = True
    Else:
        Printer.FontItalic = Italike
    End If
    If nenvizuar = True Then
        Printer.FontUnderline = True
    Else
        Printer.FontUnderline = False
    End If
    Printer.Print Tekst
Gabimi:
        If Err.Number = 482 Then
            Gabim = True
        End If
        objErrHandler.KontrolloGabimin Err.Number
End Sub


' Merr si parameter nje numer dhe kthen ekuivalentin e tij ne shkronja
Public Function OrdinalInLetters(numer As Integer) As String
    Dim letter As String
    Select Case numer
    Case 1
        OrdinalInLetters = "parë"
    Case 2
        OrdinalInLetters = "dytë"
    Case 3
        OrdinalInLetters = "tretë"
    Case 4
        OrdinalInLetters = "katërt"
    Case 5
        OrdinalInLetters = "pestë"
    Case 6
        OrdinalInLetters = "gjashtë"
    Case 7
        OrdinalInLetters = "shtatë"
    Case 8
        OrdinalInLetters = "tetë"
    Case 9
        OrdinalInLetters = "nëntë"
    Case 10
        OrdinalInLetters = "dhjetë"
    Case 11
        OrdinalInLetters = "njëmbëdhjetë"
    Case 12
        OrdinalInLetters = "dymbëdhjetë"
    End Select
End Function

Private Function DigitsInLetters(numer As Integer) As String
    Dim letter As String
    Select Case numer
    Case 1
        DigitsInLetters = "një"
    Case 2
        DigitsInLetters = "dy"
    Case 3
        DigitsInLetters = "tre"
    Case 4
        DigitsInLetters = "katër"
    Case 5
        DigitsInLetters = "pesë"
    Case 6
        DigitsInLetters = "gjashtë"
    Case 7
        DigitsInLetters = "shtatë"
    Case 8
        DigitsInLetters = "tetë"
    Case 9
        DigitsInLetters = "nëntë"
    Case 10
        DigitsInLetters = "dhjetë"
    Case 11
        DigitsInLetters = "njëmbëdhjetë"
    Case 12
        DigitsInLetters = "dymbëdhjetë"
    End Select
End Function

' Metode qe percakton orientimin e faqes qe do printohet ne varesi te variablit
' publik Portrait.
Public Sub OrientimFaqe(portret As Boolean)
    Gabim = False
    On Error GoTo Gabimi
    
    If portret = False Then
        Printer.Orientation = vbPRORLandscape
    Else
        Printer.Orientation = vbPRORPortrait
    End If
Gabimi:
    If Err.Number <> 0 Then
        Gabim = True
        objErrHandler.KontrolloGabimin Err.Number
    End If
End Sub

' Funksion qe merr si parameter nje numer dhe kthen ate numer ne
' shkronja te formatuar sipas defteses
Public Function MerrNote(Tekst As String) As String
    Dim numer As Integer
    On Error GoTo Gabimi
        numer = CInt(Tekst)
        MerrNote = Tekst + " (" + DigitsInLetters(numer) + ")"
        Exit Function
Gabimi:
    MerrNote = ""
End Function

' Funksion qe merr si parameter nje numer dhe kthen ate numer ne
' shkronja qe i perkojne numrit ordinal te formatuar sipas defteses
Public Function MerrKlase(Tekst As String) As String
    Dim numer As Integer
    On Error GoTo Gabimi
        numer = CInt(Tekst)
        If numer < 9 Then
            MerrKlase = OrdinalInLetters(numer)
        Else
            numer = numer - 8
            If numer = 0 Then
                MerrKlase = ""
            End If
            MerrKlase = OrdinalInLetters(numer)
        End If
        Exit Function
Gabimi:
        If Err.Number <> 0 Then
            MerrKlase = ""
        End If
End Function

' Funksion qe merr si parameter nje numer dhe kthen ate numer ne
' shkronja qe i perkojne numrit ordinal te formatuar sipas defteses
Public Function MerrKlasePara(Tekst As String) As String
    Dim numer As Integer
    On Error GoTo Gabimi
        numer = CInt(Tekst) - 1
        If numer <= 9 Then
            MerrKlasePara = OrdinalInLetters(numer)
        Else
            numer = numer - 9
            MerrKlasePara = OrdinalInLetters(numer)
        End If
        Exit Function
Gabimi:
    MerrKlasePara = ""
End Function

' Gjeneron nje faqe te re per tu printuar
Public Sub NewPage()
    Printer.NewPage
End Sub

Public Sub EndDoc()
    Printer.EndDoc
End Sub

' Printon nje vije nga startX, startY deri ne endX, endY duke marre parasysh
' edhe marxhinat e printerit
Public Sub PrintLine(StartX As Long, StartY As Long, endX As Long, endY As Long)
    Printer.ScaleMode = 6
    Printer.Line (StartX - LeftMargin, StartY - TopMargin)- _
        (endX - LeftMargin, endY - TopMargin)
End Sub

' Funksion qe kontrollon formatin e fletes qe suporton printer
Public Sub FormatPrinter(lPaperSize As String)
    Gabim = False
    On Error GoTo Gabimi
    If lPaperSize = "vbPRPSA3" Then
        Printer.PaperSize = vbPRPSA3
    ElseIf lPaperSize = "vbPPRSA4" Then
        Printer.PaperSize = vbPRPSA4
    End If
Gabimi:
    If Err.Number <> 0 Then
        Gabim = True
    End If
End Sub

' Ben printimin e nje tabele ne varesi te parametrave qe i jepen
Public Sub PrintTabele(posX As Long, posY As Long, gjeresi As Long, gjatesi As Long, nrRreshta As Integer, nrShtylla As Integer)
    Dim i, j As Integer
    For i = 1 To nrRreshta
        PrintLine posX, posY + ((i - 1) * gjatesi), posX + gjeresi * nrShtylla, posY + ((i - 1) * gjatesi)
    Next
    PrintLine posX, posY + ((i - 1) * gjatesi), posX + gjeresi * nrShtylla, posY + ((i - 1) * gjatesi)
    For i = 1 To nrShtylla
        PrintLine posX + ((i - 1) * gjeresi), posY, posX + ((i - 1) * gjeresi), posY + gjatesi * nrRreshta
    Next
    PrintLine posX + ((i - 1) * gjeresi), posY, posX + ((i - 1) * gjeresi), posY + gjatesi * nrRreshta
End Sub

Public Sub PrintFont(fonti As String)
    If fonti = "Courier New" Then
        Printer.Font.Name = "Courier New"
    ElseIf fonti = "Times New Roman" Then
        Printer.Font.Name = "Times New Roman"
    ElseIf fonti = "Arial" Then
        Printer.FontName = "Arial"
    ElseIf fonti = "Impact" Then
        Printer.FontName = "Impact"
    End If
End Sub

'PRINTPICTURE ****************************************************************
'Kjo procedure printon nje imazh qe ndodhet ne nje Picture
'ose Image
'StartX = pozicioni horizontal i fillimit te printimit
'StartY = pozicioni vertikal i fillimit te printimit
'objPicture = Emri i kontrollit
'Height = gjeresia(opsionale)
'Width =  gjatesia (opsionale)
'Nese keto dy te fundit nuk specifikohen atehere printimi do te
' behet ne dimensionet origjinale
Public Sub PrintPicture(StartX As Single, StartY As Single, objPicture As Object, _
                 Optional Height As Single = 0, Optional Width As Single = 0)
    
   Gabim = False
   Printer.ScaleMode = 6
   On Error GoTo Gabimi
    
    If Height > 0 And Width > 0 Then 'modifiko gjatesine dhe gjeresine
        
        Printer.PaintPicture objPicture.Picture, StartX - LeftMargin, _
                             StartY - TopMargin, Width, Height
    
    ElseIf Height = 0 And Width > 0 Then 'modifiko vetem gjeresine
    
        Printer.PaintPicture objPicture.Picture, StartX - LeftMargin, _
                             StartY - TopMargin, Width
    
    ElseIf Height > 0 And Width = 0 Then 'modifiko vetem gjatesine
        
        Printer.PaintPicture objPicture.Picture, StartX - LeftMargin, _
                             StartY - TopMargin, , Height
    Else ' asnje modifikim, dimensionet origjinale
        
        Printer.PaintPicture objPicture.Picture, StartX - LeftMargin, _
                             StartY - TopMargin
    End If
Exit Sub
Gabimi:
    If Err.Number <> 0 Then
        Gabim = True
    End If
End Sub

Public Sub PrinterColor(ngjyra As String)
    On Error GoTo Gabimi
    If ngjyra = "vbPRCMMonochrome" Then
        Printer.ColorMode = vbPRCMMonochrome
    Else
        Printer.ColorMode = vbPRCMColor
    End If
Gabimi:
End Sub
