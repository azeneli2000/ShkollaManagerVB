VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFormatDeftesatSeri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim Lendet(10) As String
Dim objPrintim As New clsPrintim
Dim objErrHandler As New clsErrorHandler
Public Gabim As Boolean
Private Const CharSjelljeUlet = 206
Private Const CharSjelljeMesme = 373
Public vitiShkollor As String

Public Sub DeftesaMesmeSeriFPa(vekTeDhenash() As String)
    Gabim = False
    If vitiShkollor = "" Then
        vitiShkollor = "2007-2008"
    End If
    If objPrintim.PrinterIsInstalled Then
        On Error GoTo Gabimi
        Printer.FontName = "Times New Roman"
        Printer.PaperSize = vbPRPSA4
        Printer.Orientation = vbPRORPortrait
        Printer.ScaleMode = 6
        ' Printo faqen para te defteses
        objPrintim.PrintLeft 31, 135, "Numri i amzës", False, 12, False
        If objPrintim.Gabim = True Then
            Gabim = True
            Exit Sub
        End If
        objPrintim.PrintLine 57, 140, 72, 140
        objPrintim.PrintLeft 59, 135, vekTeDhenash(3), False, 12, False, False
        objPrintim.PrintLeft 133, 135, "Viti shkollor 2006-2007", False, 12, False, False
        objPrintim.PrintFont "Times New Roman"
        objPrintim.PrintLeft 52, 184, "(PER KLASEN E", False, 12, False, False
        objPrintim.PrintLine 82, 189, 105, 189
        objPrintim.PrintLeft 88, 184, objPrintim.OrdinalInLetters(CInt(vekTeDhenash(10))), False, 12, False, False
        objPrintim.PrintLeft 106, 184, "TE SHKOLLES", False, 12, False, False
        objPrintim.PrintLine 135, 189, 156, 189
        objPrintim.PrintLeft 136, 184, "së mesme", False, 12, False, False
        objPrintim.PrintLeft 157, 184, ")", False, 12, False, False
        objPrintim.PrintLeft 31, 220, "Nxënësi", False, 12, False, False
        objPrintim.PrintLine 46, 225, 178, 225
        objPrintim.PrintLeft 58, 220, vekTeDhenash(0) + " " + vekTeDhenash(2) + " " + vekTeDhenash(1), True, 14, False, False
        objPrintim.PrintLeft 31, 234, "Lindur në", False, 12, False, False
        objPrintim.PrintLine 48, 239, 81, 239
        objPrintim.PrintLeft 51, 234, vekTeDhenash(6), False, 12, False, False
        objPrintim.PrintLeft 82, 234, ", të rrethit", False, 12, False, False
        objPrintim.PrintLine 99, 239, 139, 239
        objPrintim.PrintLeft 101, 234, vekTeDhenash(6), False, 12, False, False
        objPrintim.PrintLeft 140, 234, ", më", False, 12, False, False
        objPrintim.PrintLine 148, 239, 178, 239
        objPrintim.PrintLeft 150, 234, vekTeDhenash(8), False, 12, False, False
        objPrintim.PrintLeft 31, 248, "Rregjistruar për herë të parë, në mbështetje të dëftesës së klasës", False, 12, False, False
        objPrintim.PrintLine 140, 253, 178, 253
        If vekTeDhenash(10) > 1 Then
            objPrintim.PrintLeft 145, 248, "së " + objPrintim.OrdinalInLetters(CInt(vekTeDhenash(10)) - 1)
        End If
    End If
Gabimi:
    If Err.Number <> 0 Then
        objErrHandler.KontrolloGabimin Err.Number
        Gabim = True
    End If
End Sub

Public Sub DeftesaMesmeSeriFPr(tabNotash() As String, vekTeDhenash() As String)
    Gabim = False
    If objPrintim.PrinterIsInstalled Then
        On Error GoTo Gabimi
        ' Printo faqen prapa te defteses
        Dim i, j, k As Integer
        Dim gjatesiVrejte As Integer
        Dim nukEkziston As Boolean
        Printer.PaperSize = vbPRPSA4
        Printer.Orientation = vbPRORPortrait
        Printer.FontName = "Arial"
        Printer.FontSize = 10
        k = 0
        j = 0
        objPrintim.PrintTabele 21, 27, 13, 12, 1, 1
        objPrintim.PrintTabele 34, 27, 45, 12, 1, 1
        objPrintim.PrintTabele 79, 27, 115, 4, 1, 1
        objPrintim.PrintTabele 79, 31, 22, 8, 1, 4
        objPrintim.PrintTabele 167, 31, 27, 8, 1, 1
        objPrintim.PrintTabele 21, 39, 13, 5, 16, 1
        objPrintim.PrintTabele 34, 39, 45, 5, 16, 1
        objPrintim.PrintTabele 79, 39, 22, 5, 16, 4
        objPrintim.PrintTabele 167, 39, 27, 5, 16, 1
        objPrintim.PrintTabele 21, 119, 25, 15, 1, 1
        objPrintim.PrintTabele 46, 119, 33, 5, 3, 1
        objPrintim.PrintTabele 79, 119, 22.5, 5, 3, 3
        objPrintim.PrintLeft 25, 32, "Nr.", False, 10, False, False
        objPrintim.PrintLeft 43, 32, "Lëndët mësimore"
        objPrintim.PrintLeft 125, 27, "Vlerësimet"
        objPrintim.PrintLeft 82, 34, "Semestri I"
        objPrintim.PrintLeft 104, 34, "Semestri II"
        objPrintim.PrintLeft 128, 34, "Vjetore"
        objPrintim.PrintLeft 149, 34, "Riprovime"
        objPrintim.PrintLeft 170, 34, "Përfundimtare"
        ' Bejme printimin e emrave te lendeve dhe notave
        j = 0
        Do While tabNotash(j, 0) <> ""
            objPrintim.PrintLeft 25, 40 + (j * 5), CStr(j + 1), False, 10, False, False
            objPrintim.PrintLeft 36, 40 + (j * 5), tabNotash(j, 0), False, 10, False, False
            objPrintim.PrintLeft 82, 40 + (j * 5), objPrintim.MerrNote(tabNotash(j, 1)), False, 10, False, False
            objPrintim.PrintLeft 104, 40 + (j * 5), objPrintim.MerrNote(tabNotash(j, 2)), False, 10, False, False
            objPrintim.PrintLeft 126, 40 + (j * 5), objPrintim.MerrNote(tabNotash(j, 3)), False, 10, False, False
            objPrintim.PrintLeft 150, 40 + (j * 5), objPrintim.MerrNote(tabNotash(j, 4)), False, 10, False, False
            If tabNotash(j, 4) <> "" Then
                objPrintim.PrintLeft 172, 40 + (j * 5), objPrintim.MerrNote(tabNotash(j, 4)), False, 10, False, False
            Else
                'objPrintim.PrintLeft 172, 40 + (j * 5), objPrintim.MerrNote(tabNotash(j, 3)), False, 10, False, False
            End If
            j = j + 1
        Loop
        objPrintim.PrintLeft 24, 125, "Mungesa"
        objPrintim.PrintLeft 49, 120, "Të arsyeshme"
        objPrintim.PrintLeft 49, 125, "Të paarsyeshme"
        objPrintim.PrintLeft 49, 130, "Gjithsej"
'        objPrintim.PrintLeft 86, 120, CStr(CInt(vekTeDhenash(11)))
'        objPrintim.PrintLeft 108, 120, CStr(CInt(vekTeDhenash(12)))
'        objPrintim.PrintLeft 130, 120, CStr(CInt(vekTeDhenash(11)) + CInt(vekTeDhenash(12)))
'        objPrintim.PrintLeft 86, 125, CStr(CInt(vekTeDhenash(13)))
'        objPrintim.PrintLeft 108, 125, CStr(CInt(vekTeDhenash(14)))
'        objPrintim.PrintLeft 130, 125, CStr(CInt(vekTeDhenash(13)) + CInt(vekTeDhenash(14)))
'        objPrintim.PrintLeft 86, 130, CStr(CInt(vekTeDhenash(11)) + CInt(vekTeDhenash(13)))
'        objPrintim.PrintLeft 108, 130, CStr(CInt(vekTeDhenash(12)) + CInt(vekTeDhenash(14)))
'        objPrintim.PrintLeft 130, 130, CStr(CInt(vekTeDhenash(11)) + CInt(vekTeDhenash(12)) + CInt(vekTeDhenash(13)) + CInt(vekTeDhenash(14)))

'        objPrintim.PrintLeft 285, 157, CStr(CInt(vekTeDhenash(13)) + CInt(vekTeDhenash(14)))
        Do While (j < 16)
            objPrintim.PrintLeft 25, 40 + (j * 5), CStr(j + 1), False, 10, False, False
            j = j + 1
        Loop
        ' Tani bejme printimin e shenimeve.
        ' Jane 8 rreshta shenimesh te cilat do te printohen
        objPrintim.PrintLeft 21, 139, "Shënime për sjelljen dhe formimin e gjithanshëm:", False, 10, False, True
        For i = 0 To 7
            objPrintim.PrintLine 22, 149 + (8 * i), 171, 149 + (8 * i)
        Next
        
        Printer.FontName = "Courier New"
        Dim vrejtje As String
        vrejtje = Replace(CStr(vekTeDhenash(4)), vbCrLf, " ")
        vrejtje = Replace(vrejtje, vbTab, "")
        If Len(vrejtje) < 71 Then
            objPrintim.PrintLeft 23, 145, vrejtje
        Else
            Dim gjatesiVrejtje As Integer
            gjatesiVrejtje = 0
            i = 0
            Do While gjatesiVrejtje < Len(vrejtje)
                If i >= 8 Then
                    Exit Do
                End If
                If Len(vrejtje) > gjatesiVrejtje + 71 Then
                    objPrintim.PrintLeft 23, 145 + (i * 8), Mid(vrejtje, (71 * i) + 1, 71)
                    i = i + 1
                    gjatesiVrejtje = gjatesiVrejtje + 71
                Else
                    objPrintim.PrintLeft 23, 145 + (i * 8), Mid(vrejtje, (71 * i) + 1)
                    gjatesiVrejtje = Len(vrejtje)
                End If
            Loop
        End If
        Printer.FontName = "Arial"
        objPrintim.PrintLeft 21, 215, "Në mbështetje të vlerësimeve të mësipërme, nxënësi"
        objPrintim.PrintLine 105, 219, 171, 219
        objPrintim.PrintLeft 110, 215, vekTeDhenash(0) + " " + vekTeDhenash(2) + " " + vekTeDhenash(1), True
        objPrintim.PrintLeft 21, 220, "kalon klasën e"
        objPrintim.PrintLine 44, 224, 56, 224
        objPrintim.PrintLeft 46, 220, objPrintim.OrdinalInLetters(CInt(vekTeDhenash(10)))
        objPrintim.PrintLeft 21, 230, "Shënim:", False, 10, True, True
        objPrintim.PrintLeft 35, 230, "Nota më e ulët kaluese është 5 (pesë). Nota më e lartë kaluese është 10 (dhjetë).", False, 10, True
        objPrintim.PrintLeft 29, 245, "Mësuesi kujdestar", True, 12, False, False
        objPrintim.PrintLeft 129, 245, "Drejtori i shkollës", True, 12, False, False
        objPrintim.PrintLine 23, 257, 75, 257
        objPrintim.PrintLine 123, 257, 175, 257
        objPrintim.PrintLine 23, 278, 63, 278
        objPrintim.PrintLeft 25, 274, "Vlorë" + ",  më " + CStr(DateTime.Day(DateTime.Now)) + "/" + CStr(DateTime.Month(DateTime.Now)) + "/" + CStr(DateTime.Year(DateTime.Now)), True, 10, True, False
    End If
Gabimi:
    If Err.Number <> 0 Then
        objErrHandler.KontrolloGabimin Err.Number
        Gabim = True
    End If
End Sub

Public Sub DeftesaUletSeriFPa(vekTeDhenash() As String)
    Gabim = False
    If objPrintim.PrinterIsInstalled Then
        On Error GoTo Gabimi
        Printer.FontName = "Times New Roman"
        Printer.PaperSize = vbPRPSA4
        Printer.Orientation = vbPRORPortrait
        Printer.ScaleMode = 6
        ' Printo faqen para te defteses
        objPrintim.PrintLeft 31, 135, "Numri i amzës", False, 12, False
        If objPrintim.Gabim = True Then
            Gabim = True
            Exit Sub
        End If
        If vitiShkollor = "" Then
            vitiShkollor = "2007-2008"
        End If
        objPrintim.PrintLine 57, 140, 72, 140
        objPrintim.PrintLeft 59, 135, vekTeDhenash(3), False, 12, False, False
        objPrintim.PrintLeft 133, 135, "Viti shkollor " + vitiShkollor, False, 12, False, False
        objPrintim.PrintFont "Times New Roman"
        objPrintim.PrintLeft 52, 184, "(PER KLASEN E", False, 12, False, False
        objPrintim.PrintLine 82, 189, 105, 189
        objPrintim.PrintLeft 88, 184, objPrintim.OrdinalInLetters(CInt(vekTeDhenash(10))), False, 12, False, False
        objPrintim.PrintLeft 106, 184, "TE SHKOLLES", False, 12, False, False
        objPrintim.PrintLine 135, 189, 156, 189
        objPrintim.PrintLeft 136, 184, "9-vjeçare", False, 12, False, False
        objPrintim.PrintLeft 157, 184, ")", False, 12, False, False
        objPrintim.PrintLeft 31, 220, "Nxënësi", False, 12, False, False
        objPrintim.PrintLine 46, 225, 178, 225
        objPrintim.PrintLeft 58, 220, vekTeDhenash(0) + " " + vekTeDhenash(2) + " " + vekTeDhenash(1), True, 14, False, False
        objPrintim.PrintLeft 31, 234, "Lindur në", False, 12, False, False
        objPrintim.PrintLine 48, 239, 81, 239
        objPrintim.PrintLeft 51, 234, vekTeDhenash(6), False, 12, False, False
        objPrintim.PrintLeft 82, 234, ", të rrethit", False, 12, False, False
        objPrintim.PrintLine 99, 239, 139, 239
        objPrintim.PrintLeft 101, 234, vekTeDhenash(6), False, 12, False, False
        objPrintim.PrintLeft 140, 234, ", më", False, 12, False, False
        objPrintim.PrintLine 148, 239, 178, 239
        objPrintim.PrintLeft 150, 234, vekTeDhenash(8), False, 12, False, False
        objPrintim.PrintLeft 31, 248, "Rregjistruar për herë të parë, në mbështetje të dëftesës së klasës", False, 12, False, False
        objPrintim.PrintLine 140, 253, 178, 253
        If vekTeDhenash(10) > 1 Then
            objPrintim.PrintLeft 145, 248, "së " + objPrintim.OrdinalInLetters(CInt(vekTeDhenash(10)) - 1)
        End If
    End If
Gabimi:
    If Err.Number <> 0 Then
        objErrHandler.KontrolloGabimin Err.Number
        Gabim = True
    End If
End Sub

Public Sub DeftesaUletSeriFPr(tabNotash() As String, vekTeDhenash() As String)
    Gabim = False
    If objPrintim.PrinterIsInstalled Then
        On Error GoTo Gabimi
        ' Printo faqen prapa te defteses
        Dim i, j, k As Integer
        Dim gjatesiVrejte As Integer
        Dim nukEkziston As Boolean
        Printer.PaperSize = vbPRPSA4
        Printer.Orientation = vbPRORPortrait
        Printer.FontName = "Arial"
        Printer.FontSize = 10
        k = 0
        j = 0
        objPrintim.PrintTabele 21, 27, 13, 12, 1, 1
        objPrintim.PrintTabele 34, 27, 45, 12, 1, 1
        objPrintim.PrintTabele 79, 27, 115, 4, 1, 1
        objPrintim.PrintTabele 79, 31, 22, 8, 1, 4
        objPrintim.PrintTabele 167, 31, 27, 8, 1, 1
        objPrintim.PrintTabele 21, 39, 13, 5, 16, 1
        objPrintim.PrintTabele 34, 39, 45, 5, 16, 1
        objPrintim.PrintTabele 79, 39, 22, 5, 16, 4
        objPrintim.PrintTabele 167, 39, 27, 5, 16, 1
        objPrintim.PrintTabele 21, 119, 25, 15, 1, 1
        objPrintim.PrintTabele 46, 119, 33, 5, 3, 1
        objPrintim.PrintTabele 79, 119, 22.5, 5, 3, 3
        objPrintim.PrintLeft 25, 32, "Nr.", False, 10, False, False
        objPrintim.PrintLeft 43, 32, "Lëndët mësimore"
        objPrintim.PrintLeft 125, 27, "Vlerësimet"
        objPrintim.PrintLeft 82, 34, "Semestri I"
        objPrintim.PrintLeft 104, 34, "Semestri II"
        objPrintim.PrintLeft 128, 34, "Vjetore"
        objPrintim.PrintLeft 149, 34, "Riprovime"
        objPrintim.PrintLeft 170, 34, "Përfundimtare"
        ' Bejme printimin e emrave te lendeve dhe notave
        j = 0
        Do While tabNotash(j, 0) <> ""
            objPrintim.PrintLeft 25, 40 + (j * 5), CStr(j + 1), False, 10, False, False
            objPrintim.PrintLeft 36, 40 + (j * 5), tabNotash(j, 0), False, 10, False, False
            objPrintim.PrintLeft 82, 40 + (j * 5), objPrintim.MerrNote(tabNotash(j, 1)), False, 10, False, False
            objPrintim.PrintLeft 104, 40 + (j * 5), objPrintim.MerrNote(tabNotash(j, 2)), False, 10, False, False
            objPrintim.PrintLeft 126, 40 + (j * 5), objPrintim.MerrNote(tabNotash(j, 3)), False, 10, False, False
            objPrintim.PrintLeft 150, 40 + (j * 5), objPrintim.MerrNote(tabNotash(j, 4)), False, 10, False, False
            If tabNotash(j, 4) <> "" Then
                objPrintim.PrintLeft 172, 40 + (j * 5), objPrintim.MerrNote(tabNotash(j, 4)), False, 10, False, False
            Else
                'objPrintim.PrintLeft 172, 40 + (j * 5), objPrintim.MerrNote(tabNotash(j, 3)), False, 10, False, False
            End If
            j = j + 1
        Loop
        objPrintim.PrintLeft 24, 125, "Mungesa"
        objPrintim.PrintLeft 49, 120, "Të arsyeshme"
        objPrintim.PrintLeft 49, 125, "Të paarsyeshme"
        objPrintim.PrintLeft 49, 130, "Gjithsej"
'        objPrintim.PrintLeft 86, 120, CStr(CInt(vekTeDhenash(11)))
'        objPrintim.PrintLeft 108, 120, CStr(CInt(vekTeDhenash(12)))
'        objPrintim.PrintLeft 130, 120, CStr(CInt(vekTeDhenash(11)) + CInt(vekTeDhenash(12)))
'        objPrintim.PrintLeft 86, 125, CStr(CInt(vekTeDhenash(13)))
'        objPrintim.PrintLeft 108, 125, CStr(CInt(vekTeDhenash(14)))
'        objPrintim.PrintLeft 130, 125, CStr(CInt(vekTeDhenash(13)) + CInt(vekTeDhenash(14)))
'        objPrintim.PrintLeft 86, 130, CStr(CInt(vekTeDhenash(11)) + CInt(vekTeDhenash(13)))
'        objPrintim.PrintLeft 108, 130, CStr(CInt(vekTeDhenash(12)) + CInt(vekTeDhenash(14)))
'        objPrintim.PrintLeft 130, 130, CStr(CInt(vekTeDhenash(11)) + CInt(vekTeDhenash(12)) + CInt(vekTeDhenash(13)) + CInt(vekTeDhenash(14)))

'        objPrintim.PrintLeft 285, 157, CStr(CInt(vekTeDhenash(13)) + CInt(vekTeDhenash(14)))
        Do While (j < 16)
            objPrintim.PrintLeft 25, 40 + (j * 5), CStr(j + 1), False, 10, False, False
            j = j + 1
        Loop
        ' Tani bejme printimin e shenimeve.
        ' Jane 8 rreshta shenimesh te cilat do te printohen
        objPrintim.PrintLeft 21, 139, "Shënime për sjelljen dhe formimin e gjithanshëm:", False, 10, False, True
        For i = 0 To 7
            objPrintim.PrintLine 22, 149 + (8 * i), 171, 149 + (8 * i)
        Next
        Printer.FontName = "Courier New"
        Dim vrejtje As String
        vrejtje = CStr(vekTeDhenash(4))
        If Len(vrejtje) < 71 Then
            objPrintim.PrintLeft 23, 145, vekTeDhenash(4)
        Else
            Dim gjatesiVrejtje As Integer
            gjatesiVrejtje = 0
            i = 0
            Do While gjatesiVrejtje < Len(vrejtje)
                If i >= 8 Then
                    Exit Do
                End If
                If Len(vrejtje) > gjatesiVrejtje + 71 Then
                    objPrintim.PrintLeft 23, 145 + (i * 8), Mid(vrejtje, (71 * i) + 1, 71)
                    i = i + 1
                    gjatesiVrejtje = gjatesiVrejtje + 71
                Else
                    objPrintim.PrintLeft 23, 145 + (i * 8), Mid(vrejtje, (71 * i) + 1)
                    gjatesiVrejtje = Len(vrejtje)
                End If
            Loop
        End If
        Printer.FontName = "Arial"
        objPrintim.PrintLeft 21, 215, "Në mbështetje të vlerësimeve të mësipërme, nxënësi"
        objPrintim.PrintLine 105, 219, 171, 219
        objPrintim.PrintLeft 110, 215, vekTeDhenash(0) + " " + vekTeDhenash(2) + " " + vekTeDhenash(1), True
        objPrintim.PrintLeft 21, 220, "kalon klasën e"
        objPrintim.PrintLine 44, 224, 56, 224
        objPrintim.PrintLeft 46, 220, objPrintim.OrdinalInLetters(CInt(vekTeDhenash(10)))
        objPrintim.PrintLeft 21, 230, "Shënim:", False, 10, True, True
        objPrintim.PrintLeft 35, 230, "Nota më e ulët kaluese është 5 (pesë). Nota më e lartë kaluese është 10 (dhjetë).", False, 10, True
        objPrintim.PrintLeft 29, 245, "Mësuesi kujdestar", True, 12, False, False
        objPrintim.PrintLeft 129, 245, "Drejtori i shkollës", True, 12, False, False
        objPrintim.PrintLine 23, 257, 75, 257
        objPrintim.PrintLine 123, 257, 175, 257
        objPrintim.PrintLine 23, 278, 63, 278
        objPrintim.PrintLeft 25, 274, "Vlorë" + ",  më " + CStr(DateTime.Day(DateTime.Now)) + "/" + CStr(DateTime.Month(DateTime.Now)) + "/" + CStr(DateTime.Year(DateTime.Now)), True, 10, True, False
    End If
Gabimi:
    If Err.Number <> 0 Then
        objErrHandler.KontrolloGabimin Err.Number
        Gabim = True
    End If
End Sub

Public Sub NewPage()
    objPrintim.NewPage
End Sub

Public Sub EndDoc()
    objPrintim.EndDoc
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

' Metode qe percakton orientimin e faqes qe do printohet ne varesi te variablit
' publik Portrait.
Public Sub OrientimFaqe(portrait As Boolean)
    Gabim = False
    On Error GoTo Gabimi
    If portrait = False Then
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

' Metode qe vendos madhesine fletes qe do te printohet
Public Sub PaperSizeA4(madhesiA4 As Boolean)
    Gabim = False
    On Error GoTo Gabimi
    If madhesiA4 = False Then
        Printer.PaperSize = vbPRPSA3
    Else
        Printer.PaperSize = vbPRPSA4
    End If
Gabimi:
    If Err.Number <> 0 Then
        Gabim = True
        objErrHandler.KontrolloGabimin Err.Number
    End If
End Sub

Public Function PrinterIsInstalled() As Boolean
    PrinterIsInstalled = objPrintim.PrinterIsInstalled
End Function
