Attribute VB_Name = "modExcel"
Public Sub KonvertoExcel()
Dim frm As String
If (active_form Is Nothing) Then
    Exit Sub
End If
frm = active_form.Name
Dim Klasa, indeksi, vitishkollor, lloji As String
Dim txt As String
Dim cikli As Integer
txt = ""
Select Case frm
    Case "frmKonsultimeEvidenca"
        If (active_form.emrat.Visible = True) Then
            Klasa = active_form.cboKlasa.Text
            indeksi = active_form.cboIndeksi.Text
            vitishkollor = active_form.cboVitiShkollor.Text
            If (active_form.optSemestri1.Value) Then
                txt = "EVIDENCA E NOTAVE PER SEMESTRIN E PARE. KLASA " + Klasa + " " + indeksi + ", VITI SHKOLLOR " + vitishkollor
                lloji = "S1"
            ShkruajEvidencaEksel txt, active_form.emrat, active_form.Lendet, lloji, active_form.notat
            ElseIf (active_form.optSemestri2.Value) Then
                txt = "EVIDENCA E NOTAVE PER SEMESTRIN E DYTE. KLASA " + Klasa + " " + indeksi + ", VITI SHKOLLOR " + vitishkollor
                lloji = "S2"
            ShkruajEvidencaEksel txt, active_form.emrat, active_form.Lendet, lloji, active_form.notat
            ElseIf (active_form.optVjetore.Value) Then
                txt = "EVIDENCA E NOTAVE VJETORE. KLASA " + Klasa + " " + indeksi + ", VITI SHKOLLOR " + vitishkollor
                ShkruajEvidenceVjetoreEksel txt, active_form.emrat, active_form.Lendet, lloji, active_form.notat
            End If
        End If
    Case "frmKonsultimeAmza"
        If (active_form.gridaAmza.Visible = True) Then
            txt = "Amza për nxënësin " + active_form.txtEmri.Text + " " + active_form.txtAtesia.Text + " " + active_form.txtMbiemri.Text & _
            " me numër amze " + active_form.txtAmzaNo.Text
            If (active_form.optUlet.Value) Then
                cikli = 0
            Else
                cikli = 1
            End If
            ShkruajAmzaNxenesiEksel txt, active_form.gridaLabel, active_form.gridaKlasat, _
            active_form.gridaLendet, active_form.gridaAmza, active_form.gridaPjekurie, cikli
        End If
        
    Case "frmKonsultimeNota"
        If (active_form.Lendet.Visible = True) Then
            txt = "Notat "
            If (active_form.optSemestri1.Value) Then
                txt = txt + " për semestrin e parë"
                lloji = "1"
            ElseIf (active_form.optSemestri2.Value) Then
                txt = txt + " për semestrin e dytë"
                lloji = "2"
            Else
                txt = txt + " vjetore"
                lloji = "3"
            End If
            ShkruajNotaEksel txt, active_form.Lendet, active_form.Cgrid1, lloji
        End If
    Case "frmStatistikaKlasa"
        Dim currRow, currColumn As Integer
        currRow = active_form.gridaKlasat.row
        currColumn = active_form.gridaKlasat.col
        If (active_form.gridaKlasat.Rows > 0) Then
            active_form.gridaKlasat.row = 0
            active_form.gridaKlasat.col = 0
            If (active_form.gridaKlasat.Text <> "") Then
                active_form.gridaKlasat.row = currRow
                active_form.gridaKlasat.col = currColumn
                ShkruajStatistikaEkselKlasat active_form.gridaLabelKlasat, active_form.gridaKlasat, active_form.gridaLendetLabel, active_form.gridaLendet
            End If
        End If
    Case "frmStatistikaNxenes"
        currRow = active_form.gridaKlasat.row
        currColumn = active_form.gridaKlasat.col
        If (active_form.gridaKlasat.Rows > 0) Then
            active_form.gridaKlasat.row = 0
            active_form.gridaKlasat.col = 0
            If (active_form.gridaKlasat.Text <> "") Then
                active_form.gridaKlasat.row = currRow
                active_form.gridaKlasat.col = currColumn
                ShkruajStatistikaEkselNxenesit active_form.gridaLabelKlasat, active_form.gridaKlasat, active_form.gridaNxenesitLabel, active_form.gridaNxenesit
            End If
        End If
    Case "frmMesataretMomentale"
        If (active_form.amzaEmerMbiemer.Visible = True) Then
            If active_form.optSemestri1.Value Then
                txt = "Mesataret momentale sipas klasave për semestrin e parë"
            Else
                txt = "Mesataret momentale sipas klasave për semestrin e dytë"
            End If
            ShkruajMesataretMomentaleKlasaEksel active_form.amzaEmerMbiemer, active_form.emrat, _
            active_form.Lendet, active_form.notat, txt
        End If
    Case "frmMesataretCikli"
        If (active_form.Header1.Visible = True) Then
            If active_form.optSemestri1.Value Then
                txt = "Mesataret momentale sipas cikleve për semestrin e parë"
            Else
                txt = "Mesataret momentale sipas cikleve për semestrin e dytë"
            End If
            ShkruajMesataretMomentaleCikleEksel active_form.Header1, active_form.Mesataret1, _
            active_form.Header2, active_form.Mesataret2, active_form.Header3, _
            active_form.Mesataret3, txt
        End If
    Case "frmNxenesitDalluar"
        If (active_form.Nxenesit.Visible = True) Then
            If (active_form.optMesme.Value) Then
                txt = "Nxënësit e dalluar për ciklin e mesëm"
            Else
                txt = "Nxënësit e dalluar për ciklin nëntëvjeçar"
            End If
            ShkruajNxenesitDalluarEksel active_form.Header, active_form.Nxenesit, txt
        End If
End Select
End Sub

Private Sub ShkruajNxenesitDalluarEksel(Header As MSHFlexGrid, Nxenesit As MSHFlexGrid, txt As String)
    Dim FileExcel        As Workbook
    Dim FaqeExcel      As Worksheet
    Dim fso As New FileSystemObject
    If Not DirExists("C:\RaporteExcel") Then
        Dim objFile As New FileSystemObject
        objFile.CreateFolder ("C:\RaporteExcel")
    End If
    On Error GoTo Gabim:
    fso.CopyFile App.Path + "\Excel\nxenesitDalluar.xls", "C:\RaporteExcel\nxenesitDalluar.xls", True
    
    Set fso = Nothing
    
    Set FileExcel = Excel.Workbooks.Open("C:\RaporteExcel\nxenesitDalluar.xls")
    Set FaqeExcel = FileExcel.Worksheets("Nxënësit e dalluar")
    
    FaqeExcel.Cells(1, 2) = emerShkolla
    FaqeExcel.Cells(2, 2) = txt + ", Viti shkollor " + active_form.cboVitiShkollor.Text
    
    ShkruajGride Header, 4, 2, FaqeExcel
    ShkruajGride Nxenesit, 5, 2, FaqeExcel
    
    FileExcel.Close (True)
    MsgBox "Nxënësit e dalluar u ruajtën në C:\RaporteExcel\nxenesitDalluar.xls.", vbOKOnly, "Konvertimi në Excel"
    Exit Sub
Gabim:
    MsgBox "Një skedar me emër nxenesitDalluar.xls është i hapur." + Chr(10) + "" & _
    "Mbylleni skedarin para se të bëni konvertimin", vbExclamation, "Konvertimi në Excel"
End Sub

Private Sub ShkruajMesataretMomentaleCikleEksel(Header1 As MSHFlexGrid, Mesataret1 As MSHFlexGrid, _
            Header2 As MSHFlexGrid, Mesataret2 As MSHFlexGrid, Header3 As MSHFlexGrid, _
            Mesataret3 As MSHFlexGrid, txt As String)
    Dim FileExcel        As Workbook
    Dim FaqeExcel      As Worksheet
    Dim fso As New FileSystemObject
    If Not DirExists("C:\RaporteExcel") Then
        Dim objFile As New FileSystemObject
        objFile.CreateFolder ("C:\RaporteExcel")
    End If
    On Error GoTo Gabim:
    fso.CopyFile App.Path + "\Excel\mesataretMomentaleCiklet.xls", "C:\RaporteExcel\mesataretMomentaleCiklet.xls", True
    Set fso = Nothing
    
    Set FileExcel = Excel.Workbooks.Open("C:\RaporteExcel\mesataretMomentaleCiklet.xls")
    Set FaqeExcel = FileExcel.Worksheets("Mesataret momentale ciklet")
    
    FaqeExcel.Cells(1, 2) = emerShkolla
    FaqeExcel.Cells(2, 2) = txt
    FaqeExcel.Cells(3, 2) = "Viti shkollor " + active_form.cboVitiShkollor.Text
    
    FaqeExcel.Cells(5, 2) = active_form.lblUlet.Caption
    FaqeExcel.Cells(5, 5) = active_form.lblTetevjecare.Caption
    FaqeExcel.Cells(5, 8) = active_form.lblMesme.Caption
    
    ShkruajGride Header1, 6, 2, FaqeExcel
    ShkruajGride Header2, 6, 5, FaqeExcel
    ShkruajGride Header3, 6, 8, FaqeExcel
    
    ShkruajGride Mesataret1, 7, 2, FaqeExcel
    ShkruajGride Mesataret2, 7, 5, FaqeExcel
    ShkruajGride Mesataret3, 7, 8, FaqeExcel
    
    FaqeExcel.Range("B" + CStr(Mesataret1.Rows + 6), "C" + CStr(Mesataret1.Rows + 6)).Font.Bold = True
    FaqeExcel.Range("E" + CStr(Mesataret2.Rows + 6), "F" + CStr(Mesataret2.Rows + 6)).Font.Bold = True
    FaqeExcel.Range("H" + CStr(Mesataret3.Rows + 6), "I" + CStr(Mesataret3.Rows + 6)).Font.Bold = True
    
    FileExcel.Close (True)
    MsgBox "Mesataret momentale sipas cikleve u ruajtën në C:\RaporteExcel\mesataretMomentaleCiklet.xls.", vbOKOnly, "Konvertimi në Excel"
    Exit Sub
Gabim:
    MsgBox "Një skedar me emër mesataretMomentaleCiklet.xls është i hapur." + Chr(10) + "" & _
    "Mbylleni skedarin para se të bëni konvertimin", vbExclamation, "Konvertimi në Excel"
End Sub

Private Sub ShkruajMesataretMomentaleKlasaEksel(amzaEmerMbiemer As MSHFlexGrid, _
emrat As MSHFlexGrid, Lendet As MSHFlexGrid, notat As MSHFlexGrid, txt As String)
    Dim FileExcel        As Workbook
    Dim FaqeExcel      As Worksheet
    Dim fso As New FileSystemObject
    If Not DirExists("C:\RaporteExcel") Then
        Dim objFile As New FileSystemObject
        objFile.CreateFolder ("C:\RaporteExcel")
    End If
    On Error GoTo Gabim:
    fso.CopyFile App.Path + "\Excel\mesataretMomentaleKlasa.xls", "C:\RaporteExcel\mesataretMomentaleKlasa.xls", True
    Set fso = Nothing
    
    Set FileExcel = Excel.Workbooks.Open("C:\RaporteExcel\mesataretMomentaleKlasa.xls")
    Set FaqeExcel = FileExcel.Worksheets("Mesataret momentale klasa")
    
    FaqeExcel.Cells(1, 2) = emerShkolla
    FaqeExcel.Cells(2, 2) = txt
    FaqeExcel.Cells(3, 2) = "Klasa " + active_form.cboKlasa.Text + "-" + _
        active_form.cboIndeksi.Text + ", Viti shkollor " + active_form.cboVitiShkollor.Text
        
    ShkruajGride amzaEmerMbiemer, 5, 2, FaqeExcel
    
    ShkruajGride emrat, 6, 2, FaqeExcel
    
    ShkruajGride Lendet, 5, 5, FaqeExcel
    
    ShkruajGride notat, 6, 5, FaqeExcel
    
    Dim cell1, cell2, cell3 As String
    cell1 = GetRange(5 + emrat.Rows, 2)
    cell2 = GetRange(5 + emrat.Rows, 4 + Lendet.Cols)
    cell3 = GetRange(5, 4 + Lendet.Cols)
    FaqeExcel.Range(cell1, cell2).Font.Bold = True
    FaqeExcel.Range(cell3, cell2).Font.Bold = True
    
    FileExcel.Close (True)
    MsgBox "Mesataret momentale sipas klasave u ruajtën në C:\RaporteExcel\mesataretMomentaleKlasa.xls.", vbOKOnly, "Konvertimi në Excel"
    Exit Sub
Gabim:
    MsgBox "Një skedar me emër mesataretMomentaleKlasa.xls është i hapur." + Chr(10) + "" & _
    "Mbylleni skedarin para se të bëni konvertimin", vbExclamation, "Konvertimi në Excel"

End Sub

Private Sub ShkruajStatistikaEkselNxenesit(gridaLabelKlasat As MSHFlexGrid _
, gridaKlasat As MSHFlexGrid, gridaNxenesitLabel As MSHFlexGrid, gridaNxenesit As MSHFlexGrid)
    Dim FileExcel        As Workbook
    Dim FaqeExcel      As Worksheet
    Dim fso As New FileSystemObject
    If Not DirExists("C:\RaporteExcel") Then
        Dim objFile As New FileSystemObject
        objFile.CreateFolder ("C:\RaporteExcel")
    End If
    On Error GoTo Gabim:
    fso.CopyFile App.Path + "\Excel\statistikaNxenesit.xls", "C:\RaporteExcel\statistikaNxenesit.xls", True
    Set fso = Nothing
    
    Set FileExcel = Excel.Workbooks.Open("C:\RaporteExcel\statistikaNxenesit.xls")
    Set FaqeExcel = FileExcel.Worksheets("Statistika për nxënësit")
    
    FaqeExcel.Cells(1, 2) = emerShkolla
    FaqeExcel.Cells(2, 2) = "Statistika për nxënësit"
    FaqeExcel.Cells(3, 2) = active_form.cboOptions.Text
    FaqeExcel.Cells(4, 2) = active_form.cboVitiShkollor.Text
    
    FaqeExcel.Cells(6, 2) = "Raporti meshkuj-femra sipas klasave"
    
    ShkruajGride gridaLabelKlasat, 7, 2, FaqeExcel
    
    ShkruajGride gridaKlasat, 8, 2, FaqeExcel
    If gridaNxenesit.Rows > 0 Then
        gridaNxenesit.row = 0
        gridaNxenesit.col = 0
        'gridaNxenesitLabel = 0
        If (gridaKlasat.row > 0 And gridaNxenesit.Visible = True And gridaNxenesit.Text <> "") Then
            gridaKlasat.col = 0
            FaqeExcel.Cells(6, 6) = "Mesataret e nxënësve për klasën " + active_form.txtKlasaZgjedhur.Text
            ShkruajGride gridaNxenesitLabel, 7, 6, FaqeExcel
            'FaqeExcel.Range("B" + CStr(9 + gridaKlasat.Rows), "D" + CStr(10 + gridaKlasat.Rows)).Font.Bold = True
            ShkruajGride gridaNxenesit, 8, 6, FaqeExcel
        End If
    End If
    FileExcel.Close (True)
    MsgBox "Statistikat sipas klasave u ruajtën në C:\RaporteExcel\statistikaNxenesit.xls.", vbOKOnly, "Konvertimi në Excel"
    Exit Sub
Gabim:
    MsgBox "Një skedar me emër statistikaNxenesit.xls është i hapur." + Chr(10) + "" & _
    "Mbylleni skedarin para se të bëni konvertimin", vbExclamation, "Konvertimi në Excel"
End Sub

Private Sub ShkruajStatistikaEkselKlasat(gridaLabelKlasat As MSHFlexGrid _
, gridaKlasat As MSHFlexGrid, gridaLendetLabel As MSHFlexGrid, gridaLendet As MSHFlexGrid)
    Dim FileExcel        As Workbook
    Dim FaqeExcel      As Worksheet
    Dim fso As New FileSystemObject
    If Not DirExists("C:\RaporteExcel") Then
        Dim objFile As New FileSystemObject
        objFile.CreateFolder ("C:\RaporteExcel")
    End If
    On Error GoTo Gabim:
    fso.CopyFile App.Path + "\Excel\statistikaKlasat.xls", "C:\RaporteExcel\statistikaKlasat.xls", True
    Set fso = Nothing
    
    Set FileExcel = Excel.Workbooks.Open("C:\RaporteExcel\statistikaKlasat.xls")
    Set FaqeExcel = FileExcel.Worksheets("Statistika sipas klasave")
    
    FaqeExcel.Cells(1, 2) = emerShkolla
    FaqeExcel.Cells(2, 2) = "Statistika sipas klasave"
    FaqeExcel.Cells(3, 2) = active_form.cboOptions.Text
    FaqeExcel.Cells(4, 2) = active_form.cboVitiShkollor.Text
    
    FaqeExcel.Cells(6, 2) = "Mungesat dhe mesataret për të gjitha klasat"
    
    ShkruajGride gridaLabelKlasat, 7, 2, FaqeExcel
    
    ShkruajGride gridaKlasat, 8, 2, FaqeExcel
    If gridaLendet.Rows > 0 Then
        gridaLendet.row = 0
        gridaLendet.col = 0
        'gridaLendetLabel = 0
        If (gridaKlasat.row > 0 And gridaLendet.Visible = True And gridaLendet.Text <> "") Then
            gridaKlasat.col = 0
            FaqeExcel.Cells(9 + gridaKlasat.Rows, 2) = "Mesataret sipas lëndëve për klasën " + active_form.txtKlasaZgjedhur.Text
            ShkruajGride gridaLendetLabel, 10 + gridaKlasat.Rows, 2, FaqeExcel
            FaqeExcel.Range("B" + CStr(9 + gridaKlasat.Rows), "D" + CStr(10 + gridaKlasat.Rows)).Font.Bold = True
            ShkruajGride gridaLendet, 11 + gridaKlasat.Rows, 2, FaqeExcel
        End If
    End If
    FileExcel.Close (True)
    MsgBox "Statistikat sipas klasave u ruajtën në C:\RaporteExcel\statistikaKlasat.xls.", vbOKOnly, "Konvertimi në Excel"
    Exit Sub
Gabim:
    MsgBox "Një skedar me emër statistikaKlasat.xls është i hapur." + Chr(10) + "" & _
    "Mbylleni skedarin para se të bëni konvertimin", vbExclamation, "Konvertimi në Excel"
End Sub

Private Sub ShkruajNotaEksel(txt As String, Lendet As Cgrid, Cgrid1 As Cgrid, lloji As String)
    Dim FileExcel        As Workbook
    Dim FaqeExcel      As Worksheet
    Dim fso As New FileSystemObject
    If Not DirExists("C:\RaporteExcel") Then
        Dim objFile As New FileSystemObject
        objFile.CreateFolder ("C:\RaporteExcel")
    End If
    On Error GoTo Gabim:
    fso.CopyFile App.Path + "\Excel\konsultimeNota.xls", "C:\RaporteExcel\konsultimeNota.xls", True
    Set fso = Nothing
    
    Set FileExcel = Excel.Workbooks.Open("C:\RaporteExcel\konsultimeNota.xls")
    Set FaqeExcel = FileExcel.Worksheets("Notat")
    
    FaqeExcel.Cells(1, 2) = emerShkolla
    FaqeExcel.Cells(2, 2) = txt
    
    FaqeExcel.Cells(3, 2) = "Amza"
    FaqeExcel.Cells(4, 2) = "Nxënësi"
    FaqeExcel.Cells(5, 2) = "Klasa"
    FaqeExcel.Cells(6, 2) = "Viti shkollor"
    
    FaqeExcel.Cells(3, 3) = active_form.txtAmzaNo.Text
    FaqeExcel.Cells(4, 3) = active_form.txtEmri.Text + " " + active_form.txtAtesia.Text + " " + active_form.txtMbiemri.Text
    FaqeExcel.Cells(5, 3) = active_form.txtKlasa.Text + "-" + active_form.txtIndeksi.Text + " "
    FaqeExcel.Cells(6, 3) = active_form.cboVitiShkollor.Text

    FaqeExcel.Cells(8, 2) = "Lëndët"
    If lloji = "1" Or lloji = "2" Then
        FaqeExcel.Cells(8, 3) = "Notat dhe mungesat"
    Else
        FaqeExcel.Cells(8, 3) = "S1"
        FaqeExcel.Cells(8, 4) = "S2"
        FaqeExcel.Cells(8, 5) = "V"
        FaqeExcel.Cells(8, 6) = "R"
        FaqeExcel.Range("C8", "F8").HorizontalAlignment = -4108
    End If
    
    'shkruaj Lendet
    ShkruajGride2 Lendet, 9, 2, FaqeExcel, Lendet.Width, 255
    
    'shkruaj notat
    ShkruajGride2 Cgrid1, 9, 3, FaqeExcel, 300, 255
    
    FileExcel.Close (True)
    MsgBox "Notat për nxënësin u ruajtën në C:\RaporteExcel\konsultimeNota.xls.", vbOKOnly, "Konvertimi në Excel"
    Exit Sub
Gabim:
    MsgBox "Një skedar me emër konsultimeNota.xls është i hapur." + Chr(10) + "" & _
    "Mbylleni skedarin para se të bëni konvertimin", vbExclamation, "Konvertimi në Excel"

End Sub

Private Sub ShkruajAmzaNxenesiEksel(txt As String, gridaLabel As MSHFlexGrid, gridaKlasat As MSHFlexGrid _
, gridaLendet As MSHFlexGrid, gridaAmza As MSHFlexGrid, gridaPjekurie As MSFlexGrid, cikli As Integer)

    Dim FileExcel        As Workbook
    Dim FaqeExcel      As Worksheet
    Dim fso As New FileSystemObject
    If Not DirExists("C:\RaporteExcel") Then
        Dim objFile As New FileSystemObject
        objFile.CreateFolder ("C:\RaporteExcel")
    End If
    On Error GoTo Gabim:
    fso.CopyFile App.Path + "\Excel\amzaNxenesi.xls", "C:\RaporteExcel\amzaNxenesi.xls", True
    Set fso = Nothing
    
    Set FileExcel = Excel.Workbooks.Open("C:\RaporteExcel\amzaNxenesi.xls")
    Set FaqeExcel = FileExcel.Worksheets("Amza për nxënësin")
    
    FaqeExcel.Cells(1, 2) = emerShkolla
    FaqeExcel.Cells(2, 2) = txt
    
    FaqeExcel.Cells(4, 2) = "Klasa"
    FaqeExcel.Cells(4, 3) = "Viti shkollor"
    
    'shkruaj klasat dhe vitet shkollor
    ShkruajGride gridaKlasat, 5, 2, FaqeExcel
    
    'shkruaj lendet
    ShkruajGride gridaLendet, 4, 4, FaqeExcel
    
    ShkruajGride gridaAmza, 5, 4, FaqeExcel
    
    If (gridaPjekurie.Visible = True) Then
        If cikli = 0 Then
            FaqeExcel.Cells(23, 2) = "Notat në provimet e lirimit"
            ShkruajGride1 gridaPjekurie, 23, 4, FaqeExcel
        Else
            aqeExcel.Cells(13, 2) = "Notat në provimet e maturës"
            ShkruajGride1 gridaPjekurie, 13, 4, FaqeExcel
        End If
    End If
        
    FileExcel.Close (True)
    MsgBox "Amza për nxënësin u ruajt në C:\RaporteExcel\amzaNxenesi.xls.", vbOKOnly, "Konvertimi në Excel"

    Exit Sub
Gabim:
    MsgBox "Një skedar me emër amzaNxenesi.xls është i hapur." + Chr(10) + "" & _
    "Mbylleni skedarin para se të bëni konvertimin", vbExclamation, "Konvertimi në Excel"
End Sub

Private Sub ShkruajGride(grid As MSHFlexGrid, top As Integer, left As Integer, FaqeExcel As Worksheet)
    Dim i, j As Integer
    For i = 0 To grid.Rows - 1
        grid.row = i
        For j = 0 To grid.Cols - 1
            grid.col = j
            FaqeExcel.Cells(top + i, left + j) = CStr(grid.Text)
        Next
    Next
End Sub

Private Sub ShkruajGride1(grid As MSFlexGrid, top As Integer, left As Integer, FaqeExcel As Worksheet)
    Dim i, j As Integer
    For i = 0 To grid.Rows - 1
        grid.row = i
        For j = 0 To grid.Cols - 1
            grid.col = j
            FaqeExcel.Cells(top + i, left + j) = CStr(grid.Text)
        Next
    Next
End Sub

Private Sub ShkruajGride2(grid As Cgrid, top As Integer, left As Integer, _
FaqeExcel As Worksheet, celWidth As Integer, celHeight As Integer)
    
    Dim i, j, r, C, k As Integer
    i = 1
    j = 1
    k = 1
    r = CInt(grid.Height / celHeight)
    C = CInt(grid.Width / celWidth)
    Dim alfa(100) As String
    Dim str, strAll As String
    str = "CDEFGHIJKLMNOPQRSTUVWXYZ"
    strAll = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    If (grid.Name = "Cgrid1") Then
        Do While (k <= C)
            If (k <= 24) Then
                alfa(k) = Mid(str, k, 1)
            Else
                Dim pos1, pos2 As Integer
                pos2 = (k - 24) Mod 26
                If (pos2 <> 0) Then
                    pos1 = (k - 24) \ 26 + 1
                End If
                If (pos2 <> 0) Then
                    alfa(k) = Mid(strAll, pos1, 1) + Mid(strAll, pos2, 1)
                Else
                    alfa(k) = Mid(strAll, pos1, 1) + "Z"
                End If
            End If
            k = k + 1
        Loop
    End If
    For i = 1 To r
        For j = 1 To C
            FaqeExcel.Cells(top + i - 1, left + j - 1) = grid.Text(i, j)
            If (grid.Name = "Cgrid1") Then
                Dim rng As String
                rng = alfa(j) + CStr(i + 8)
                FaqeExcel.Range(rng).Font.Color = grid.CellForeColor(i, j)
            End If
       Next
    Next
End Sub


Private Sub ShkruajEvidenceVjetoreEksel(txt As String, emrat As MSHFlexGrid, Lendet As MSHFlexGrid, lloji As String, notat As MSHFlexGrid)
    Dim FileExcel        As Workbook
    Dim FaqeExcel      As Worksheet
    Dim fso As New FileSystemObject
    If Not DirExists("C:\RaporteExcel") Then
        Dim objFile As New FileSystemObject
        objFile.CreateFolder ("C:\RaporteExcel")
    End If
    On Error GoTo Gabim:
    fso.CopyFile App.Path + "\Excel\evidencaVjetore.xls", "C:\RaporteExcel\evidencaVjetore.xls", True
    Set fso = Nothing
    
    Set FileExcel = Excel.Workbooks.Open("C:\RaporteExcel\evidencaVjetore.xls")
    Set FaqeExcel = FileExcel.Worksheets("Evidenca vjetore")
    
    FaqeExcel.Cells(1, 2) = emerShkolla
    FaqeExcel.Cells(2, 2) = txt
    
    FaqeExcel.Cells(5, 2) = "Nr."
    FaqeExcel.Cells(5, 3) = "Amza"
    FaqeExcel.Cells(5, 4) = "Emri Mbiemri"
    
    'shkruaj emrat
    Dim i, j As Integer
    For i = 0 To emrat.Rows - 1
        For j = 0 To emrat.Cols - 1
            emrat.col = j
            emrat.row = i
            FaqeExcel.Cells(6 + i, 2 + j) = CStr(emrat.Text)
            
        Next
    Next
    'shkruaj lendet
    Lendet.row = 0
    For i = 0 To Lendet.Cols - 1
        Lendet.col = i
        FaqeExcel.Cells(4, 5 + 3 * i) = CStr(Lendet.Text)
        FaqeExcel.Cells(5, 5 + 3 * i) = "S1"
        FaqeExcel.Cells(5, 5 + 3 * i + 1) = "S2"
        FaqeExcel.Cells(5, 5 + 3 * i + 2) = "V"
    Next
    For i = 0 To notat.Rows - 1
        notat.row = i
        For j = 0 To notat.Cols - 1
            notat.col = j
            FaqeExcel.Cells(6 + i, 5 + j) = notat.Text
        Next
    Next
    FileExcel.Close (True)
    MsgBox "Evidenca vjetore u ruajt në C:\RaporteExcel\evidencaVjetore.xls.", vbOKOnly, "Konvertimi në Excel"
    Exit Sub
Gabim:
    MsgBox "Një skedar me emër evidencaVjetore.xls është i hapur." + Chr(10) + "" & _
    "Mbylleni skedarin para se të bëni konvertimin", vbExclamation, "Konvertimi në Excel"

End Sub

Private Sub ShkruajEvidencaEksel(txt As String, emrat As MSHFlexGrid, Lendet As MSHFlexGrid, lloji As String, notat As MSHFlexGrid)
    Dim FileExcel        As Workbook
    Dim FaqeExcel      As Worksheet
    Dim fso As New FileSystemObject
    If Not DirExists("C:\RaporteExcel") Then
        Dim objFile As New FileSystemObject
        objFile.CreateFolder ("C:\RaporteExcel")
    End If
    On Error GoTo Gabim:
    fso.CopyFile App.Path + "\Excel\evidenca.xls", "C:\RaporteExcel\evidenca.xls", True
    Set fso = Nothing
    
    Set FileExcel = Excel.Workbooks.Open("C:\RaporteExcel\evidenca.xls")
    Set FaqeExcel = FileExcel.Worksheets("Evidenca")
    
        
    FaqeExcel.Cells(1, 2) = emerShkolla
    FaqeExcel.Cells(2, 2) = txt
    
    FaqeExcel.Cells(4, 2) = "Nr."
    FaqeExcel.Cells(4, 3) = "Amza"
    FaqeExcel.Cells(4, 4) = "Emri Mbiemri"
    'shkruaj emrat
    Dim i, j As Integer
    For i = 0 To emrat.Rows - 1
        For j = 0 To emrat.Cols - 1
            emrat.col = j
            emrat.row = i
            FaqeExcel.Cells(5 + i, 2 + j) = CStr(emrat.Text)
            
        Next
    Next
    'shkruaj lendet
    Lendet.row = 0
    For i = 0 To Lendet.Cols - 1
        Lendet.col = i
        FaqeExcel.Cells(4, 5 + 2 * (i + 1) - 2) = CStr(Lendet.Text)
        FaqeExcel.Cells(4, 5 + 2 * (i + 1) - 2 + 1) = lloji
    Next
    'shkruaj notat
    Dim nrMax As Integer
    Lendet.col = 1
    notat.row = 0
    notat.col = 0
    nrMax = Round(Lendet.CellWidth / (notat.CellWidth + 15))
    Dim notatStr As String
    notatStr = ""
    For i = 0 To notat.Rows - 1
        notat.row = i
        For j = 0 To notat.Cols - 1
            If (j Mod nrMax = 0) Then
                If (j <> 0) Then
                    FaqeExcel.Cells(5 + i, 5 + 2 * (j / nrMax) - 2) = notatStr + " "
                End If
                notat.col = j
                notatStr = notat.Text + " "
            'notat perfundimtare
            ElseIf (j Mod nrMax = nrMax - 1) Then
                notat.col = j
                FaqeExcel.Cells(5 + i, 5 + 2 * (j / nrMax) - 2 + 1) = notat.Text
                If (j = notat.Cols - 1) Then
                    FaqeExcel.Cells(5 + i, 5 + 2 * (j / nrMax) - 2) = notatStr + " "
                End If
            Else
                notat.col = j
                notatStr = notatStr + notat.Text + " "
            End If
        Next
    Next
    FileExcel.Close (True)
    MsgBox "Evidenca për semestrin u ruajt në C:\RaporteExcel\evidenca.xls.", vbOKOnly, "Konvertimi në Excel"

    Exit Sub
Gabim:
    MsgBox "Një skedar me emër evidenca.xls është i hapur." + Chr(10) + "" & _
    "Mbylleni skedarin para se të bëni konvertimin", vbExclamation, "Konvertimi në Excel"
End Sub

'gjen adresen e qelizes ne excel kur jepet rreshti dhe kolona
'vlen vetem per kolonat me indeks jo me te madh se  26 * 26 + 26
Private Function GetRange(row As Integer, col As Integer) As String
Dim alfa As String
alfa = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Dim pos1, pos2 As Integer
pos1 = col \ 26
pos2 = col Mod 26

If (col \ 26 = 0) Then
    GetRange = Mid(alfa, col Mod 26, 1) + CStr(row)
ElseIf (col = 26) Then
    GetRange = "Z" + CStr(row)
Else
    If (pos2 <> 0) Then
        pos1 = col \ 26 + 1
    End If
    If (pos2 <> 0) Then
        GetRange = Mid(alfa, pos1, 1) + Mid(alfa, pos2, 1) + CStr(row)
    Else
        GetRange = Mid(strAll, pos1, 1) + "Z" + CStr(row)
    End If
End If
End Function

