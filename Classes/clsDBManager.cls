VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDBManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'
' * Author           :
' * Web Site         :
' * E-Mail           :
' * Date             : 11/22/2004
' * Time             : 20:52
' * Module Name      : clsDBManager
' * Module Filename  : clsDBManager.cls
' * *******************************************************************
' * Comments         :
' *
' *
' * *******************************************************************

Option Explicit

Public dbPath As String
Const dbName As String = "DProgNec.dat"
Dim nota_rreshti As Integer
Dim nota_shtylla As Integer
Dim nResult As Integer
Public nr_lenda_provimet As Integer
Public nr_nxenesit_provimet As Integer

'local variable(s) to hold property value(s)
Private mvarErrorHandler As New clsErrorHandler
Private mvarActionName As BUS_MANAGER_ENUM





Public Property Let actionName(ByVal vData As BUS_MANAGER_ENUM)
    mvarActionName = vData
End Property


Public Property Get actionName() As BUS_MANAGER_ENUM
   actionName = mvarActionName
End Property


Public Property Set ErrorHandler(ByVal vData As clsErrorHandler)
    Set mvarErrorHandler = vData
End Property


Public Property Get ErrorHandler() As clsErrorHandler
   Set ErrorHandler = mvarErrorHandler
End Property


Private Sub Class_Terminate()
   Set ErrorHandler = Nothing
End Sub


Public Sub andi()
   ' Dim nrAmza As String
   Dim i                As Integer, j             As Integer
   Dim vleresimi        As String
   Dim a                As Integer
   Dim b                As Integer
   Dim q_res2           As String
   Dim s1, s2, q_res, q_res1, q_res11, s3 As String
   Dim numerAmze        As String
   Dim cikli            As Integer
   Dim strCikli         As String
   Dim emri             As String
   Dim mbiemri          As String
   Dim ugjet            As Boolean
   Dim nrVitiShkollor, vitishkollor  As String
   Dim Klasa            As String
   Dim indeksi          As String
   Dim atesia           As String
   Dim memesia          As String
   Dim Vendlindja       As String
   Dim Datelindja       As String
   Dim seksi            As String
   Dim Vrejtje          As String
   Dim conn             As New ADODB.Connection
   
   With conn
      .Provider = "Microsoft.Access.OLEDB.10.0"
      .Properties("Data Provider").Value = "SQLOLEDB"
      .Properties("Data Source").Value = EmerServer + "\SHM"
      .Properties("User ID").Value = "sa"
      .Properties("Password").Value = "vision"
      .Properties("Initial Catalog").Value = "DProgNec"
      If (conn.state = 1) Then
        .Close
      End If
      .Open
   End With
   
   'conn.ConnectionString = getConnectionString()
   'conn.Open
   Dim tmp              As New ADODB.Recordset
   Dim tmp11            As New ADODB.Recordset
   Select Case Me.actionName

      Case RUAJ_INFORMACIONE
         Dim emriShkolla      As String
         Dim website          As String
         Dim email            As String
         Dim strSql           As String
         emriShkolla = LCase(active_form.txtAdresa.Text)
         website = (active_form.txtWebsite.Text)
         email = (active_form.txtEmail.Text)
         strSql = " INSERT INTO TBL_PERDORUES ( PERD_EMRI_SHKOLLA, PERD_WEBSITE, PERD_EMAIL ) VALUES ('" & emriShkolla & "','" & website & "', '" & email & "')"
         ex (strSql)

      Case MERR_NOTA_AMZA
         If active_form.optMesme.Value Then
            cikli = 1
            strCikli = "1"
         Else
            cikli = 0
            strCikli = "0"
         End If
         
         If active_form.txtAmzaNo.Text <> "" Then
            numerAmze = active_form.txtAmzaNo.Text
            q_res = "SELECT AMZA_EMRI , AMZA_MBIEMRI , AMZA_KLASA , AMZA_INDEKSI , AMZA_ATESIA , AMZA_MEMESIA , AMZA_SEKSI , AMZA_VENDLINDJA , AMZA_DATELINDJA,AMZA_VITI_SHKOLLOR FROM TBL_AMZA WHERE ( AMZA_NR_AMZA = '" & numerAmze & "' AND AMZA_CIKLI = " & strCikli & ")"
            'q_res = "SELECT * FROM TBL_AMZA WHERE ( AMZA_NR_AMZA = '" & numerAmze & "' AND AMZA_CIKLI = " & strCikli & ")"
            tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
            'get_rec (q_res)
            If Not tmp.EOF Then
               active_form.txtEmri.Text = tmp.Fields("AMZA_EMRI")
               active_form.txtMbiemri.Text = tmp.Fields("AMZA_MBIEMRI")
               active_form.txtAtesia.Text = tmp.Fields("AMZA_ATESIA")
               active_form.txtKlasa.Text = tmp.Fields("AMZA_KLASA")
               active_form.txtIndeksi.Text = tmp.Fields("AMZA_INDEKSI")
               active_form.txtMemesia.Text = tmp.Fields("AMZA_MEMESIA")
               active_form.txtSeksi.Text = tmp.Fields("AMZA_SEKSI")
               active_form.txtVendlindja.Text = tmp.Fields("AMZA_VENDLINDJA")
               active_form.txtDatelindja.Text = tmp.Fields("AMZA_DATELINDJA")
               vitishkollor = tmp.Fields("AMZA_VITI_SHKOLLOR")
               Klasa = tmp.Fields("AMZA_KLASA")
               inicializoGridaAmza9Vjecare active_form.gridaKlasat, active_form.gridaAmza, vitishkollor, CInt(Klasa), cikli
               tmp.Close
Else:
               tmp.Close

               active_form.txtEmri.Text = ""
               active_form.txtMbiemri.Text = ""
               active_form.txtKlasa.Text = ""
               active_form.txtIndeksi.Text = ""
               active_form.txtAtesia.Text = ""
               active_form.txtMemesia.Text = ""
               active_form.txtSeksi.Text = ""
               active_form.txtVendlindja.Text = ""
               active_form.txtDatelindja.Text = ""
               active_form.Label4.Visible = False
               active_form.gridaAmza.Visible = False
               active_form.gridaKlasat.Visible = False
               active_form.gridaLendet.Visible = False
               active_form.gridaLabel.Visible = False
               active_form.rtbShenime.Visible = False
               active_form.fsbVertikali.Visible = False
               active_form.fsbHorizontali.Visible = False
               active_form.Label1.Visible = False
               active_form.cboVitiShkollor.Visible = False
               active_form.lblVitiShkollor.Visible = False
               active_form.Label5.Visible = False
               active_form.Label6.Visible = False
               active_form.Label7.Visible = False
               active_form.gridaPjekurie.Visible = False
               MsgBox "   Nxenesi me numrin e amzes se dhene nuk eshte i regjistruar ne ciklin qe eshte dhene ", vbInformation, "  Konsultime Amza"
               active_form.txtAmzaNo.Text = ""
               Exit Sub
            End If
         Else
         
            emri = active_form.txtEmri.Text
            mbiemri = active_form.txtMbiemri.Text
            atesia = active_form.txtAtesia.Text
            q_res = "SELECT AMZA_NR_AMZA , AMZA_KLASA , AMZA_INDEKSI , AMZA_ATESIA , AMZA_MEMESIA , AMZA_SEKSI , AMZA_VENDLINDJA , AMZA_DATELINDJA,AMZA_EMRI, AMZA_MBIEMRI,AMZA_ATESIA,AMZA_CIKLI,AMZA_VITI_SHKOLLOR FROM TBL_AMZA WHERE ( AMZA_EMRI LIKE '" & emri & "%' AND AMZA_MBIEMRI LIKE '" & mbiemri & "%' AND AMZA_ATESIA LIKE '" & atesia & "%' ) AND  (AMZA_CIKLI = " & strCikli & ") ORDER BY AMZA_EMRI, AMZA_ATESIA, AMZA_MBIEMRI"
            tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
            If Not tmp.EOF Then
                If tmp.RecordCount = 1 Then
                    active_form.txtAmzaNo.Text = tmp.Fields("AMZA_NR_AMZA")
                    active_form.txtKlasa.Text = tmp.Fields("AMZA_KLASA")
                    active_form.txtIndeksi.Text = tmp.Fields("AMZA_INDEKSI")
                    active_form.txtEmri.Text = tmp.Fields("AMZA_EMRI")
                    active_form.txtMbiemri.Text = tmp.Fields("AMZA_MBIEMRI")
                    active_form.txtAtesia.Text = tmp.Fields("AMZA_ATESIA")
                    active_form.txtMemesia.Text = tmp.Fields("AMZA_MEMESIA")
                    active_form.txtSeksi.Text = tmp.Fields("AMZA_SEKSI")
                    active_form.txtVendlindja.Text = tmp.Fields("AMZA_VENDLINDJA")
                    active_form.txtDatelindja.Text = tmp.Fields("AMZA_DATELINDJA")
                    vitishkollor = tmp.Fields("AMZA_VITI_SHKOLLOR")
                    Klasa = tmp.Fields("AMZA_KLASA")
                    inicializoGridaAmza9Vjecare active_form.gridaKlasat, active_form.gridaAmza, vitishkollor, CInt(Klasa), cikli
                    tmp.Close
                Else
                    'nqs ka me teper se nje nxenes me keto te dhena atehere shfaq formen e kerkimit
                    'inicializo gridat ne forme
                    Dim nrRreshta As Integer
                    nrRreshta = tmp.RecordCount
                    Call inicializoGridaZgjidhNxenes(nrRreshta)
                    'mbush griden me te dhena
                    Dim Index As Integer
                    Index = 0
                    Do While Not tmp.EOF
                        frmZgjidhNxenes.gridaNxenesit.row = Index
                        frmZgjidhNxenes.gridaNxenesit.col = 0
                        frmZgjidhNxenes.gridaNxenesit.Text = tmp.Fields("AMZA_NR_AMZA").Value
                        frmZgjidhNxenes.gridaNxenesit.col = 1
                        frmZgjidhNxenes.gridaNxenesit.Text = tmp.Fields("AMZA_EMRI").Value + " " + tmp.Fields("AMZA_ATESIA").Value + " " + tmp.Fields("AMZA_MBIEMRI").Value
                        frmZgjidhNxenes.gridaNxenesit.col = 2
                        frmZgjidhNxenes.gridaNxenesit.Text = tmp.Fields("AMZA_CIKLI").Value
                        tmp.MoveNext
                        Index = Index + 1
                    Loop
                    tmp.Close
                    ShkollaManager.frmZgjidhNxenes.show vbModal
                    numerAmze = ShkollaManager.frmZgjidhNxenes.Amza
                    strCikli = ShkollaManager.frmZgjidhNxenes.cikli
                    If (numerAmze = "" Or strCikli = "") Then
                        'pastro
                        active_form.txtAmzaNo.Text = ""
                        active_form.txtKlasa.Text = ""
                        active_form.txtIndeksi.Text = ""
                        active_form.txtAtesia.Text = ""
                        active_form.txtMemesia.Text = ""
                        active_form.txtSeksi.Text = ""
                        active_form.txtVendlindja.Text = ""
                        active_form.txtDatelindja.Text = ""
                        active_form.gridaAmza.Visible = False
                        active_form.gridaKlasat.Visible = False
                        active_form.gridaLendet.Visible = False
                        active_form.gridaPjekurie.Visible = False
                        active_form.gridaLabel.Visible = False
                        active_form.fsbVertikali.Visible = False
                        active_form.fsbHorizontali.Visible = False
                        active_form.Label4.Visible = False
                        active_form.Label5.Visible = False
                        active_form.Label1.Visible = False
                        active_form.Label7.Visible = False
                        active_form.Label6.Visible = False
                        active_form.lblVitiShkollor.Visible = False
                        active_form.cboVitiShkollor.Visible = False
                        active_form.rtbShenime.Visible = False
                        active_form.txtEmri.Text = ""
                        active_form.txtMbiemri.Text = ""
                        Exit Sub
                    End If
                    q_res = "SELECT AMZA_NR_AMZA,AMZA_EMRI , AMZA_MBIEMRI , AMZA_KLASA , AMZA_INDEKSI , AMZA_ATESIA , AMZA_MEMESIA , AMZA_SEKSI , AMZA_VENDLINDJA , AMZA_DATELINDJA,AMZA_VITI_SHKOLLOR FROM TBL_AMZA WHERE ( AMZA_NR_AMZA = '" & numerAmze & "' AND AMZA_CIKLI = " & strCikli & ")"
                    tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
                    If Not tmp.EOF Then
                        active_form.txtAmzaNo.Text = tmp.Fields("AMZA_NR_AMZA")
                        active_form.txtKlasa.Text = tmp.Fields("AMZA_KLASA")
                        active_form.txtIndeksi.Text = tmp.Fields("AMZA_INDEKSI")
                        active_form.txtMemesia.Text = tmp.Fields("AMZA_MEMESIA")
                        active_form.txtSeksi.Text = tmp.Fields("AMZA_SEKSI")
                        active_form.txtVendlindja.Text = tmp.Fields("AMZA_VENDLINDJA")
                        active_form.txtDatelindja.Text = tmp.Fields("AMZA_DATELINDJA")
                        active_form.txtEmri.Text = tmp.Fields("AMZA_EMRI")
                        active_form.txtMbiemri.Text = tmp.Fields("AMZA_MBIEMRI")
                        active_form.txtAtesia.Text = tmp.Fields("AMZA_ATESIA")
                        vitishkollor = tmp.Fields("AMZA_VITI_SHKOLLOR")
                        Klasa = tmp.Fields("AMZA_KLASA")
                        inicializoGridaAmza9Vjecare active_form.gridaKlasat, active_form.gridaAmza, vitishkollor, CInt(Klasa), cikli
                        tmp.Close
                    End If
                End If
               
Else:
               tmp.Close
               active_form.txtAmzaNo.Text = ""

               active_form.txtKlasa.Text = ""
               active_form.txtIndeksi.Text = ""
               active_form.txtAtesia.Text = ""
               active_form.txtMemesia.Text = ""
               active_form.txtSeksi.Text = ""
               active_form.txtVendlindja.Text = ""
               active_form.txtDatelindja.Text = ""
               active_form.gridaAmza.Visible = False
               active_form.gridaKlasat.Visible = False
               active_form.gridaLendet.Visible = False
               active_form.gridaPjekurie.Visible = False
               active_form.gridaLabel.Visible = False
               active_form.fsbVertikali.Visible = False
               active_form.fsbHorizontali.Visible = False
               active_form.Label4.Visible = False
               active_form.Label5.Visible = False
               active_form.Label1.Visible = False
               active_form.Label7.Visible = False
               active_form.Label6.Visible = False
               active_form.lblVitiShkollor.Visible = False
               active_form.cboVitiShkollor.Visible = False
               active_form.rtbShenime.Visible = False
               MsgBox "   Nxenesi me kete emer dhe mbiemer  nuk eshte i regjistruar ne ciklin qe eshte dhene ", vbInformation, "  Konsultime Amza"
               active_form.txtEmri.Text = ""
               active_form.txtMbiemri.Text = ""
               Exit Sub
            End If
         End If

         numerAmze = active_form.txtAmzaNo.Text
         Dim viti_shkollor    As String
         Dim gjatesia_gride   As Integer
         gjatesia_gride = active_form.gridaAmza.Rows - 1
         q_res = "SELECT DISTINCT KL_KLASA,KL_VITI_SHKOLLOR FROM TBL_KLASA WHERE KL_NR_AMZA = '" & numerAmze & "' AND KL_CIKLI = " & strCikli & " ORDER BY KL_KLASA,KL_VITI_SHKOLLOR"
         tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
         Do While Not tmp.EOF
            For i = 0 To gjatesia_gride
               active_form.gridaKlasat.col = 0
               active_form.gridaKlasat.row = i
               Klasa = active_form.gridaKlasat.Text
               active_form.gridaKlasat.col = 1
               active_form.gridaKlasat.row = i
               viti_shkollor = active_form.gridaAmza.Text
               If Klasa = tmp.Fields("KL_KLASA") Then
                  If viti_shkollor = "" Then
                     active_form.gridaKlasat.col = 1
                     active_form.gridaKlasat.row = i
                     active_form.gridaKlasat.Text = tmp.Fields("KL_VITI_SHKOLLOR")
                     GoTo fund_amza

                  End If
               End If
            Next i
fund_amza:
            tmp.MoveNext

         Loop
         tmp.Close
         Dim tmp1             As New ADODB.Recordset
         Dim lenda            As String
         Dim nota             As String
         q_res = "SELECT NT_KLASA, NT_VITI_SHKOLLOR, NT_EMRI_LENDA, NT_VLERESIMI FROM TBL_NOTA WHERE ( NT_NR_AMZA = '" & numerAmze & "' AND NT_CIKLI = " & strCikli & " ) AND NT_VJETORE = 1 ;"
         q_res1 = "SELECT NT_KLASA, NT_VITI_SHKOLLOR, NT_EMRI_LENDA, NT_VLERESIMI FROM TBL_NOTA WHERE ( NT_NR_AMZA = '" & numerAmze & "' AND NT_CIKLI = " & strCikli & " ) AND NT_RIPROVIM = 1 ;"
         tmp1.Open q_res1, conn, adOpenDynamic, adLockOptimistic
         tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
         'get_rec (q_res)
         Do While Not tmp.EOF

            For i = 0 To active_form.gridaAmza.Rows - 1

               active_form.gridaKlasat.row = i
               active_form.gridaKlasat.col = 0
               Klasa = active_form.gridaKlasat.Text
               active_form.gridaKlasat.row = i
               active_form.gridaKlasat.col = 1
               viti_shkollor = active_form.gridaKlasat.Text
               If tmp.Fields("NT_KLASA") = Klasa And tmp.Fields("NT_VITI_SHKOLLOR") = viti_shkollor Then

                  For j = 0 To active_form.gridaAmza.Cols - 2
                     active_form.gridaLendet.row = 0
                     active_form.gridaLendet.col = j
                     lenda = active_form.gridaLendet.Text
                     If active_form.gridaLendet.Text = tmp.Fields("NT_EMRI_LENDA") Then
                        active_form.gridaAmza.row = i
                        active_form.gridaAmza.col = j
                        active_form.gridaAmza.CellAlignment = 4
                        active_form.gridaAmza.CellFontBold = True
                        nota = tmp.Fields("NT_VLERESIMI")
                        If nota = "4" Then
                           If Not tmp1.EOF Then
                              tmp1.MoveFirst
                              ugjet = False
                              Do While Not tmp1.EOF And Not ugjet
                                 If tmp1.Fields("NT_VITI_SHKOLLOR") = viti_shkollor And tmp1.Fields("NT_EMRI_LENDA") = lenda Then
                                    nota = tmp1.Fields("NT_VLERESIMI")
                                    ugjet = True
                                 End If
                                 tmp1.MoveNext
                              Loop
                           End If
                           active_form.gridaAmza.Text = nota

                        End If
                        active_form.gridaAmza.Text = nota
                     End If
                  Next j

               End If

            Next i
            tmp.MoveNext
         Loop

         tmp.Close
         tmp1.Close
         '*****Ne qofte se nxenesi eshte larguar nga shkolla shkruan ne amze daten qe eshte larguar ne vitin shkollor perkates. * * * * * *

         q_res = "SELECT AMZA_NR_AMZA, AMZA_VITI_SHKOLLOR, AMZA_VREJTJE, AMZA_VITI_LARGIMIT FROM TBL_AMZA WHERE AMZA_NR_AMZA = '" & numerAmze & "'  AND AMZA_CIKLI = " & strCikli
         tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
         If Not tmp.EOF Then
            a = active_form.gridaAmza.Rows
            b = active_form.gridaAmza.Cols
            For i = 0 To a - 1
               active_form.gridaKlasat.row = i
               active_form.gridaKlasat.col = 1
               viti_shkollor = active_form.gridaKlasat.Text
               If viti_shkollor = tmp.Fields("AMZA_VITI_LARGIMIT") Then
                  active_form.gridaAmza.col = b - 1
                  active_form.gridaAmza.row = i
                  active_form.gridaAmza.Text = tmp.Fields("AMZA_VREJTJE")
               End If
            Next i
         End If
         tmp.Close
        active_form.gridaKlasat.row = active_form.gridaKlasat.Rows - 2
        active_form.gridaKlasat.col = 1
        Dim vitShkollor As String
        vitShkollor = active_form.gridaKlasat.Text
         
         
         'mbush grilen e provimeve me emrat e provimeve sipas ciklit
         q_res = "SELECT NTA_EMRI_PROVIMI FROM TBL_NOTAAMZ WHERE NTA_CIKLI = " & cikli & " AND NTA_VITI_SHKOLLOR = '" & vitShkollor & "'"
         tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
         a = 0
         'gjej numrin e provimeve
         If Not tmp.EOF Then
            Do While Not tmp.EOF
               a = a + 1
               tmp.MoveNext
            Loop
            'active_form.gridaPjekurie.Cols = a
            'active_form.gridaPjekurie.Height = 2 * active_form.gridaPjekurie.RowHeight(0) + 30
            'active_form.gridaPjekurie.ColWidth(0) = 1500
            'active_form.gridaPjekurie.Width = active_form.gridaPjekurie.ColWidth(1) * (a) + 120
            'active_form.gridaPjekurie.Rows = 2
            'active_form.gridaPjekurie.Cols = a
            'tmp.MoveFirst
            'Call shfaq_grile2(1, a + 1)

            inicializo_grida_provime active_form.gridaPjekurie, a
            b = 0

            tmp.MoveFirst
            active_form.gridaPjekurie.row = 0
            active_form.gridaPjekurie.col = 0
            Do While Not tmp.EOF
               active_form.gridaPjekurie.col = b
               active_form.gridaPjekurie.ColWidth(b) = 1500
               active_form.gridaPjekurie.CellFontBold = True

               active_form.gridaPjekurie.Text = tmp.Fields("NTA_EMRI_PROVIMI")
               active_form.gridaPjekurie.CellAlignment = 4  ' flexAlignCenterCenter
               b = b + 1
               tmp.MoveNext

            Loop

            'Dim tmp1             As New ADODB.Recordset
            q_res = "SELECT  NT_VLERESIMI,NT_EMRI_LENDA FROM TBL_NOTA WHERE ( NT_NR_AMZA = '" & numerAmze & "' AND NT_CIKLI = " & strCikli & " ) AND NT_PROVIM = 1 "
            tmp1.Open q_res, conn, adOpenDynamic, adLockOptimistic
            ' hedh vleresimet ne grilen e provimeve
            active_form.gridaPjekurie.col = 0
            active_form.gridaPjekurie.row = 1
            b = 0
            Do While Not tmp1.EOF
               Dim cl As Integer
               cl = 1
               For cl = 0 To a - 1
                    active_form.gridaPjekurie.col = cl
                    active_form.gridaPjekurie.row = 0
                    If tmp1.Fields("NT_EMRI_LENDA") = active_form.gridaPjekurie.Text Then
                        active_form.gridaPjekurie.row = 1
                        active_form.gridaPjekurie.CellFontBold = True
                        active_form.gridaPjekurie.Text = tmp1.Fields("NT_VLERESIMI")
                    End If
               Next
               tmp1.MoveNext
            Loop
            tmp1.Close
            active_form.gridaPjekurie.Visible = True
         Else
            active_form.gridaPjekurie.Visible = False
         End If
         tmp.Close
         active_form.gridaAmza.Visible = True
         active_form.gridaKlasat.Visible = True
         active_form.gridaLendet.Visible = True
         active_form.gridaLabel.Visible = True
         'active_form.gridaPjekurie.Visible = True
         active_form.Label4.Visible = True
         active_form.Label5.Visible = True
         active_form.lblVitiShkollor.Visible = True
         active_form.cboVitiShkollor.Visible = True
         If cikli Then
            active_form.Label7.Visible = True
            active_form.fsbVertikali.Visible = False
         Else
            active_form.Label6.Visible = True
            active_form.fsbVertikali.Visible = True
         End If
         active_form.Label1.Visible = True
         active_form.rtbShenime.Visible = True
         If active_form.gridaAmza.Cols > 7 Then
            active_form.fsbHorizontali.Visible = True
            Dim nr               As Integer
            nr = active_form.gridaAmza.Cols + 1 - 8
            Dim distanca         As Integer
            distanca = CInt(12000 / nr)
            ' Dim vlera            As Integer
            'vlera = CInt((active_form.fsbHorizontali.Value / nr) * 8)
            With active_form.fsbHorizontali
               .LargeChange = 2 * distanca
               .SmallChange = distanca

            End With
         End If

         If strCikli = "0" Then
            If active_form.gridaAmza.Cols < 7 Then
               active_form.fsbVertikali.left = 2120 + (active_form.gridaAmza.Cols + 1) * 1500
            End If
         End If

      Case MERR_NOTA_MOMENTALEI

         ' Dim nota             As String
         '*** hedhn notat e nje nxenesi ne gride
         nrVitiShkollor = active_form.cboVitiShkollor.Text

         '*** ketu percaktohet cikli i nxenesit

         If active_form.optMesme.Value Then
            cikli = 1
            strCikli = "1"
         Else
            cikli = 0
            strCikli = "0"
         End If

         '***
         '*** ketu kryhet hedhja e te dhenave plotesuese te nxenesit
         If active_form.txtAmzaNo.Text <> "" Then
            numerAmze = active_form.txtAmzaNo.Text

            q_res = " SELECT AMZA_NR_AMZA, AMZA_EMRI, AMZA_MBIEMRI, AMZA_ATESIA, AMZA_DATELINDJA, AMZA_VREJTJE, AMZA_VENDLINDJA,KL_KLASA,KL_INDEKSI, KL_VITI_SHKOLLOR, PERD_ADRESA, PERD_QYTETI FROM TBL_AMZA, TBL_KLASA, TBL_PERDORUES WHERE ( AMZA_NR_AMZA = KL_NR_AMZA AND AMZA_CIKLI = KL_CIKLI) AND (AMZA_NR_AMZA = '" & numerAmze & "' AND AMZA_CIKLI = " & strCikli & ") AND KL_VITI_SHKOLLOR = '" & nrVitiShkollor & "' ORDER BY AMZA_EMRI, AMZA_MBIEMRI"
            tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic

            q_res1 = "SELECT SJ_SJELLJE FROM TBL_SJELLJE WHERE (SJ_NR_AMZA = '" & _
               numerAmze & "') AND (SJ_CIKLI = " & cikli & ") AND (SJ_VITI_SHKOLLOR = '" & nrVitiShkollor & "')"
            tmp11.Open q_res1, conn, adOpenDynamic, adLockOptimistic

            'get_rec (q_res)
            If Not tmp.EOF Then
               active_form.txtEmri.Text = tmp.Fields("AMZA_EMRI")
               active_form.txtMbiemri.Text = tmp.Fields("AMZA_MBIEMRI")
               active_form.txtAtesia = tmp.Fields("AMZA_ATESIA")
               active_form.txtKlasa.Text = tmp.Fields("KL_KLASA")
               active_form.txtIndeksi.Text = tmp.Fields("KL_INDEKSI")

               If active_form.Name = "frmKonsultimeNota" Then
                  If Not IsNull(tmp.Fields("QYTETI")) Then
                     active_form.Qyteti = tmp.Fields("PERD_QYTETI")
                  End If
                  If Not IsNull(tmp.Fields("PERD_ADRESA")) Then
                     active_form.Adresa = tmp.Fields("PERD_ADRESA")
                  End If
                  active_form.Vendlindja = tmp.Fields("AMZA_VENDLINDJA")
                  active_form.Datelindja = tmp.Fields("AMZA_DATELINDJA")
                  If Not IsNull(tmp11.Fields("SJ_SJELLJE")) Then

                     active_form.Vrejtje = tmp11.Fields("SJ_SJELLJE")
                  End If
               End If
               tmp.Close
               tmp11.Close
Else:
               tmp.Close
               tmp11.Close
               active_form.txtEmri.Text = ""
               active_form.txtMbiemri.Text = ""
               active_form.txtAtesia.Text = ""
               active_form.txtKlasa.Text = ""
               active_form.txtIndeksi.Text = ""
               ' Atesia, Vendlindja, Datelindja, Vrejtje,  sherbejne vetem per
               ' printimin e defteses

               active_form.Label3.Visible = False
               active_form.txtData.Visible = False
               active_form.Cgrid1.Visible = False
               active_form.Lendet.Visible = False
               active_form.Label2.Visible = False
               active_form.lblNotatDheMungesat.Visible = False
               scrbar = False
               scrbarver = False
               active_form.cmdPrint.Enabled = False
               MsgBox "   Nxenesi me numrin e amzes se dhene nuk eshte i regjistruar ne vitin shkollor qe eshte dhene" & vbCrLf & "   dhe ne ciklin qe eshte zgjedhur", vbInformation, "  Konsultime Nota"
               active_form.txtAmzaNo.Text = ""
               Exit Sub
            End If
         Else
            emri = active_form.txtEmri.Text
            mbiemri = active_form.txtMbiemri.Text
            atesia = active_form.txtAtesia.Text
            q_res = " SELECT AMZA_NR_AMZA, AMZA_EMRI, AMZA_MBIEMRI, AMZA_ATESIA, AMZA_DATELINDJA, AMZA_VREJTJE, AMZA_VENDLINDJA,KL_KLASA,KL_INDEKSI, KL_VITI_SHKOLLOR FROM TBL_AMZA, TBL_KLASA WHERE (AMZA_NR_AMZA = KL_NR_AMZA AND AMZA_CIKLI = KL_CIKLI ) AND ( AMZA_EMRI = '" & emri & "' AND AMZA_MBIEMRI = '" & mbiemri & "' AND AMZA_ATESIA = '" & atesia & "' AND AMZA_CIKLI = " & strCikli & ")  AND KL_VITI_SHKOLLOR = '" & nrVitiShkollor & "'"
            'get_rec (q_res)
            tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
            If Not tmp.EOF Then
               active_form.txtAmzaNo.Text = tmp.Fields("AMZA_NR_AMZA")
               active_form.txtKlasa.Text = tmp.Fields("KL_KLASA")
               active_form.txtIndeksi.Text = tmp.Fields("KL_INDEKSI")
               tmp.Close
               active_form.Label3.Visible = True
               active_form.txtData.Visible = True
Else:
               tmp.Close
               active_form.txtEmri.Text = ""
               active_form.txtMbiemri.Text = ""
               active_form.txtAtesia.Text = ""
               active_form.Label3.Visible = False
               active_form.txtData.Visible = False
               active_form.Cgrid1.Visible = False
               active_form.Lendet.Visible = False
               active_form.Label2.Visible = False
               active_form.lblNotatDheMungesat.Visible = False
               MsgBox "   Nxenesi me kete emer dhe mbiemer  nuk eshte i regjistruar ne vitin shkollor qe eshte dhene" & vbCrLf & "   dhe ne ciklin qe eshte zgjedhur", vbInformation, "  Konsultime Nota"
               Exit Sub
            End If
         End If
         numerAmze = active_form.txtAmzaNo.Text

         '***** Merren emri i shkolles dhe rrethi **********
         If active_form.Name = "frmKonsultimeNota" Then
            q_res = "SELECT PERD_EMRI_SHKOLLA, PERD_RRETHI FROM TBL_PERDORUES"
            tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
            If Not tmp.EOF Then
               If Not IsNull(tmp.Fields("PERD_ADRESA")) Then
                  active_form.Rrethi = tmp.Fields("PERD_ADRESA")
               End If
               If Not IsNull(tmp.Fields("PERD_EMRI_SHKOLLA")) Then
                  active_form.ShkEmri = tmp.Fields("PERD_EMRI_SHKOLLA")
               End If
            End If
            tmp.Close
         End If

         '*****

         '*** fillimisht percaktohen numri maksimal i notave  qe mund te kete marre ky nxenes ne nje lende te caktuar

         q_res = "SELECT  NT_EMRI_LENDA FROM TBL_NOTA WHERE (NT_NR_AMZA = '" & numerAmze & "' AND NT_CIKLI = " & strCikli & " ) AND NT_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND NT_PROVIM = 0  ORDER BY NT_EMRI_LENDA, NT_DATA  "
         'get_rec (q_res)
         tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
         Dim max              As Integer
         max = 1
         'Dim nota             As String
         'Dim lenda            As String
         If Not tmp.EOF Then
            nota = tmp.Fields("NT_EMRI_LENDA")
            i = 1
            Do While Not tmp.EOF
               If tmp.Fields("NT_EMRI_LENDA") = nota Then
                  i = i + 1
               Else
                  If max < i Then
                     max = i
                  End If
                  i = 1
                  nota = tmp.Fields("NT_EMRI_LENDA")
               End If
               tmp.MoveNext
            Loop
            If i > max Then
               max = i
            End If
         End If
         tmp.Close

         '*****

         '***** ketu percaktohet numri i lendeve qe zhvillon ky nxenes
         Klasa = active_form.txtKlasa.Text
         q_res = "SELECT DISTINCT NT_EMRI_LENDA FROM TBL_NOTA WHERE (NT_NR_AMZA = '" & numerAmze & "' AND NT_CIKLI = " & strCikli & " ) AND NT_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND NT_PROVIM = 0  ORDER BY NT_EMRI_LENDA  "
         'get_rec (q_res)
         tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
         Dim nrLendet         As Integer
         nrLendet = 0
         Do While Not tmp.EOF
            nrLendet = nrLendet + 1
            tmp.MoveNext
         Loop
         tmp.Close

         '*******
         If max > 0 And nrLendet > 0 Then
            Call shfaq_grile1(max, nrLendet)
         Else

            active_form.Label3.Visible = False
            active_form.txtData.Visible = False
            active_form.Cgrid1.Visible = False
            active_form.Lendet.Visible = False
            active_form.Label2.Visible = False
            active_form.lblNotatDheMungesat.Visible = False
            scrbar = False
            scrbarver = False
            active_form.cmdPrint.Enabled = False
            MsgBox "Nxenesi nuk ka marre akoma asnje note ", vbInformation, "Konsultimi i notave."
         End If

         Dim numerAmza        As String
         numerAmza = active_form.txtAmzaNo.Text

         '**** ketu hidhen emrat e lendeve ne gride
         i = 1
         'q_res = "SELECT DISTINCT  NT_EMRI_LENDA FROM TBL_NOTA WHERE (NT_NR_AMZA = '" & numerAmze & "' AND NT_CIKLI = " & strCikli & " ) AND NT_VITI_SHKOLLOR = '" & nrVitiShkollor & "'  ORDER BY NT_EMRI_LENDA  "
         tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
         'get_rec (q_res)
         Do While Not tmp.EOF
            active_form.Lendet.Text(i, 1) = tmp.Fields("NT_EMRI_LENDA")
            i = i + 1
            tmp.MoveNext
         Loop
         tmp.Close

         '*******

         '******* ketu hidhen notat  ne gride
         'ReDim Preserve matricaNotat(20, 20)
         Dim changedRange     As Integer
         q_res = "SELECT NT_EMRI_LENDA , NT_VLERESIMI , NT_DATA ,NT_TEKST_DATA, NT_MOMENTALES1 , NT_MOMENTALES2 ,NT_DETYREKONTROLLI, NT_SEMESTRI1 , NT_SEMESTRI2 ,NT_VJETORE, NT_MUNGESE_ME, NT_MUNGESE_PA FROM TBL_NOTA WHERE ( NT_NR_AMZA = '" & numerAmze & "' AND NT_CIKLI = " & strCikli & " ) AND NT_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND NT_PROVIM = 0 ORDER BY NT_EMRI_LENDA, NT_DATA  , NT_SEMESTRI1  DESC, NT_SEMESTRI2  DESC, NT_VJETORE  DESC"
         tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
         'get_rec (q_res)
         If Not tmp.EOF Then
            nota = tmp.Fields("NT_EMRI_LENDA")
            i = 1
            j = 1
            Do While Not tmp.EOF
               If tmp.Fields("NT_EMRI_LENDA") = nota Then
                  If (tmp.Fields("NT_MOMENTALES1") Or tmp.Fields("NT_MOMENTALES2")) And tmp.Fields("NT_DETYREKONTROLLI") = False Then
                     active_form.Cgrid1.CellForeColor(i, j) = &H80000008
                     active_form.Cgrid1.CellFontBold(i, j) = True
                  End If

                  If tmp.Fields("NT_SEMESTRI1") Or tmp.Fields("NT_SEMESTRI2") Then
                     active_form.Cgrid1.CellForeColor(i, j) = &HC00000
                     active_form.Cgrid1.CellFontBold(i, j) = True
                  End If
                  If tmp.Fields("NT_VJETORE") Then
                     active_form.Cgrid1.CellForeColor(i, j) = &HFF&
                     active_form.Cgrid1.CellFontBold(i, j) = True
                  End If
                  If tmp.Fields("NT_DETYREKONTROLLI") = True Then
                     active_form.Cgrid1.CellForeColor(i, j) = &H8000&
                     active_form.Cgrid1.CellFontBold(i, j) = True
                  End If
                  If tmp.Fields("NT_MUNGESE_ME") = True Or tmp.Fields("NT_MUNGESE_PA") = True Then
                     active_form.Cgrid1.CellFontBold(i, j) = True
                  End If
                  active_form.Cgrid1.Text(i, j) = tmp.Fields("NT_VLERESIMI")

                  ' Ben ndryshimin e dimensionit te matrices se notave
                  If j > 20 Then
                     changedRange = j
                     'ReDim Preserve matricaNotat(20, 21)
                     'ReDim Preserve matricaNotat(20, changedRange)
                     If j > changedRange Then
                     End If
                  End If
                  If Not IsNull(tmp.Fields("NT_TEKST_DATA")) Then
                     matricaNotat(i, j) = tmp.Fields("NT_TEKST_DATA")
                  End If

                  If (tmp.Fields("NT_MOMENTALES1") Or tmp.Fields("NT_MOMENTALES2")) And tmp.Fields("NT_DETYREKONTROLLI") = False Then
                     active_form.Cgrid1.CellForeColor(i, j) = &H80000008
                     active_form.Cgrid1.CellFontBold(i, j) = True
                  End If

                  If tmp.Fields("NT_SEMESTRI1") Or tmp.Fields("NT_SEMESTRI2") Then
                     active_form.Cgrid1.CellForeColor(i, j) = &HC00000
                     active_form.Cgrid1.CellFontBold(i, j) = True
                  End If
                  If tmp.Fields("NT_VJETORE") Then
                     active_form.Cgrid1.CellForeColor(i, j) = &HFF&
                     active_form.Cgrid1.CellFontBold(i, j) = True
                  End If
                  If tmp.Fields("NT_DETYREKONTROLLI") = True Then
                     active_form.Cgrid1.CellForeColor(i, j) = &H8000&
                     active_form.Cgrid1.CellFontBold(i, j) = True
                  End If
                  If tmp.Fields("NT_MUNGESE_ME") Or tmp.Fields("NT_MUNGESE_PA") Then
                     active_form.Cgrid1.CellFontBold(i, j) = True
                  End If
               Else
                  If j > changedRange Then
                     changedRange = j
                  End If
                  'ReDim Preserve matricaNotat(20, changedRange)
                  i = i + 1
                  j = 1
                  active_form.Cgrid1.Text(i, j) = tmp.Fields("NT_VLERESIMI")

                  matricaNotat(i, j) = tmp.Fields("NT_TEKST_DATA")
                  If (tmp.Fields("NT_MOMENTALES1") Or tmp.Fields("NT_MOMENTALES2")) And tmp.Fields("NT_DETYREKONTROLLI") = False Then
                     active_form.Cgrid1.CellForeColor(i, j) = &H80000008
                     active_form.Cgrid1.CellFontBold(i, j) = True
                  End If

                  If tmp.Fields("NT_SEMESTRI1") Or tmp.Fields("NT_SEMESTRI2") Then
                     active_form.Cgrid1.CellForeColor(i, j) = &HC00000
                     active_form.Cgrid1.CellFontBold(i, j) = True
                  End If
                  If tmp.Fields("NT_VJETORE") Then
                     active_form.Cgrid1.CellForeColor(i, j) = &HFF&
                     active_form.Cgrid1.CellFontBold(i, j) = True
                  End If
                  If tmp.Fields("NT_DETYREKONTROLLI") = True Then
                     active_form.Cgrid1.CellForeColor(i, j) = &H8000&
                     active_form.Cgrid1.CellFontBold(i, j) = True
                  End If
                  If tmp.Fields("NT_MUNGESE_ME") Or tmp.Fields("NT_MUNGESE_PA") Then
                     active_form.Cgrid1.CellFontBold(i, j) = True
                  End If
               End If
               j = j + 1
               nota = tmp.Fields("NT_EMRI_LENDA")

               tmp.MoveNext
            Loop

         End If
         tmp.Close
         For i = 1 To nrLendet
            For j = 1 To max
               active_form.Cgrid1.CellEditMode(i, j) = False

            Next j
            active_form.Lendet.CellEditMode(i, 1) = False
            active_form.Lendet.CellFontBold(i, 1) = True
         Next i

         '***********
      Case MERR_NOTA_EVIDENCA

         Dim Nxenesit(60)     As String
         Dim Lendet(40)       As String
         Dim lendetIndeksi(60, 40) As Integer
         Dim k                As Integer
         Dim maksiNota        As Integer
         Dim lendetNxenesit(60, 30) As Integer
         nrVitiShkollor = active_form.cboVitiShkollor.Text
         Klasa = active_form.cboKlasa.Text
         indeksi = active_form.cboIndeksi.Text
         cikli = ktheCiklin(Klasa, CStr(nrVitiShkollor))
         For i = LBound(Nxenesit) To UBound(Nxenesit)
            Nxenesit(i) = ""
            For j = LBound(Lendet) To UBound(Lendet)
               Lendet(j) = ""
               lendetIndeksi(i, j) = 0
            Next j
         Next i

         If cikli Then
            strCikli = "1"
         Else
            strCikli = "0"
         End If
         '*** q_res1 nxjerr numrat e amzes, emrat dhe mbiemrat e nje klase te caktuar.
         q_res1 = "SELECT AMZA_EMRI , AMZA_MBIEMRI, AMZA_NR_AMZA, KL_KLASA, KL_INDEKSI, KL_VITI_SHKOLLOR FROM TBL_AMZA, TBL_KLASA  WHERE (AMZA_NR_AMZA = KL_NR_AMZA AND AMZA_CIKLI = KL_CIKLI) AND (KL_KLASA = '" & Klasa & "') AND (KL_INDEKSI = '" & indeksi & "') AND (KL_VITI_SHKOLLOR = '" & nrVitiShkollor & "') AND (AMZA_LARGUAR = 0) ORDER BY AMZA_EMRI, AMZA_MBIEMRI "

         '***q_res2 nxjerr lendet qe zhvillon nje klase e dhene gjate nje viti shkollor te dhene
         If strCikli = "1" Then
            q_res2 = "SELECT LN_EMRI FROM TBL_LENDA WHERE LN_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND LN_KLASA" & Klasa & " = 1  ORDER BY LN_NUMRI_MESME "
         Else
            q_res2 = "SELECT LN_EMRI FROM TBL_LENDA WHERE LN_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND LN_KLASA" & Klasa & " = 1  ORDER BY LN_NUMRI_ULET "
         End If
         a = 0  '* do te ruaje gjatesine e rekordsetit q_res1

         Dim recset           As New ADODB.Recordset
         recset.Open q_res1, conn, adOpenDynamic, adLockOptimistic
         ' get_rec (q_res1)
         Do While Not recset.EOF
            a = a + 1
            recset.MoveNext
         Loop

         recset.Close
         'a = a + 1

         b = 0  '*  do te ruaje gjatesine e rekordsetit q_res2

         recset.Open q_res2, conn, adOpenDynamic, adLockOptimistic
         'get_rec (q_res2)
         Do While Not recset.EOF
            b = b + 1
            recset.MoveNext
         Loop

         recset.Close

         '*** inicializohet grida ne formen e evidences
         If (a < 1) Or (b < 1) Then

            MsgBox " Klasa e zgjedhur nga ju nuk ekziston.Nuk ka asnje nxenes per momentin " & Chr(10) & "ose nuk ka asnje lende te regjistruar per kete vit shkollor.", vbInformation, "Konsultimi i evidencave."
            active_form.cboKlasa.ListIndex = -1
            active_form.cboIndeksi.ListIndex = -1
            active_form.fsbVertikali.Visible = False
            active_form.fsbHorizontali.Visible = False

            Exit Sub
         End If

         maksiNota = 0
         If Not active_form.optVjetore.Value Then
            If active_form.optSemestri1.Value Then
               q_res = "SELECT NT_NR_AMZA, NT_EMRI_LENDA, NT_VLERESIMI, NT_SEMESTRI1, NT_MOMENTALES1, AMZA_LARGUAR FROM TBL_NOTA, TBL_AMZA WHERE (AMZA_NR_AMZA = NT_NR_AMZA AND AMZA_CIKLI = NT_CIKLI )  AND ( NT_KLASA = '" & Klasa & "' AND NT_INDEKSI = '" & indeksi & "' AND NT_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND AMZA_LARGUAR = 0) AND ( (NT_MOMENTALES1 = 1 OR  NT_SEMESTRI1 = 1) AND NT_MUNGESE_PA = 0 AND NT_MUNGESE_ME = 0 )  ORDER BY  NT_NR_AMZA, NT_EMRI_LENDA"
            Else
               q_res = "SELECT NT_NR_AMZA, NT_EMRI_LENDA, NT_VLERESIMI, NT_SEMESTRI2, NT_MOMENTALES2, AMZA_LARGUAR FROM TBL_NOTA, TBL_AMZA WHERE (AMZA_NR_AMZA = NT_NR_AMZA AND AMZA_CIKLI = NT_CIKLI)  AND( NT_KLASA = '" & Klasa & "' AND NT_INDEKSI = '" & indeksi & "' AND NT_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND AMZA_LARGUAR = 0) AND ( (NT_MOMENTALES2 = 1 OR NT_SEMESTRI2  = 1) AND NT_MUNGESE_PA = 0 AND NT_MUNGESE_ME = 0 ) ORDER BY  NT_NR_AMZA, NT_EMRI_LENDA"
            End If
            tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
            If tmp.EOF Then
               tmp.Close
               maksiNota = 0
            Else
               numerAmze = tmp.Fields("NT_NR_AMZA")
               lenda = tmp.Fields("NT_EMRI_LENDA")
               tmp.MoveFirst
               i = 0
               Do While Not tmp.EOF

                  If tmp.Fields("NT_EMRI_LENDA") = lenda And tmp.Fields("NT_NR_AMZA") = numerAmze Then
                     i = i + 1
                  Else
                     If maksiNota < i Then
                        maksiNota = i
                     End If
                     numerAmze = tmp.Fields("NT_NR_AMZA")
                     lenda = tmp.Fields("NT_EMRI_LENDA")
                     i = 1
                  End If
                  tmp.MoveNext
               Loop
               tmp.Close
               If maksiNota < i Then
                  maksiNota = i
               End If

            End If

            If maksiNota < 7 Then
               maksiNota = 7
            Else
               maksiNota = maksiNota + 1
            End If

         End If

         If active_form.optVjetore.Value Then
            'Call inicializoGridaEmratNotatVjetore(a, b)
            Call inicializoGridaEvidencaVjetore(active_form.amzaEmerMbiemer, active_form.emrat, active_form.Lendet, active_form.notat, a, b)
         Else
            ' Call inicializoGridaEmratNotat(a, b, maksiNota)
            Call inicializoGridaEvidenca(active_form.amzaEmerMbiemer, active_form.emrat, active_form.Lendet, active_form.notat, a, b, maksiNota)
         End If
         active_form.fsbHorizontali.Value = 0
         active_form.fsbVertikali.Value = 0

         '*** hidhen ne gride numri i amzes, emri , mbiemri per cdo nxenes te nje klase te dhene
         tmp.Open q_res1, conn, adOpenDynamic, adLockOptimistic
         i = 0
         Do While Not tmp.EOF
            'active_form.emrat.Text(i, 1) = tmp.Fields("AMZA_NR_AMZA") & ".  " & tmp.Fields("AMZA_EMRI") & "   " & tmp.Fields("AMZA_MBIEMRI")
            'active_form.Amza.Text(i, 1) = tmp.Fields("AMZA_NR_AMZA")
            active_form.emrat.row = i
            active_form.emrat.col = 0
            active_form.emrat.Text = i + 1
            
            active_form.emrat.row = i
            active_form.emrat.col = 1
            active_form.emrat.Text = tmp.Fields("AMZA_NR_AMZA")

            active_form.emrat.row = i
            active_form.emrat.col = 2
            active_form.emrat.Text = tmp.Fields("AMZA_EMRI") & Space(2) & tmp.Fields("AMZA_MBIEMRI")
            Nxenesit(i) = tmp.Fields("AMZA_NR_AMZA")
            tmp.MoveNext
            i = i + 1
         Loop
         tmp.Close
         j = 0
         tmp.Open q_res2, conn, adOpenDynamic, adLockOptimistic
         Do While Not tmp.EOF

            'active_form.Cgrid2.Text(1, j) = tmp.Fields("LN_EMRI")
            active_form.Lendet.col = j
            active_form.Lendet.row = 0
            active_form.Lendet.Text = tmp.Fields("LN_EMRI")
            Lendet(j) = tmp.Fields("LN_EMRI")
            'lendetIndeksi(j) = 0
            tmp.MoveNext
            j = j + 1
         Loop
         tmp.Close

         If active_form.optSemestri1.Value Then
            '*** hidhen notat e semestrit te pare ne gride .Per nje nxenes te dhene ne nje lende tecaktuar notat hidhen ne nje cell te grides
            q_res = "SELECT NT_NR_AMZA, NT_EMRI_LENDA, NT_VLERESIMI, NT_SEMESTRI1, NT_MOMENTALES1, AMZA_LARGUAR FROM TBL_NOTA, TBL_AMZA WHERE (AMZA_NR_AMZA = NT_NR_AMZA AND AMZA_CIKLI = NT_CIKLI )  AND ( NT_KLASA = '" & Klasa & "' AND NT_INDEKSI = '" & indeksi & "' AND NT_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND AMZA_LARGUAR = 0) AND ( (NT_MOMENTALES1 = 1 OR  NT_SEMESTRI1 = 1) AND NT_MUNGESE_PA = 0 AND NT_MUNGESE_ME = 0 )  ORDER BY  NT_MOMENTALES1  DESC,NT_SEMESTRI1  DESC,NT_DATA"
            tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
            'get_rec (q_res)
            Do While Not tmp.EOF

               lenda = tmp.Fields("NT_EMRI_LENDA")
               numerAmze = tmp.Fields("NT_NR_AMZA")
               For i = 0 To 60
                  If Nxenesit(i) = numerAmze Then
                     b = i
                  End If
               Next i
               For i = 0 To 40
                  If Lendet(i) = lenda Then
                     a = i
                     lendetIndeksi(b, a) = lendetIndeksi(b, a) + 1
                  End If
               Next i

               If tmp.Fields("NT_SEMESTRI1") Then
                  active_form.notat.row = b
                  active_form.notat.col = a * maksiNota + maksiNota - 1

                  active_form.notat.Text = tmp.Fields("NT_VLERESIMI")

                  active_form.notat.CellForeColor = &H800000
                  active_form.notat.CellFontBold = True
               Else
                  active_form.notat.row = b
                  active_form.notat.col = a * maksiNota + lendetIndeksi(b, a) - 1

                  active_form.notat.Text = tmp.Fields("NT_VLERESIMI")

               End If
               tmp.MoveNext
            Loop
            tmp.Close

         End If
         If active_form.optSemestri2.Value Then
            '*** hidhen notat e semestrit te dyte ne gride .Per nje nxenes te dhene ne njelende tecaktuar notat hidhen ne nje cell te grides
            q_res = "SELECT NT_NR_AMZA, NT_EMRI_LENDA, NT_VLERESIMI, NT_SEMESTRI2, NT_MOMENTALES2, AMZA_LARGUAR FROM TBL_NOTA, TBL_AMZA WHERE (AMZA_NR_AMZA = NT_NR_AMZA AND AMZA_CIKLI = NT_CIKLI)  AND( NT_KLASA = '" & Klasa & "' AND NT_INDEKSI = '" & indeksi & "' AND NT_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND AMZA_LARGUAR = 0) AND ( (NT_MOMENTALES2 = 1 OR NT_SEMESTRI2  = 1) AND NT_MUNGESE_PA = 0 AND NT_MUNGESE_ME = 0 ) ORDER BY  NT_MOMENTALES2  DESC, NT_SEMESTRI2  DESC, NT_DATA"
            tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
            'get_rec (q_res)
            Do While Not tmp.EOF

               lenda = tmp.Fields("NT_EMRI_LENDA")
               numerAmze = tmp.Fields("NT_NR_AMZA")
               For i = 0 To 60
                  If Nxenesit(i) = numerAmze Then
                     b = i
                  End If
               Next i
               For i = 0 To 40
                  If Lendet(i) = lenda Then
                     a = i
                     lendetIndeksi(b, a) = lendetIndeksi(b, a) + 1
                  End If
               Next i

               If tmp.Fields("NT_SEMESTRI2") Then
                  active_form.notat.row = b
                  active_form.notat.col = a * maksiNota + maksiNota - 1

                  active_form.notat.Text = tmp.Fields("NT_VLERESIMI")

                  active_form.notat.CellForeColor = &H800000
                  active_form.notat.CellFontBold = True
               Else
                  active_form.notat.row = b
                  active_form.notat.col = a * maksiNota + lendetIndeksi(b, a) - 1

                  active_form.notat.Text = tmp.Fields("NT_VLERESIMI")

               End If
               tmp.MoveNext
            Loop
            tmp.Close

         End If

         If active_form.optVjetore.Value Then
            '*** hidhen notat vjetore ne gride .Per nje nxenes te dhene ne njelende tecaktuar notat hidhen ne nje cell te grides-VJETORE
            q_res = "SELECT NT_NR_AMZA, NT_EMRI_LENDA, NT_VLERESIMI, NT_SEMESTRI1, NT_SEMESTRI2, NT_VJETORE, AMZA_LARGUAR FROM TBL_NOTA, TBL_AMZA WHERE (AMZA_NR_AMZA = NT_NR_AMZA AND AMZA_CIKLI = NT_CIKLI) AND  ( NT_KLASA = '" & Klasa & "' AND NT_INDEKSI = '" & indeksi & "' AND NT_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND AMZA_LARGUAR = 0) And (NT_SEMESTRI2 = 1 OR NT_SEMESTRI1 = 1 OR NT_VJETORE = 1) ORDER BY  NT_SEMESTRI1  DESC, NT_SEMESTRI2  DESC, NT_VJETORE  DESC"
            tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
            'get_rec (q_res)
            Do While Not tmp.EOF

               lenda = tmp.Fields("NT_EMRI_LENDA")
               numerAmze = tmp.Fields("NT_NR_AMZA")
               For i = 0 To 60
                  If Nxenesit(i) = numerAmze Then
                     b = i
                  End If
               Next i
               For i = 0 To 40
                  If Lendet(i) = lenda Then
                     a = i
                     'lendetIndeksi(b, a) = lendetIndeksi(b, a) + 1
                  End If
               Next i

               If tmp.Fields("NT_SEMESTRI1") Then
                  active_form.notat.row = b
                  active_form.notat.col = a * 3
                  active_form.notat.Text = Space(3) & tmp.Fields("NT_VLERESIMI")
                  active_form.notat.CellForeColor = &H800000
               ElseIf tmp.Fields("NT_SEMESTRI2") Then
                  active_form.notat.row = b
                  active_form.notat.col = a * 3 + 1
                  active_form.notat.Text = Space(3) & tmp.Fields("NT_VLERESIMI")
                  active_form.notat.CellForeColor = &H800000
               Else
                  active_form.notat.row = b
                  active_form.notat.col = a * 3 + 2
                  active_form.notat.Text = Space(3) & tmp.Fields("NT_VLERESIMI")
                  active_form.notat.CellForeColor = &HC0&
               End If

               active_form.notat.CellFontBold = True
               tmp.MoveNext
            Loop
            tmp.Close

         End If

         For i = 0 To active_form.notat.Rows - 1
            For j = 0 To active_form.notat.Cols - 1
               active_form.notat.row = i
               active_form.notat.col = j
               If active_form.notat.Text <> "" Then
                  If Len(Trim(active_form.notat.Text)) = 1 Then
                     active_form.notat.Text = Space(1) & active_form.notat.Text
                  End If
               End If
            Next j
         Next i

         active_form.notat.Visible = True
         active_form.emrat.Visible = True
         active_form.Lendet.Visible = True
         active_form.amzaEmerMbiemer.Visible = True
         active_form.cmdPrint.Enabled = True

      Case NXENES_KALO_KLASE

         Dim klasaEre         As String
         Dim indeksiIri       As String
         Dim nrVitiShkollorIri As String
         nrVitiShkollorIri = active_form.cboVitiShkollor2.Text
         klasaEre = active_form.cboKlasa2.Text
         indeksiIri = active_form.cboIndeksi2.Text

         nrVitiShkollor = active_form.cboVitiShkollor.Text
         Klasa = active_form.cboKlasa.Text
         indeksi = active_form.cboIndeksi.Text
         cikli = ktheCiklin(Klasa, CStr(nrVitiShkollor))
         If cikli Then
            strCikli = "1"
         Else
            strCikli = "0"
         End If
         Dim nxenesi          As String
         Dim klasa_nxenesi    As String
         Dim indeksi_nxenesi  As String

         If nrVitiShkollor = nrVitiShkollorIri Then
            'For I = 0 To active_form.lista2.ListCount - 1
            '  nxenesi = Trim(active_form.lista2.List(I))

            ' numerAmze = active_form.lista2Amza.List(I)

            ' q_res = "UPDATE TBL_KLASA SET KL_KLASA = '" & klasaEre & "' , KL_INDEKSI = '" & indeksiIri & "' WHERE KL_NR_AMZA = '" & numerAmze & "' AND KL_CIKLI = " & strcikli & " AND KL_VITI_SHKOLLOR = '" & nrVitiShkollorIri & "'"
            '  conn.Execute q_res
            '  q_res1 = "UPDATE TBL_AMZA SET AMZA_KLASA = '" & klasaEre & "', AMZA_INDEKSI = '" & indeksiIri & "', AMZA_VITI_SHKOLLOR = '" & nrVitiShkollorIri & "' WHERE AMZA_NR_AMZA = '" & numerAmze & "' AND AMZA_CIKLI = " & strcikli
            ' conn.Execute q_res1
            ' q_res1 = "UPDATE TBL_NOTA SET NT_INDEKSI = '" & indeksiIri & "' WHERE NT_NR_AMZA = '" & numerAmze & "' AND NT_CIKLI = " & strcikli & " AND NT_VITI_SHKOLLOR = '" & nrVitiShkollorIri & "' AND NT_KLASA = '" & Klasa & "'"
            ' conn.Execute q_res1
            'Next I
            q_res1 = "DELETE FROM TBL_KLASA WHERE  KL_KLASA = '" & Klasa & "' AND KL_INDEKSI = '" & indeksi & "' AND KL_VITI_SHKOLLOR = '" & nrVitiShkollor & "'"
            conn.Execute q_res1
            q_res1 = "DELETE FROM TBL_KLASA WHERE  KL_KLASA = '" & klasaEre & "' AND KL_INDEKSI = '" & indeksiIri & "' AND KL_VITI_SHKOLLOR = '" & nrVitiShkollorIri & "'"
            conn.Execute q_res1

            For i = 0 To active_form.lista2.ListCount - 1

               numerAmze = active_form.lista2Amza.List(i)
               q_res = "INSERT INTO TBL_KLASA (KL_NR_AMZA, KL_CIKLI, KL_VITI_SHKOLLOR,KL_KLASA, KL_INDEKSI) VALUES ('" & numerAmze & "', " & strCikli & " , '" & nrVitiShkollorIri & "', '" & klasaEre & "', '" & indeksiIri & "')"
               conn.Execute q_res
               q_res1 = "UPDATE TBL_AMZA SET AMZA_KLASA = '" & klasaEre & "', AMZA_INDEKSI = '" & indeksiIri & "', AMZA_VITI_SHKOLLOR = '" & nrVitiShkollorIri & "' WHERE AMZA_NR_AMZA = '" & numerAmze & "' AND AMZA_CIKLI = " & strCikli
               conn.Execute q_res1
               q_res1 = "UPDATE TBL_NOTA SET NT_KLASA = '" & klasaEre & "', NT_INDEKSI = '" & indeksiIri & "' WHERE NT_NR_AMZA = '" & numerAmze & "' AND NT_CIKLI = " & strCikli & " AND NT_VITI_SHKOLLOR = '" & nrVitiShkollor & "'"
               conn.Execute q_res1
            Next i
            For i = 0 To active_form.lista1.ListCount - 1

               numerAmze = active_form.listaAmza.List(i)
               q_res = "INSERT INTO TBL_KLASA (KL_NR_AMZA, KL_CIKLI, KL_VITI_SHKOLLOR,KL_KLASA, KL_INDEKSI) VALUES ('" & numerAmze & "', " & strCikli & " , '" & nrVitiShkollor & "', '" & Klasa & "', '" & indeksi & "')"
               conn.Execute q_res
               q_res1 = "UPDATE TBL_AMZA SET AMZA_KLASA = '" & Klasa & "', AMZA_INDEKSI = '" & indeksi & "', AMZA_VITI_SHKOLLOR = '" & nrVitiShkollor & "' WHERE AMZA_NR_AMZA = '" & numerAmze & "' AND AMZA_CIKLI = " & strCikli
               conn.Execute q_res1
               q_res1 = "UPDATE TBL_NOTA SET NT_KLASA = '" & Klasa & "', NT_INDEKSI = '" & indeksi & "' WHERE NT_NR_AMZA = '" & numerAmze & "' AND NT_CIKLI = " & strCikli & " AND NT_VITI_SHKOLLOR = '" & nrVitiShkollor & "'"
               conn.Execute q_res1

            Next i
         Else
            q_res1 = "DELETE FROM TBL_KLASA WHERE  KL_KLASA = '" & klasaEre & "' AND KL_INDEKSI = '" & indeksiIri & "' AND KL_VITI_SHKOLLOR = '" & nrVitiShkollorIri & "'"
            conn.Execute q_res1
            For i = 0 To active_form.lista2.ListCount - 1
               nxenesi = Trim(active_form.lista2.List(i))

               numerAmze = active_form.lista2Amza.List(i)

               q_res11 = "SELECT KL_NR_AMZA, KL_CIKLI, KL_KLASA, KL_INDEKSI FROM TBL_KLASA WHERE KL_NR_AMZA = '" & numerAmze & "' AND KL_CIKLI = " & strCikli & " AND KL_VITI_SHKOLLOR = '" & nrVitiShkollorIri & "'"
               tmp11.Open q_res11, conn, adOpenDynamic, adLockOptimistic
               If tmp11.EOF Then
                  q_res = "INSERT INTO TBL_KLASA (KL_NR_AMZA, KL_CIKLI, KL_VITI_SHKOLLOR,KL_KLASA, KL_INDEKSI) VALUES ('" & numerAmze & "', " & strCikli & " , '" & nrVitiShkollorIri & "', '" & klasaEre & "', '" & indeksiIri & "')"
                  conn.Execute q_res
                  q_res1 = "UPDATE TBL_AMZA SET AMZA_KLASA = '" & klasaEre & "', AMZA_INDEKSI = '" & indeksiIri & "', AMZA_VITI_SHKOLLOR = '" & nrVitiShkollorIri & "' WHERE AMZA_NR_AMZA = '" & numerAmze & "' AND AMZA_CIKLI = " & strCikli
                  conn.Execute q_res1
               Else
                  klasa_nxenesi = tmp11.Fields("KL_KLASA")
                  indeksi_nxenesi = tmp11.Fields("KL_INDEKSI")
                  MsgBox "Nxenesi me emrin " & nxenesi & " per vitin shkollor " & nrVitiShkollorIri & " eshte regjistruar ne klasen " & klasa_nxenesi & "-" & indeksi_nxenesi & " .", vbInformation, "Konfigurimi i klasave."
               End If
               tmp11.Close
            Next i

         End If
         'ex (q_res)

         MsgBox "Nxenesit u hodhen ne klasen e re.", vbInformation, "Konfigurimi i klasave."
         active_form.lista1.Clear
         active_form.lista2.Clear
         active_form.cboKlasa.ListIndex = -1
         active_form.cboIndeksi.ListIndex = -1
         active_form.cboVitiShkollor.ListIndex = -1
         active_form.cboKlasa2.ListIndex = -1
         active_form.cboIndeksi2.ListIndex = -1
         active_form.cboVitiShkollor2.ListIndex = -1

      Case KALO_LENDE
         'Dim lenda As Integer
         If active_form.optMesme.Value = True Then
            strCikli = "1"
            cikli = 1
         Else
            strCikli = "0"
            cikli = 0
         End If
         Dim lendaNumri       As Integer
         Klasa = active_form.cboKlasa.Text
         nrVitiShkollor = active_form.cboVitiShkollor.Text
         Dim ndodhet          As Boolean
         q_res1 = " SELECT LN_EMRI, LN_VITI_SHKOLLOR FROM TBL_LENDA"
         tmp.Open q_res1, conn, adOpenDynamic, adLockOptimistic
         'ex (q_res)
         q_res2 = "UPDATE TBL_LENDA SET LN_KLASA" & Klasa & " = 0 WHERE  LN_VITI_SHKOLLOR = '" & nrVitiShkollor & "'"
         conn.Execute q_res2
         If Not tmp.EOF Then
            For i = 0 To active_form.lista2.ListCount - 1
               lenda = active_form.lista2.List(i)
               lendaNumri = CInt(active_form.listaNumri.List(i))
               ndodhet = False
               tmp.MoveFirst
               Do While Not tmp.EOF And Not ndodhet
                  If (tmp.Fields("LN_EMRI") = lenda) And tmp.Fields("LN_VITI_SHKOLLOR") = nrVitiShkollor Then
                     ndodhet = True
                     tmp.MoveLast
                  Else
                     tmp.MoveNext
                  End If
               Loop
               If Not ndodhet Then
                  If cikli Then
                     q_res = "INSERT INTO TBL_LENDA (LN_EMRI, LN_VITI_SHKOLLOR,LN_KLASA" & Klasa & ", LN_NUMRI_MESME ) VALUES ( '" & lenda & "', '" & nrVitiShkollor & "', 1 ," & lendaNumri & ")"
                  Else
                     q_res = "INSERT INTO TBL_LENDA (LN_EMRI, LN_VITI_SHKOLLOR,LN_KLASA" & Klasa & ", LN_NUMRI_ULET ) VALUES ( '" & lenda & "', '" & nrVitiShkollor & "', 1 ," & lendaNumri & ")"
                  End If
               Else
                  If cikli Then
                     q_res = "UPDATE TBL_LENDA SET LN_KLASA" & Klasa & " = 1, LN_NUMRI_MESME = " & lendaNumri & " WHERE ( LN_EMRI = '" & lenda & "' AND LN_VITI_SHKOLLOR = '" & nrVitiShkollor & "')"
                  Else
                     q_res = "UPDATE TBL_LENDA SET LN_KLASA" & Klasa & " = 1, LN_NUMRI_ULET = " & lendaNumri & " WHERE ( LN_EMRI = '" & lenda & "' AND LN_VITI_SHKOLLOR = '" & nrVitiShkollor & "' )"
                  End If

               End If
               conn.Execute q_res
            Next i
         Else
            For i = 0 To active_form.lista2.ListCount - 1
               lenda = active_form.lista2.List(i)
               lendaNumri = CInt(active_form.listaNumri.List(i))
               If cikli Then
                  q_res = "INSERT INTO TBL_LENDA (LN_EMRI, LN_VITI_SHKOLLOR,LN_KLASA" & Klasa & ", LN_NUMRI_MESME ) VALUES ( '" & lenda & "', '" & nrVitiShkollor & "', 1 ," & lendaNumri & ")"
               Else
                  q_res = "INSERT INTO TBL_LENDA (LN_EMRI, LN_VITI_SHKOLLOR,LN_KLASA" & Klasa & ", LN_NUMRI_ULET ) VALUES ( '" & lenda & "', '" & nrVitiShkollor & "', 1 ," & lendaNumri & ")"
               End If

               conn.Execute q_res
               'ex (q_res)
            Next i
         End If
         tmp.Close
      Case LENDE_HIDH_PROVIME

         If active_form.optMesme.Value = True Then
            strCikli = "1"
         Else
            strCikli = "0"
         End If

         Klasa = active_form.cboKlasa.Text
         nrVitiShkollor = active_form.cboVitiShkollor.Text
         q_res = "DELETE FROM TBL_NOTAAMZ WHERE NTA_CIKLI = " & strCikli & " AND NTA_VITI_SHKOLLOR = '" & nrVitiShkollor & "'"
         conn.Execute q_res
         Dim provimi          As String
         q_res1 = "SELECT NTA_EMRI_PROVIMI FROM TBL_NOTAAMZ WHERE NTA_CIKLI = " & strCikli & " AND NTA_VITI_SHKOLLOR = '" & nrVitiShkollor & "'"
         tmp.Open q_res1, conn, adOpenDynamic, adLockOptimistic

         If Not tmp.EOF Then
            For i = 0 To active_form.listaProvimet.ListCount - 1
               provimi = active_form.listaProvimet.List(i)
               ndodhet = False
               tmp.MoveFirst
               Do While Not tmp.EOF And Not ndodhet
                  If LCase(provimi) = LCase(tmp.Fields("NTA_EMRI_PROVIMI")) Then

                     ndodhet = True
                  End If
                  tmp.MoveNext
               Loop

               If Not ndodhet Then

                  q_res = "INSERT INTO TBL_NOTAAMZ (NTA_CIKLI, NTA_VITI_SHKOLLOR, NTA_EMRI_PROVIMI) VALUES(" & strCikli & ", '" & nrVitiShkollor & "', '" & provimi & "')"
                  conn.Execute q_res
               End If
            Next i

         Else

            For i = 0 To active_form.listaProvimet.ListCount - 1
               provimi = active_form.listaProvimet.List(i)

               q_res = "INSERT INTO TBL_NOTAAMZ (NTA_CIKLI, NTA_VITI_SHKOLLOR, NTA_EMRI_PROVIMI) VALUES(" & strCikli & ", '" & nrVitiShkollor & "', '" & provimi & "')"
               conn.Execute q_res

            Next i
         End If
         tmp.Close

      Case HIDH_PROVIME
         'shfaq grilen sipas provimeve
         Klasa = active_form.txtKlasa.Text
         cikli = ktheCiklin(Klasa, CStr(active_form.cboVitiShkollor.Text))
         If cikli = True Then
            strCikli = "1"
            active_form.Label5.Visible = True
            ' active_form.Label3.Left = 2640
         Else
            strCikli = "0"

            active_form.Label3.Visible = True
         End If
         q_res = "SELECT NTA_EMRI_PROVIMI FROM TBL_NOTAAMZ WHERE NTA_CIKLI = " & cikli
         tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
         a = 0
         'gjej numrin e provimeve
         Do While Not tmp.EOF
            a = a + 1
            tmp.MoveNext
         Loop
         Call shfaq_grile2(2, a)
         tmp.MoveFirst

         b = 0
         'mbushet grila me emrat e provimeve
         Do While Not tmp.EOF
            b = b + 1
            active_form.provimet.Text(1, b) = tmp.Fields("NTA_EMRI_PROVIMI")
            tmp.MoveNext
         Loop
         tmp.Close

      Case HIDH_GJENERALITETE
         Dim Mp               As Processa
         Set Mp = New Processa
         Dim nrAmza           As String
         numerAmze = active_form.txtAmzaNo.Text
         emri = active_form.txtEmri.Text
         mbiemri = active_form.txtMbiemri.Text
         Klasa = active_form.cboKlasa.Text
         indeksi = active_form.cboIndeksi.Text
         atesia = active_form.txtAtesia.Text
         memesia = active_form.txtMemesia.Text
         seksi = active_form.cboSeksi.Text
         Vendlindja = active_form.txtVendlindja.Text
         Datelindja = active_form.data.Value
         nrVitiShkollor = active_form.cboVitiShkollor.Text
         cikli = ktheCiklin(Klasa, CStr(nrVitiShkollor))
         If cikli = True Then
            strCikli = "1"
         Else
            strCikli = "0"
         End If
         'vrejtje = active_form.rtbShenime.Text
         'hidhen gjeneralitetet
         q_res = "SELECT AMZA_NR_AMZA, AMZA_EMRI, AMZA_MBIEMRI, AMZA_ATESIA , AMZA_VITI_SHKOLLOR FROM TBL_AMZA WHERE AMZA_NR_AMZA = '" & numerAmze & "' AND AMZA_EMRI = '" & emri & "' AND AMZA_MBIEMRI = '" & mbiemri & "' AND AMZA_ATESIA = '" & atesia & "'"
         tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
         Dim vitishkollorIvjeter As String
         If Not tmp.EOF Then
            vitishkollorIvjeter = tmp.Fields("AMZA_VITI_SHKOLLOR")
            nrAmza = tmp.Fields("AMZA_NR_AMZA")
            MsgBox "Nxenesi me keto te dhena ekziston tashme ne amze." & Chr(10) & "Ne qofte se ka qene i larguar po e regjistrojme perseri me te dhenat ekzistuese.", vbInformation, "Hedhje gjeneralitetesh"
            active_form.txtAmzaNo.Text = ""
            active_form.txtEmri.Text = ""
            active_form.txtMbiemri.Text = ""
            active_form.cboKlasa.ListIndex = -1
            active_form.cboIndeksi.ListIndex = -1
            active_form.txtAtesia.Text = ""
            active_form.txtMemesia.Text = ""
            active_form.cboSeksi.ListIndex = -1
            active_form.txtVendlindja.Text = ""
            active_form.txtData.Visible = True

            active_form.cboVitiShkollor.ListIndex = -1
            q_res = "UPDATE TBL_AMZA SET AMZA_LARGUAR = 0, AMZA_KLASA = '" & Klasa & "',AMZA_INDEKSI = '" & indeksi & "', AMZA_VITI_SHKOLLOR = '" & nrVitiShkollor & "', AMZA_MEMESIA = '" & memesia & "', AMZA_SEKSI = '" & seksi & "', AMZA_VENDLINDJA = '" & Vendlindja & "', AMZA_DATELINDJA = '" & Datelindja & "' WHERE AMZA_EMRI = '" & emri & "' AND AMZA_MBIEMRI = '" & mbiemri & "' AND AMZA_ATESIA = '" & atesia & "'"
            conn.Execute q_res
            If nrVitiShkollor = vitishkollorIvjeter Then
               q_res = "UPDATE TBL_AMZA SET AMZA_VREJTJE = '' WHERE AMZA_NR_AMZA = '" & numerAmze & "' AND AMZA_CIKLI = " & strCikli & " AND AMZA_VITI_SHKOLLOR = '" & nrVitiShkollor & "'"
               conn.Execute q_res
            End If
            On Error GoTo fund_hedhje_amze
            q_res = "INSERT INTO TBL_KLASA ( KL_NR_AMZA, KL_CIKLI, KL_VITI_SHKOLLOR, KL_KLASA, KL_INDEKSI) VALUES ('" & nrAmza & "'," & strCikli & ", '" & nrVitiShkollor & "', '" & Klasa & "', '" & indeksi & "')"
            conn.Execute q_res
            active_form.txtAmzaNo.Text = ""
            active_form.txtEmri.Text = ""
            active_form.txtMbiemri.Text = ""
            active_form.cboKlasa.ListIndex = -1
            active_form.cboIndeksi.ListIndex = -1
            active_form.txtAtesia.Text = ""
            active_form.txtMemesia.Text = ""
            active_form.cboSeksi.ListIndex = -1
            active_form.txtVendlindja.Text = ""
            active_form.txtData.Visible = True

            active_form.cboVitiShkollor.ListIndex = -1
            tmp.Close
            GoTo fund_hedhje_amze
         End If
         tmp.Close

         q_res1 = "SELECT AMZA_NR_AMZA, AMZA_VITI_SHKOLLOR FROM TBL_AMZA WHERE AMZA_CIKLI = " & strCikli
         tmp.Open q_res1, conn, adOpenDynamic, adLockOptimistic
         ugjet = False
         Do While Not tmp.EOF And Not ugjet
            If tmp.Fields("AMZA_NR_AMZA") = numerAmze Then
               ugjet = True
               vitishkollorIvjeter = tmp.Fields("AMZA_VITI_SHKOLLOR")
            End If
            tmp.MoveNext
         Loop
         tmp.Close
         Dim ugjet1           As Boolean
         q_res2 = "SELECT AMZA_NR_AMZA, AMZA_EMRI, AMZA_MBIEMRI , AMZA_ATESIA, AMZA_VITI_SHKOLLOR FROM TBL_AMZA WHERE AMZA_CIKLI = " & strCikli
         tmp.Open q_res2, conn, adOpenDynamic, adLockOptimistic
         ugjet1 = False
         Do While Not tmp.EOF And Not ugjet1
            If LCase(tmp.Fields("AMZA_EMRI")) = LCase(emri) And LCase(tmp.Fields("AMZA_MBIEMRI")) = LCase(mbiemri) And LCase(tmp.Fields("AMZA_ATESIA")) = LCase(atesia) Then
               ugjet1 = True
               nrAmza = tmp.Fields("AMZA_NR_AMZA")
               vitishkollorIvjeter = tmp.Fields("AMZA_VITI_SHKOLLOR")
            End If
            tmp.MoveNext
         Loop
         tmp.Close
         Dim sig              As Integer
         q_res2 = "SELECT AMZA_EMRI FROM TBL_AMZA"
         tmp.Open q_res2, conn, adOpenDynamic, adLockOptimistic
         sig = 1
         Do While Not tmp.EOF
            tmp.MoveNext
            sig = sig + 1
         Loop
         tmp.Close
         If lic.RegistryNLic = "OEH-00000000000" Then
            If sig < 30 Then
hidh:                                                                                                                                                                      If Not (ugjet Or ugjet1) Then

               q_res = " INSERT INTO TBL_AMZA (AMZA_NR_AMZA, AMZA_CIKLI, AMZA_EMRI, AMZA_MBIEMRI, AMZA_ATESIA, AMZA_MEMESIA, AMZA_SEKSI, AMZA_VENDLINDJA, AMZA_DATELINDJA, AMZA_VITI_SHKOLLOR, AMZA_INDEKSI, AMZA_KLASA,AMZA_LARGUAR ) VALUES ('" & numerAmze & "', " & strCikli & ", '" & emri & "', '" & mbiemri & "', '" & atesia & "', '" & memesia & "', '" & seksi & "', '" & Vendlindja & "', '" & Datelindja & "', '" & nrVitiShkollor & "', '" & indeksi & "', '" & Klasa & "',0)"
               conn.Execute q_res
               q_res = "INSERT INTO TBL_KLASA (KL_NR_AMZA, KL_CIKLI, KL_VITI_SHKOLLOR, KL_KLASA, KL_INDEKSI) VALUES ('" & numerAmze & "', " & strCikli & ", '" & nrVitiShkollor & "', '" & Klasa & "', '" & indeksi & "')"
               conn.Execute q_res
               MsgBox "Nxenesi i ri u hodh ne amze.", vbInformation, "Hedhje gjeneralitetesh."

               active_form.txtAmzaNo.Text = ""
               active_form.txtEmri.Text = ""
               active_form.txtMbiemri.Text = ""
               active_form.cboKlasa.ListIndex = -1
               active_form.cboIndeksi.ListIndex = -1
               active_form.txtAtesia.Text = ""
               active_form.txtMemesia.Text = ""
               active_form.cboSeksi.ListIndex = -1
               active_form.txtVendlindja.Text = ""
               active_form.txtData.Visible = True

               active_form.cboVitiShkollor.ListIndex = -1
            End If

            If ugjet Or ugjet1 Then
               If ugjet1 Then
                  MsgBox "Nxenesi me keto te dhena ekziston tashme ne amze.", vbInformation, "Hedhje gjeneralitetesh"
                  q_res = "UPDATE TBL_AMZA SET AMZA_LARGUAR = 0, AMZA_KLASA = '" & Klasa & "' ,AMZA_INDEKSI = '" & indeksi & "' , AMZA_VITI_SHKOLLOR = '" & nrVitiShkollor & "', AMZA_MEMESIA = '" & memesia & "', AMZA_SEKSI = '" & seksi & "', AMZA_VENDLINDJA = '" & Vendlindja & "', AMZA_DATELINDJA = '" & Datelindja & "' WHERE AMZA_EMRI = '" & emri & "' AND AMZA_MBIEMRI = '" & mbiemri & "' AND AMZA_ATESIA = '" & atesia & "'"
                  conn.Execute q_res
                  active_form.txtAmzaNo.Text = ""
                  active_form.txtEmri.Text = ""
                  active_form.txtMbiemri.Text = ""
                  active_form.cboKlasa.ListIndex = -1
                  active_form.cboIndeksi.ListIndex = -1
                  active_form.txtAtesia.Text = ""
                  active_form.txtMemesia.Text = ""
                  active_form.cboSeksi.ListIndex = -1
                  active_form.txtVendlindja.Text = ""
                  active_form.txtData.Visible = True

                  active_form.cboVitiShkollor.ListIndex = -1
                  If nrVitiShkollor = vitishkollorIvjeter Then
                     q_res = "UPDATE TBL_AMZA SET AMZA_VREJTJE = '' WHERE AMZA_EMRI = '" & emri & "' AND AMZA_MBIEMRI = '" & mbiemri & "' AND AMZA_ATESIA = '" & atesia & "' AND AMZA_CIKLI = " & strCikli & " AND AMZA_VITI_SHKOLLOR = '" & nrVitiShkollor & "'"
                     conn.Execute q_res

                  End If
                  On Error GoTo fund_hedhje_amze

                  q_res = "INSERT INTO TBL_KLASA ( KL_NR_AMZA, KL_CIKLI, KL_VITI_SHKOLLOR, KL_KLASA, KL_INDEKSI) VALUES ('" & nrAmza & "'," & strCikli & ", '" & nrVitiShkollor & "', '" & Klasa & "', '" & indeksi & "')"
                  conn.Execute q_res
                  active_form.txtAmzaNo.Text = ""
                  active_form.txtEmri.Text = ""
                  active_form.txtMbiemri.Text = ""
                  active_form.cboKlasa.ListIndex = -1
                  active_form.cboIndeksi.ListIndex = -1
                  active_form.txtAtesia.Text = ""
                  active_form.txtMemesia.Text = ""
                  active_form.cboSeksi.ListIndex = -1
                  active_form.txtVendlindja.Text = ""
                  active_form.txtData.Visible = True

                  active_form.cboVitiShkollor.ListIndex = -1
               Else
                  MsgBox "Numri i amzes qe ju kerkoni te hidhni eksiston tashme ne amze." & Chr(10) & "Jepni nxenesit  nje numer te ri.", vbInformation, "Hedhje gjeneralitetesh."
               End If

            End If
         Else
            MsgBox "Nuk mund te fusni me shume se 30 nxenes ne versionin demostrativ", vbInformation, "Versioni demostrativ"
         End If
      Else
         GoTo hidh
      End If
fund_hedhje_amze:
   Case MODIFIKO_NOTA
      Dim notaEvjeter      As String
      Dim notaEre          As String
      Dim rreshti          As Integer
      Dim shtylla          As Integer
      Dim tipiNota         As Integer
      Dim kushti           As String
      tipiNota = active_form.txtTipi.Text
      Select Case tipiNota
         Case 1
            kushti = " AND (NT_MOMENTALES1 = 1 OR NT_MOMENTALES2 = 1 )"
         Case 2
            kushti = " AND NT_SEMESTRI1 = 1 "
         Case 3
            kushti = " AND NT_SEMESTRI2 = 1 "
         Case 4
            kushti = " AND NT_VJETORE = 1 "
         Case 5
            kushti = " AND NT_RIPROVIM = 1 "

      End Select
      rreshti = rreshtiModifiko     '  active_form.Cgrid1.CurrentRow
      shtylla = shtyllaModifiko     'active_form.Cgrid1.CurrentCol
      If active_form.optMesme.Value Then
         strCikli = "1"
      Else
         strCikli = "0"
      End If
      numerAmze = active_form.txtAmzaNo.Text
      Dim data             As String
      data = active_form.txtData.Text
      nrVitiShkollor = active_form.cboVitiShkollor.Text
      lenda = active_form.Lendet.Text(rreshti, 1)
      notaEvjeter = Trim(active_form.Cgrid1.Text(rreshti, shtylla))
      'Dim data_nota        As Date
      'data_nota = active_form.txt.Value
      'data_nota = CDate(data)
      'data_nota = CDate(active_form.mskData.Text)
      'data_nota = active_form.datePiker.Value
      If modifikoNota = 1 Then
         q_res = "DELETE FROM TBL_NOTA WHERE ( NT_NR_AMZA = '" & numerAmze & "' AND NT_CIKLI = " & strCikli & ") AND ( NT_EMRI_LENDA = '" & lenda & "' AND NT_TEKST_DATA = '" & data & "') AND NT_VITI_SHKOLLOR = '" & nrVitiShkollor & "'" & kushti
         conn.Execute q_res
         active_form.Cgrid1.Text(rreshti, shtylla) = ""
      ElseIf modifikoNota = 2 Then
         notaEre = InputBox("Jepni noten e re.", "Modifikoni noten e zgjedhur")
         If notaEre = "" Then
            Exit Sub
         End If
         If notaEvjeter = "m" Then
            If notaEre = "m" Then
               q_res = " UPDATE TBL_NOTA SET NT_VLERESIMI = '" & notaEre & "' WHERE ( NT_NR_AMZA = '" & numerAmze & "' AND NT_CIKLI = " & strCikli & ") AND ( NT_EMRI_LENDA = '" & lenda & "' AND NT_TEKST_DATA = '" & data & "') AND NT_VITI_SHKOLLOR = '" & nrVitiShkollor & "'" & kushti

               If notaEre <> notaEvjeter Then
                  conn.Execute q_res
                  'ex (q_res)
                  active_form.Cgrid1.Text(rreshti, shtylla) = notaEre
               End If
            Else
               MsgBox "Mungesa e hedhur nuk eshte ne formatin e duhur.", vbInformation, "Modifikim note."
            End If
         Else
            If notaEre = "4" Or notaEre = "5" Or notaEre = "6" Or notaEre = "7" Or notaEre = "8" Or notaEre = "9" Or notaEre = "10" Then
               q_res = " UPDATE TBL_NOTA SET NT_VLERESIMI = '" & notaEre & "' WHERE ( NT_NR_AMZA = '" & numerAmze & "' AND NT_CIKLI = " & strCikli & ") AND ( NT_EMRI_LENDA = '" & lenda & "' AND NT_TEKST_DATA = '" & data & "') AND NT_VITI_SHKOLLOR = '" & nrVitiShkollor & "'" & kushti

               If notaEre <> notaEvjeter Then
                  conn.Execute q_res
                  'ex (q_res)
                  active_form.Cgrid1.Text(rreshti, shtylla) = notaEre
               End If

            Else
               MsgBox "Nota e hedhur nuk eshte ne formatin e duhur.", vbInformation, "Modifikim note."
            End If
         End If
      Else
      End If
   

   Case MODIFIKO_AMZA2
      If cikel Then
         strCikli = "1"
      Else
         strCikli = "0"
      End If
      numerAmze = active_form.txtAmzaNo.Text
      emri = active_form.txtEmri.Text
      mbiemri = active_form.txtMbiemri.Text
      Klasa = active_form.cboKlasa.Text
      indeksi = active_form.cboIndeksi.Text
      atesia = active_form.txtAtesia.Text
      memesia = active_form.txtMemesia.Text
      seksi = active_form.cboSeksi.Text
      nrVitiShkollor = active_form.cboVitiShkollor.Text
      Vendlindja = active_form.txtVendlindja.Text
      Datelindja = active_form.txtDatelindja.Value
      nrVitiShkollor = active_form.cboVitiShkollor.Text
      'ugjet = False
      q_res1 = "SELECT AMZA_NR_AMZA FROM TBL_AMZA WHERE AMZA_CIKLI = " & strCikli & " AND AMZA_NR_AMZA = '" & numerAmze & "'"
      tmp.Open q_res1, conn, adOpenDynamic, adLockOptimistic
      If Not tmp.EOF And numerAmze <> nrAmzaModifiko Then
         MsgBox "Numri i ri i amzes i dhene nga ju ekziston tashme ne amze." & Chr(10) & "Jepni nxenesit nje numer te ri.", vbInformation, "Modifikimi i gjeneraliteteve."
         active_form.txtAmzaNo.Text = ""
         active_form.txtAmzaNo.SetFocus

         Exit Sub
      End If
      tmp.Close
      q_res1 = "SELECT * FROM TBL_AMZA WHERE AMZA_CIKLI = " & strCikli & " AND AMZA_EMRI = '" & emri & "' AND AMZA_MBIEMRI = '" & mbiemri & "' AND AMZA_ATESIA = '" & atesia & "'"
      tmp.Open q_res1, conn, adOpenDynamic, adLockOptimistic
      If Not tmp.EOF And (emri <> emerModifiko Or mbiemri <> mbiemerModifiko Or atesia <> atesiaModifiko) Then
         MsgBox "Ekziston nje nxenes me keto te dhena ne amze." & Chr(10) & "Rishikoni perseri te dhenat.", vbInformation, "Modifikimi i gjeneraliteteve."
         active_form.txtEmri.Text = ""
         active_form.txtMbiemri.Text = ""
         active_form.txtAtesia.Text = ""
         active_form.txtEmri.SetFocus
         Exit Sub
      End If
      tmp.Close
      If nrVitiShkollor <> nrVitiShkollorModifiko Then
         q_res1 = "SELECT * FROM TBL_KLASA WHERE KL_NR_AMZA = '" & numerAmze & "' AND KL_CIKLI = " & strCikli & " AND KL_VITI_SHKOLLOR = '" & nrVitiShkollor & "'"
         tmp.Open q_res1, conn, adOpenDynamic, adLockOptimistic
         If Not tmp.EOF Then
            MsgBox "Nxenesi eshte regjistruar nje here per vitin shkollor te dhene nga ju" & Chr(10) & "keshtu qe ju nuk mund ta regjistroni perseri ne kete vit shkollor ." & Chr(10) & " Jepni nje vit shkollor te ri.", vbInformation, "Modifikimi i gjeneraliteteve."
            Exit Sub
         End If
         tmp.Close
      End If

      If nrVitiShkollor = nrVitiShkollorModifiko Then
         If Klasa <> klasaModifiko Then
            q_res1 = "SELECT * FROM TBL_NOTA WHERE NT_NR_AMZA = '" & numerAmze & "' AND NT_CIKLI = " & strCikli & " AND NT_VITI_SHKOLLOR = '" & nrVitiShkollor & "'"
            tmp.Open q_res1, conn, adOpenDynamic, adLockOptimistic
            If Not tmp.EOF Then
               MsgBox "Nxenesi eshte regjistruar nje here per vitin shkollor te dhene nga ju" & Chr(10) & "ne klasen  " & klasaModifiko & "  keshtu qe ju nuk mund t'i ndryshoni klasen ." & Chr(10) & "Jepni nxenesit klasen ekzistuese.", vbInformation, "Modifikimi i gjeneraliteteve."
               tmp.Close
               Exit Sub
            End If
            tmp.Close
         End If
      End If

      
      
      
      q_res = " UPDATE TBL_AMZA SET AMZA_NR_AMZA = '" & numerAmze & "', AMZA_EMRI = '" & emri & "', AMZA_MBIEMRI = '" & mbiemri & "', AMZA_KLASA = '" & Klasa & "', AMZA_INDEKSI = '" & indeksi & "', AMZA_VITI_SHKOLLOR = '" & nrVitiShkollor & "', AMZA_ATESIA = '" & atesia & "', AMZA_MEMESIA = '" & memesia & "', AMZA_SEKSI = '" & seksi & "', AMZA_VENDLINDJA = '" & Vendlindja & "', AMZA_DATELINDJA = '" & Datelindja & "' WHERE (AMZA_NR_AMZA = '" & nrAmzaModifiko & "' AND AMZA_CIKLI = " & strCikli & ")"
      'q_res = " UPDATE TBL_AMZA SET AMZA_NR_AMZA = '" & numerAmze & "', AMZA_EMRI = '" & emri & "',  AMZA_MBIEMRI = '" & mbiemri & "', AMZA_KLASA = '" & Klasa & "', AMZA_ATESIA = '" & atesia & "', AMZA_MEMESIA = '" & memesia & "', AMZA_SEKSI = '" & seksi & "' WHERE (AMZA_NR_AMZA = '" & nrAmza & "' AND AMZA_CIKLI = " & strCikli & ")"
      conn.Execute q_res
      q_res = " UPDATE TBL_KLASA SET KL_NR_AMZA = '" & numerAmze & "' WHERE (KL_NR_AMZA = '" & nrAmzaModifiko & "' AND KL_CIKLI = " & strCikli & ")"
      conn.Execute q_res
      q_res = " UPDATE TBL_NOTA SET NT_NR_AMZA = '" & numerAmze & "' WHERE (NT_NR_AMZA = '" & nrAmzaModifiko & "' AND NT_CIKLI = " & strCikli & ")"
      conn.Execute q_res
      q_res = " UPDATE TBL_SJELLJE SET SJ_NR_AMZA = '" & numerAmze & "' WHERE (SJ_NR_AMZA = '" & nrAmzaModifiko & "' AND SJ_CIKLI = " & strCikli & ")"
      conn.Execute q_res

      If nrVitiShkollor = nrVitiShkollorModifiko And Klasa = klasaModifiko And indeksi <> indeksiModifiko Then
         q_res = "UPDATE TBL_NOTA SET NT_INDEKSI = '" & indeksi & "' WHERE NT_NR_AMZA = '" & numerAmze & "' AND NT_CIKLI = " & strCikli & " AND NT_VITI_SHKOLLOR = '" & nrVitiShkollor & "'"
         conn.Execute q_res
         q_res = "UPDATE TBL_KLASA SET KL_INDEKSI = '" & indeksi & "' WHERE KL_NR_AMZA = '" & numerAmze & "' AND KL_CIKLI = " & strCikli & " AND KL_VITI_SHKOLLOR = '" & nrVitiShkollor & "'"
         conn.Execute q_res
      End If
      
      If nrVitiShkollorModifiko <> nrVitiShkollor Then
         'nqs viti shkollor ndryshon hiqet nxenesi nga klasa e meparshme dhe vendoset ne klasen e re per vitin e ri shkollor
         q_res = "DELETE FROM TBL_KLASA WHERE KL_NR_AMZA = '" & numerAmze & "' AND KL_CIKLI = " & strCikli & " AND KL_VITI_SHKOLLOR = '" & nrVitiShkollorModifiko & "'"
         conn.Execute q_res
         q_res = "INSERT INTO TBL_KLASA ( KL_NR_AMZA, KL_CIKLI, KL_VITI_SHKOLLOR, KL_KLASA, KL_INDEKSI) VALUES ('" & numerAmze & "', " & strCikli & ", '" & nrVitiShkollor & "', '" & Klasa & "', '" & indeksi & "')"
         conn.Execute q_res
      Else
         q_res = "UPDATE TBL_KLASA SET KL_KLASA = '" & Klasa & "', KL_INDEKSI = '" & indeksi & "' WHERE KL_NR_AMZA = '" & numerAmze & "' AND KL_CIKLI = " & strCikli & " AND KL_VITI_SHKOLLOR = '" & nrVitiShkollor & "'"
         conn.Execute q_res
      End If
      
      'Pasqyrimi i modifikimeve te gjeneraliteteteve ne programin per pagesat e nxenesve
      q_res = "if  exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[TBL_REGJISTRIMI]') and " & _
            " OBJECTPROPERTY(id, N'IsUserTable') = 1) BEGIN" & _
            " UPDATE TBL_REGJISTRIMI SET REG_NR_AMZA = '" + numerAmze + "', REG_VITI_SHKOLLOR = '" + nrVitiShkollor + "', " & _
            " REG_KLASA = '" + Klasa + "' WHERE (REG_NR_AMZA = '" + nrAmzaModifiko + "' AND REG_CIKLI = " + strCikli + " AND REG_VITI_SHKOLLOR = '" + nrVitiShkollorModifiko + "') " & _
            " UPDATE TBL_PAGESA SET PG_NR_AMZA = '" + numerAmze + "', PG_VITI_SHKOLLOR = '" + nrVitiShkollor + "' " & _
            " WHERE (PG_NR_AMZA = '" + nrAmzaModifiko + "' AND PG_CIKLI = " + strCikli + " AND PG_VITI_SHKOLLOR = '" + nrVitiShkollorModifiko + "') END"
      conn.Execute q_res

      MsgBox "Te dhenat e modifikuara u regjistruan.", vbInformation, "Konsultimi i amzes."
      active_form.txtAmzaNo.Text = ""
      active_form.txtEmri.Text = ""
      active_form.txtMbiemri.Text = ""
      active_form.cboKlasa.Text = ""
      active_form.cboIndeksi.Text = ""
      active_form.txtAtesia.Text = ""
      active_form.txtMemesia.Text = ""
      active_form.cboSeksi.Text = ""
      active_form.txtVendlindja.Text = ""
      'active_form.txtDatelindja.Text = ""
      active_form.cboVitiShkollor.Text = ""
      active_form.txtMbulo.Visible = True
   Case LENDE_KLASA_KLIK

      active_form.lista2.Clear
      active_form.listaNumri.Clear
      If active_form.optMesme.Value Then
         strCikli = "1"
      Else
         strCikli = "0"
      End If
      Klasa = active_form.cboKlasa.List(active_form.cboKlasa.ListIndex)
      nrVitiShkollor = active_form.cboVitiShkollor.Text
      Dim Gabim As Boolean
      Gabim = False
      Dim vitFillimi As Integer
      vitFillimi = CInt(Mid(nrVitiShkollor, 1, 4))
      If (Klasa = "9") Then
            If (strCikli = "1" And vitFillimi >= 2008) Then
                Gabim = True
                MsgBox "Për këtë vit shkollor cikli i mesëm nuk ka klasë 9!", vbInformation, "Konfigurimi i lëndëve"
            End If
            If (strCikli = "0" And vitFillimi <= 2007) Then
                Gabim = True
                MsgBox "Për këtë vit shkollor shkolla tetëvjeçare nuk ka klasë 9!", vbInformation, "Konfigurimi i lëndëve"
            End If
      End If
      If Gabim Then
        Exit Sub
      End If
      If nrVitiShkollor <> "" Then
         If strCikli = "1" Then
            q_res = " SELECT LN_EMRI, LN_NUMRI_MESME FROM TBL_LENDA WHERE ( LN_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND LN_KLASA" & Klasa & " = 1 ) ORDER BY LN_NUMRI_MESME"
         Else
            q_res = " SELECT LN_EMRI, LN_NUMRI_ULET FROM TBL_LENDA WHERE ( LN_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND LN_KLASA" & Klasa & " = 1 ) ORDER BY LN_NUMRI_ULET"
         End If
         tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
         Do While Not tmp.EOF
            active_form.lista2.AddItem tmp.Fields("LN_EMRI")
            If strCikli = "1" Then
               active_form.listaNumri.AddItem tmp.Fields("LN_NUMRI_MESME")
            Else
               active_form.listaNumri.AddItem tmp.Fields("LN_NUMRI_ULET")
            End If

            tmp.MoveNext
         Loop
         tmp.Close
      Else
         MsgBox " Jepni vitin shkollor para se te zgjidhni klasen", vbInformation, "Konfigurimi i lendeve."
      End If

   Case LENDE_VITI_SHKOLLOR_KLIK

      If active_form.cboKlasa.Text <> "" Then
         active_form.lista2.Clear

         Klasa = active_form.cboKlasa.List(active_form.cboKlasa.ListIndex)
         nrVitiShkollor = active_form.cboVitiShkollor.Text
         If nrVitiShkollor <> "" Then
            q_res = " SELECT LN_EMRI FROM TBL_LENDA WHERE ( LN_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND LN_KLASA" & Klasa & " = 1 )"
            tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
            Do While Not tmp.EOF
               active_form.lista2.AddItem tmp.Fields("LN_EMRI")
               tmp.MoveNext
            Loop
            tmp.Close

         End If
      End If

      strCikli = ""
      nrVitiShkollor = active_form.cboVitiShkollor.Text
      If active_form.optMesme.Value Then
         strCikli = "1"
      End If
      If active_form.optTetevjecare.Value Then
         strCikli = "0"
      End If
      If strCikli <> "" And nrVitiShkollor <> "" Then
         q_res = "SELECT NTA_EMRI_PROVIMI FROM TBL_NOTAAMZ WHERE  NTA_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND NTA_CIKLI = " & strCikli
         tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
         Do While Not tmp.EOF

            active_form.listaProvimet.AddItem tmp.Fields("NTA_EMRI_PROVIMI")
            tmp.MoveNext
         Loop
         tmp.Close
      End If

   Case KALO_KLASE_KERKO

      active_form.listaAmza.Clear
      active_form.lista1.Clear
      nrVitiShkollor = active_form.cboVitiShkollor.Text
      Klasa = active_form.cboKlasa.Text
      indeksi = active_form.cboIndeksi.Text
      q_res = "SELECT AMZA_NR_AMZA,AMZA_EMRI, AMZA_MBIEMRI FROM TBL_AMZA,TBL_KLASA WHERE (AMZA_NR_AMZA = KL_NR_AMZA AND AMZA_CIKLI = KL_CIKLI ) AND ( KL_KLASA = '" & Klasa & "' AND KL_INDEKSI = '" & indeksi & "' AND KL_VITI_SHKOLLOR = '" & nrVitiShkollor & "') AND AMZA_LARGUAR = 0 ORDER BY AMZA_EMRI, AMZA_MBIEMRI"
      tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
      Do While Not tmp.EOF
         active_form.lista1.AddItem tmp.Fields("AMZA_EMRI") & "     " & tmp.Fields("AMZA_MBIEMRI")
         active_form.listaAmza.AddItem tmp.Fields("AMZA_NR_AMZA")
         tmp.MoveNext
      Loop
      tmp.Close

   Case MERR_AMZA_E_MESME

      q_res = "SELECT a.LN_EMRI FROM ( SELECT DISTINCT LN_EMRI, LN_NUMRI_MESME FROM TBL_LENDA " & _
            " WHERE (LN_KLASA10 = 1 OR LN_KLASA11 = 1 OR LN_KLASA12 = 1) OR (LN_KLASA9 = 1 AND CONVERT(INT,SUBSTRING(LN_VITI_SHKOLLOR,0,4)) <= 2007)) " & _
            " as a ORDER BY a.LN_NUMRI_MESME "
      a = 0
      tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
      Do While Not tmp.EOF
         a = a + 1
         tmp.MoveNext
      Loop
      tmp.Close
      'a = a + 2
      inicializoGridaAmzaEmesme active_form.gridaKlasat, active_form.gridaAmza, active_form.gridaLendet, active_form.gridaLabel, a
      active_form.gridaLabel.row = 0
      active_form.gridaLabel.col = 0
      active_form.gridaLabel.ColWidth(0) = 800
      active_form.gridaLabel.Text = "Klasa"
      active_form.gridaLabel.col = 1
      active_form.gridaLabel.ColWidth(1) = 1200
      active_form.gridaLabel.Text = "Viti shkollor"
      i = 0
      active_form.gridaLendet.row = 0
      tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
      Do While Not tmp.EOF
         active_form.gridaLendet.col = i

         active_form.gridaLendet.Text = tmp.Fields("LN_EMRI")
         i = i + 1
         tmp.MoveNext
      Loop
      tmp.Close

   Case MERR_AMZA_TETEVJECARE

      q_res = "SELECT   a.LN_EMRI FROM ( SELECT DISTINCT LN_EMRI, LN_NUMRI_ULET FROM TBL_LENDA " & _
            " WHERE (LN_KLASA1 = 1 OR LN_KLASA2 = 1 OR LN_KLASA3 = 1 OR LN_KLASA4 = 1 " & _
            " OR LN_KLASA5 = 1 OR LN_KLASA6 = 1 OR LN_KLASA7 = 1 OR LN_KLASA8 = 1) Or (LN_KLASA9 = 1 AND CONVERT(INT,SUBSTRING(LN_VITI_SHKOLLOR,1,4)) >=2008 ) ) as " & _
            " a ORDER BY a.LN_NUMRI_ULET "
      a = 0
      tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
      Do While Not tmp.EOF
         a = a + 1
         tmp.MoveNext
      Loop
      tmp.Close
      'a = a + 2
      inicializoGridaTetevjecare active_form.gridaKlasat, active_form.gridaAmza, active_form.gridaLendet, active_form.gridaLabel, a
      active_form.gridaLabel.row = 0
      active_form.gridaLabel.col = 0
      active_form.gridaLabel.ColWidth(0) = 800
      active_form.gridaLabel.Text = "Klasa"
      active_form.gridaLabel.col = 1
      active_form.gridaLabel.ColWidth(1) = 1200
      active_form.gridaLabel.Text = "Viti shkollor"
      i = 0
      active_form.gridaAmza.row = 0
      tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
      Do While Not tmp.EOF
         active_form.gridaLendet.col = i

         active_form.gridaLendet.Text = tmp.Fields("LN_EMRI")
         i = i + 1
         tmp.MoveNext
      Loop
      tmp.Close

   Case SHFAQ_KLASAT_STATISTIKA
      Dim nrM              As Integer
      Dim nrF              As Integer
      nrVitiShkollor = active_form.cboVitiShkollor.Text
      q_res1 = "SELECT DISTINCT KL_KLASA, KL_INDEKSI FROM TBL_KLASA, TBL_AMZA WHERE (AMZA_NR_AMZA = KL_NR_AMZA AND AMZA_CIKLI = KL_CIKLI ) AND KL_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND AMZA_LARGUAR = 0"

      i = 0
      tmp.Open q_res1, conn, adOpenDynamic, adLockOptimistic
      If Not tmp.EOF Then
         Do While Not tmp.EOF
            i = i + 1
            tmp.MoveNext
         Loop

         tmp.Close
      Else
         tmp.Close

      End If

      ' active_form.gridaKlasat.Rows = i + 1
      inicializoGridaNxenesit active_form.gridaKlasat, active_form.gridaLabelKlasat, i
      q_res = "SELECT KL_KLASA,KL_INDEKSI,KL_VITI_SHKOLLOR,AMZA_SEKSI FROM TBL_AMZA, TBL_KLASA WHERE (AMZA_NR_AMZA = KL_NR_AMZA AND AMZA_CIKLI = KL_CIKLI) AND KL_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND AMZA_LARGUAR = 0 ORDER BY KL_KLASA, KL_INDEKSI"
      tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
      If tmp.EOF Then
         'MsgBox "Per kete vit shkollor nuk ka asnje klase te regjistruar.", vbInformation, "Statistika per nxenesit."
         Exit Sub
      End If
      tmp.MoveFirst
      Klasa = tmp.Fields("KL_KLASA")
      indeksi = tmp.Fields("KL_INDEKSI")
      nrM = 0
      nrF = 0
      If tmp.Fields("AMZA_SEKSI") = "mashkull" Then
         nrM = nrM + 1
      Else
         nrF = nrF + 1
      End If
      i = 0
      tmp.MoveNext
      Do While Not tmp.EOF
         If Klasa = tmp.Fields("KL_KLASA") And indeksi = tmp.Fields("KL_INDEKSI") Then
            If tmp.Fields("AMZA_SEKSI") = "mashkull" Then
               nrM = nrM + 1
            Else
               nrF = nrF + 1
            End If
         Else
            If Klasa = "10" Or Klasa = "11" Or Klasa = "12" Then
               'active_form.lista1.AddItem Trim(Klasa) & "-" & (indeksi) & Space(35) & str(nrM) & Space(35) & str(nrF)
               active_form.gridaKlasat.row = i
               active_form.gridaKlasat.col = 0
               active_form.gridaKlasat.CellAlignment = 1
               active_form.gridaKlasat.Text = Klasa & "-" & indeksi
               active_form.gridaKlasat.col = 1
               active_form.gridaKlasat.Text = str(nrM)
               active_form.gridaKlasat.col = 2
               active_form.gridaKlasat.Text = str(nrF)
               i = i + 1

            Else
               'active_form.lista1.AddItem "  " & Trim(Klasa) & "-" & (indeksi) & Space(35) & str(nrM) & Space(35) & str(nrF)
               active_form.gridaKlasat.row = i
               active_form.gridaKlasat.col = 0
               active_form.gridaKlasat.CellAlignment = 1
               active_form.gridaKlasat.Text = "  " & Klasa & "-" & indeksi
               active_form.gridaKlasat.col = 1
               active_form.gridaKlasat.Text = str(nrM)
               active_form.gridaKlasat.col = 2
               active_form.gridaKlasat.Text = str(nrF)
               i = i + 1
            End If
            Klasa = tmp.Fields("KL_KLASA")
            indeksi = tmp.Fields("KL_INDEKSI")
            nrM = 0
            nrF = 0
            If tmp.Fields("AMZA_SEKSI") = "mashkull" Then
               nrM = nrM + 1
            Else
               nrF = nrF + 1
            End If
         End If
         tmp.MoveNext
      Loop
      If Klasa = "10" Or Klasa = "11" Or Klasa = "12" Then
         'active_form.lista1.AddItem Trim(Klasa) & "-" & (indeksi) & Space(35) & str(nrM) & Space(35) & str(nrF)
         active_form.gridaKlasat.row = i
         active_form.gridaKlasat.col = 0
         active_form.gridaKlasat.CellAlignment = 1
         active_form.gridaKlasat.Text = Klasa & "-" & indeksi
         active_form.gridaKlasat.col = 1
         active_form.gridaKlasat.Text = str(nrM)
         active_form.gridaKlasat.col = 2
         active_form.gridaKlasat.Text = str(nrF)
         i = i + 1
      Else
         'active_form.lista1.AddItem "  " & Trim(Klasa) & "-" & (indeksi) & Space(35) & str(nrM) & Space(35) & str(nrF)
         active_form.gridaKlasat.row = i
         active_form.gridaKlasat.col = 0
         active_form.gridaKlasat.CellAlignment = 1
         active_form.gridaKlasat.Text = "  " & Klasa & "-" & indeksi
         active_form.gridaKlasat.col = 1
         active_form.gridaKlasat.Text = str(nrM)
         active_form.gridaKlasat.col = 2
         active_form.gridaKlasat.Text = str(nrF)
         i = i + 1
      End If
      tmp.Close

   Case SHFAQ_MESATARE_NXENESIT

      i = active_form.gridaKlasat.RowSel
      active_form.gridaKlasat.col = 0
      active_form.gridaKlasat.row = i
      s3 = Trim(active_form.gridaKlasat.Text)
      's3 = Trim(active_form.lista1.List(active_form.lista1.ListIndex))
      'Klasa = ktheKlase(s3)
      ' indeksi = ktheIndeksi(s3)
      nrVitiShkollor = active_form.cboVitiShkollor.Text
      Klasa = ""
      ugjet = False
      j = 1
      Do While Not ugjet
         If Mid(s3, j, 1) <> "-" Then
            Klasa = Klasa & Mid(s3, j, 1)
            j = j + 1
         Else
            ugjet = True
         End If
      Loop

      indeksi = Mid(s3, j + 1)
      cikli = ktheCiklin(Klasa, CStr(nrVitiShkollor))
      If cikli Then
         strCikli = "1"
      Else
         strCikli = "0"
      End If

      Dim opsioni          As String
      opsioni = active_form.cboOptions.Text
      If opsioni = "Semestri I" Then
         q_res = "SELECT NT_NR_AMZA, NT_VLERESIMI FROM TBL_NOTA WHERE NT_NR_AMZA IN ( SELECT KL_NR_AMZA FROM TBL_KLASA, TBL_AMZA WHERE ( AMZA_NR_AMZA = KL_NR_AMZA AND AMZA_CIKLI = KL_CIKLI) AND (KL_KLASA = '" & Klasa & "' AND KL_INDEKSI = '" & indeksi & "' AND KL_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND AMZA_LARGUAR = 0)) AND NT_SEMESTRI1 = 1   ORDER BY NT_NR_AMZA"
      ElseIf opsioni = "SemestriII" Then
         q_res = "SELECT NT_NR_AMZA, NT_VLERESIMI FROM TBL_NOTA WHERE NT_NR_AMZA IN ( SELECT KL_NR_AMZA FROM TBL_KLASA, TBL_AMZA WHERE (AMZA_NR_AMZA = KL_NR_AMZA AND AMZA_CIKLI = KL_CIKLI ) AND (KL_KLASA = '" & Klasa & "' AND KL_INDEKSI = '" & indeksi & "') AND ( KL_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND AMZA_LARGUAR = 0))  AND NT_SEMESTRI2 = 1  ORDER BY NT_NR_AMZA"
      Else
         q_res = "SELECT NT_NR_AMZA, NT_VLERESIMI FROM TBL_NOTA WHERE NT_NR_AMZA IN ( SELECT KL_NR_AMZA FROM TBL_KLASA, TBL_AMZA WHERE (AMZA_NR_AMZA = KL_NR_AMZA AND AMZA_CIKLI = KL_CIKLI ) AND (KL_KLASA = '" & Klasa & "' AND KL_INDEKSI = '" & indeksi & "') AND ( KL_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND AMZA_LARGUAR = 0)) AND NT_VJETORE = 1  ORDER BY NT_NR_AMZA"
      End If
      q_res1 = "SELECT AMZA_NR_AMZA,AMZA_EMRI,AMZA_MBIEMRI,AMZA_LARGUAR,AMZA_CIKLI, KL_NR_AMZA, KL_CIKLI, KL_KLASA, KL_INDEKSI, KL_VITI_SHKOLLOR FROM TBL_AMZA, TBL_KLASA WHERE ( AMZA_NR_AMZA = KL_NR_AMZA AND AMZA_CIKLI = KL_CIKLI ) AND (  KL_KLASA= '" & Klasa & "' AND KL_INDEKSI = '" & indeksi & "' AND KL_VITI_SHKOLLOR = '" & nrVitiShkollor & "') AND AMZA_LARGUAR = 0 ORDER BY AMZA_EMRI, AMZA_MBIEMRI "
      i = 0
      tmp.Open q_res1, conn, adOpenDynamic, adLockOptimistic
      Do While Not tmp.EOF
         i = i + 1
         tmp.MoveNext
      Loop
      tmp.Close

      inicializoGridaMestarjaNxenesit active_form.gridaNxenesit, active_form.gridaNxenesitLabel, i

      i = 0

      tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
      Do While Not tmp.EOF
         i = i + 1
         tmp.MoveNext
      Loop
      tmp.Close
      If i = 0 Then
         recset.Open q_res1, conn, adOpenDynamic, adLockOptimistic
         j = 0
         Do While Not recset.EOF
            emri = recset.Fields("AMZA_EMRI")
            mbiemri = recset.Fields("AMZA_MBIEMRI")
            'active_form.lista2.AddItem emri
            'active_form.lista3.AddItem mbiemri
            'active_form.lista4.AddItem "---"
            active_form.gridaNxenesit.row = j
            active_form.gridaNxenesit.col = 0
            active_form.gridaNxenesit.Text = emri
            active_form.gridaNxenesit.col = 1
            active_form.gridaNxenesit.Text = mbiemri
            active_form.gridaNxenesit.col = 2
            active_form.gridaNxenesit.CellAlignment = 1
            active_form.gridaNxenesit.Text = "  ---"
            j = j + 1
            recset.MoveNext
         Loop
         recset.Close
      Else
         Dim strmes           As String
         Dim sum              As Double
         Dim mes              As Double

         'Dim nrLendet As Integer

         tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
         recset.Open q_res1, conn, adOpenDynamic, adLockOptimistic
         j = 0
         Do While Not recset.EOF
            emri = recset.Fields("AMZA_EMRI")
            mbiemri = recset.Fields("AMZA_MBIEMRI")
            numerAmze = recset.Fields("AMZA_NR_AMZA")
            sum = 0
            i = 0
            tmp.MoveFirst

            Do While Not tmp.EOF
               If numerAmze = tmp.Fields("NT_NR_AMZA") Then
                  sum = sum + Val(tmp.Fields("NT_VLERESIMI"))
                  i = i + 1
               End If
               tmp.MoveNext
            Loop

            If i = 0 Then
               strmes = "---"
            Else
               mes = sum / i
               strmes = str(mes)
            End If

            'active_form.lista2.AddItem emri
            'active_form.lista3.AddItem mbiemri
            'active_form.lista4.AddItem strmes
            active_form.gridaNxenesit.row = j
            active_form.gridaNxenesit.col = 0
            active_form.gridaNxenesit.Text = emri
            active_form.gridaNxenesit.col = 1
            active_form.gridaNxenesit.Text = mbiemri
            active_form.gridaNxenesit.col = 2
            active_form.gridaNxenesit.CellAlignment = 1
            If strmes = "4" Or strmes = "5" Or strmes = "6" Or strmes = "7" Or strmes = "8" Or strmes = "9" Or strmes = "10" Then
               active_form.gridaNxenesit.Text = strmes
            ElseIf IsNumeric(strmes) Then
               active_form.gridaNxenesit.Text = Format(strmes, ".0")
            Else
                active_form.gridaNxenesit.Text = strmes
            End If
            recset.MoveNext
            j = j + 1
         Loop
         recset.Close
         tmp.Close
      End If

   Case SHFAQ_KLASAT_MESATARE
      'Dim nr               As Integer
      Dim nra              As Integer
      Dim nrpa             As Integer
      'Dim strmes           As String
      opsioni = active_form.cboOptions.Text
      nrVitiShkollor = active_form.cboVitiShkollor.Text
      q_res = "SELECT DISTINCT KL_KLASA, KL_INDEKSI FROM TBL_KLASA WHERE KL_VITI_SHKOLLOR = '" & nrVitiShkollor & "' ORDER BY KL_KLASA, KL_INDEKSI "
      Dim numerKlasash     As Integer
      numerKlasash = 0
      tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
      Do While Not tmp.EOF
         numerKlasash = numerKlasash + 1
         tmp.MoveNext
      Loop
      tmp.Close
      inicializoGridaKlasat active_form.gridaKlasat, active_form.gridaLabelKlasat, numerKlasash
      active_form.gridaNr.Rows = numerKlasash
      q_res1 = "SELECT NT_KLASA, NT_INDEKSI, NT_VLERESIMI,NT_MOMENTALES1, NT_MOMENTALES2, NT_MUNGESE_PA, NT_MUNGESE_ME, NT_SEMESTRI1, NT_SEMESTRI2, NT_VJETORE FROM TBL_NOTA WHERE NT_VITI_SHKOLLOR = '" & nrVitiShkollor & "'  ORDER BY NT_KLASA, NT_INDEKSI "
      tmp.Open q_res1, conn, adOpenDynamic, adLockOptimistic
      i = 0
      Dim C As Long
      Do While Not tmp.EOF
         C = C + 1
         tmp.MoveNext
      Loop
      tmp.Close
      If C = 0 Then
         recset.Open q_res, conn, adOpenDynamic, adLockOptimistic
         j = 0
         Do While Not recset.EOF
            Klasa = recset.Fields("KL_KLASA")
            indeksi = recset.Fields("KL_INDEKSI")
            If Klasa = "10" Or Klasa = "11" Or Klasa = "12" Then
               ' active_form.lista1.AddItem Klasa & "-" & indeksi & Space(30) & "0" & Space(30) & "0" & Space(30) & "--"
               active_form.gridaKlasat.row = j
               active_form.gridaKlasat.col = 0
               active_form.gridaKlasat.CellAlignment = 1
               active_form.gridaKlasat.Text = Klasa & "-" & indeksi
               active_form.gridaKlasat.col = 1
               active_form.gridaKlasat.Text = Space(10) & " 0"
               active_form.gridaKlasat.col = 2
               active_form.gridaKlasat.Text = Space(10) & " 0"
               active_form.gridaKlasat.col = 3
               active_form.gridaKlasat.Text = Space(10) & "---"
            Else
               ' active_form.lista1.AddItem "  " & Klasa & "-" & indeksi & Space(30) & "0" & Space(30) & "--"
               active_form.gridaKlasat.row = j
               active_form.gridaKlasat.col = 0
               active_form.gridaKlasat.CellAlignment = 1
               active_form.gridaKlasat.Text = "  " & Klasa & "-" & indeksi
               active_form.gridaKlasat.col = 1
               active_form.gridaKlasat.Text = Space(10) & " 0"
               active_form.gridaKlasat.col = 2
               active_form.gridaKlasat.Text = Space(10) & " 0"
               active_form.gridaKlasat.col = 3
               active_form.gridaKlasat.Text = Space(10) & "---"
            End If
            j = j + 1
            recset.MoveNext
         Loop
         recset.Close
      Else
         If opsioni = "Semestri I" Then
            q_res1 = "SELECT NT_KLASA, NT_INDEKSI, NT_VLERESIMI,NT_MOMENTALES1, NT_MUNGESE_PA, NT_MUNGESE_ME, NT_SEMESTRI1 FROM TBL_NOTA WHERE NT_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND ((NT_MOMENTALES1 = 1 AND ( NT_MUNGESE_PA = 1 OR NT_MUNGESE_ME = 1)) OR NT_SEMESTRI1 = 1)  ORDER BY NT_KLASA, NT_INDEKSI "
            tmp.Open q_res1, conn, adOpenDynamic, adLockOptimistic
            If tmp.EOF Then
               tmp.Close
               recset.Open q_res, conn, adOpenDynamic, adLockOptimistic
               j = 0
               Do While Not recset.EOF
                  Klasa = recset.Fields("KL_KLASA")
                  indeksi = recset.Fields("KL_INDEKSI")
                  If Klasa = "10" Or Klasa = "11" Or Klasa = "12" Then
                     ' active_form.lista1.AddItem Klasa & "-" & indeksi & Space(30) & "0" & Space(30) & "0" & Space(30) & "--"
                     active_form.gridaKlasat.row = j
                     active_form.gridaKlasat.col = 0
                     active_form.gridaKlasat.CellAlignment = 1
                     active_form.gridaKlasat.Text = Klasa & "-" & indeksi
                     active_form.gridaKlasat.col = 1
                     active_form.gridaKlasat.Text = Space(10) & " 0"
                     active_form.gridaKlasat.col = 2
                     active_form.gridaKlasat.Text = Space(10) & " 0"
                     active_form.gridaKlasat.col = 3
                     active_form.gridaKlasat.Text = Space(10) & "---"
                  Else
                     ' active_form.lista1.AddItem "  " & Klasa & "-" & indeksi & Space(30) & "0" & Space(30) & "--"
                     active_form.gridaKlasat.row = j
                     active_form.gridaKlasat.col = 0
                     active_form.gridaKlasat.CellAlignment = 1
                     active_form.gridaKlasat.Text = "  " & Klasa & "-" & indeksi
                     active_form.gridaKlasat.col = 1
                     active_form.gridaKlasat.Text = Space(10) & " 0"
                     active_form.gridaKlasat.col = 2
                     active_form.gridaKlasat.Text = Space(10) & " 0"
                     active_form.gridaKlasat.col = 3
                     active_form.gridaKlasat.Text = Space(10) & "---"
                  End If
                  j = j + 1
                  recset.MoveNext
               Loop
               recset.Close
            Else
               recset.Open q_res, conn, adOpenDynamic, adLockOptimistic
               j = 0
               Do While Not recset.EOF
                  Klasa = recset.Fields("KL_KLASA")
                  indeksi = recset.Fields("KL_INDEKSI")
                  sum = 0
                  nra = 0
                  nrpa = 0
                  nr = 0
                  tmp.MoveFirst

                  Do While Not tmp.EOF
                     If tmp.Fields("NT_KLASA") = Klasa And tmp.Fields("NT_INDEKSI") = indeksi Then
                        If tmp.Fields("NT_VLERESIMI") = "m" Then
                           If tmp.Fields("NT_MUNGESE_PA") = True Then
                              nrpa = nrpa + 1
                           Else
                              nra = nra + 1
                           End If
                        Else
                           If tmp.Fields("NT_SEMESTRI1") = True Then
                              sum = sum + Val(tmp.Fields("NT_VLERESIMI"))
                              nr = nr + 1
                           End If
                        End If
                     End If
                     tmp.MoveNext
                  Loop
                  If nr = 0 Then
                     strmes = "---"
                  Else
                     mes = sum / nr
                     strmes = str(mes)
                  End If
                  If Klasa = "10" Or Klasa = "11" Or Klasa = "12" Then
                     ' active_form.lista1.AddItem Klasa & "-" & indeksi & Space(25) & str(nrpa) & Space(30) & str(nra) & Space(35) & strmes
                     active_form.gridaKlasat.row = j
                     active_form.gridaKlasat.col = 0
                     active_form.gridaKlasat.CellAlignment = 1
                     active_form.gridaKlasat.Text = Klasa & "-" & indeksi
                     active_form.gridaKlasat.col = 1
                     active_form.gridaKlasat.Text = Space(10) & str(nrpa)
                     active_form.gridaKlasat.col = 2
                     active_form.gridaKlasat.Text = Space(10) & str(nra)
                     active_form.gridaKlasat.col = 3
                     If strmes = "4" Or strmes = "5" Or strmes = "6" Or strmes = "7" Or strmes = "8" Or strmes = "9" Or strmes = "10" Then
                        active_form.gridaKlasat.Text = Space(5) & strmes
                     ElseIf IsNumeric(strmes) Then
                        active_form.gridaKlasat.Text = Space(5) & Format(strmes, ".0")
                     Else
                        active_form.gridaLendet.Text = strmes
                     End If
                  Else
                     'active_form.lista1.AddItem "  " & Klasa & "-" & indeksi & Space(25) & str(nrpa) & Space(30) & str(nra) & Space(35) & strmes
                     active_form.gridaKlasat.row = j
                     active_form.gridaKlasat.col = 0
                     active_form.gridaKlasat.CellAlignment = 1
                     active_form.gridaKlasat.Text = "  " & Klasa & "-" & indeksi
                     active_form.gridaKlasat.col = 1
                     active_form.gridaKlasat.Text = Space(10) & str(nrpa)
                     active_form.gridaKlasat.col = 2
                     active_form.gridaKlasat.Text = Space(10) & str(nra)
                     active_form.gridaKlasat.col = 3
                     Dim strmesDbl As Double
                     If strmes = "4" Or strmes = "5" Or strmes = "6" Or strmes = "7" Or strmes = "8" Or strmes = "9" Or strmes = "10" Then
                        active_form.gridaKlasat.Text = Space(10) & strmes
                     ElseIf IsNumeric(strmes) Then
                        strmesDbl = CDbl(strmes)
                        ' Per shkak te nje bug te windows bejme kete kontroll
                        If strmes > 100 Then
                            strmesDbl = strmesDbl / 100
                        ElseIf strmes > 10 Then
                            strmesDbl = strmesDbl / 10
                        End If
                        active_form.gridaKlasat.Text = Space(10) & Round(strmesDbl, 2)
                    Else
                        active_form.gridaLendet.Text = strmes
                    End If
                  End If
                  recset.MoveNext
                  j = j + 1
               Loop
               recset.Close
               tmp.Close
            End If
         ElseIf opsioni = "Semestri II" Then
            q_res1 = "SELECT NT_KLASA, NT_INDEKSI, NT_VLERESIMI,NT_MOMENTALES1, NT_MUNGESE_PA, NT_MUNGESE_ME, NT_SEMESTRI2 FROM TBL_NOTA WHERE NT_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND ((NT_MOMENTALES2 = 1 AND ( NT_MUNGESE_PA = 1 OR NT_MUNGESE_ME = 1)) OR NT_SEMESTRI2 = 1)  ORDER BY NT_KLASA, NT_INDEKSI "
            tmp.Open q_res1, conn, adOpenDynamic, adLockOptimistic
            If tmp.EOF Then
               tmp.Close
               recset.Open q_res, conn, adOpenDynamic, adLockOptimistic
               j = 0
               Do While Not recset.EOF
                  Klasa = recset.Fields("KL_KLASA")
                  indeksi = recset.Fields("KL_INDEKSI")
                  If Klasa = "10" Or Klasa = "11" Or Klasa = "12" Then
                     ' active_form.lista1.AddItem Klasa & "-" & indeksi & Space(30) & "0" & Space(30) & "0" & Space(30) & "--"
                     active_form.gridaKlasat.row = j
                     active_form.gridaKlasat.col = 0
                     active_form.gridaKlasat.CellAlignment = 1
                     active_form.gridaKlasat.Text = Klasa & "-" & indeksi
                     active_form.gridaKlasat.col = 1
                     active_form.gridaKlasat.Text = Space(10) & " 0"
                     active_form.gridaKlasat.col = 2
                     active_form.gridaKlasat.Text = Space(10) & " 0"
                     active_form.gridaKlasat.col = 3
                     active_form.gridaKlasat.Text = Space(10) & "---"
                  Else
                     ' active_form.lista1.AddItem "  " & Klasa & "-" & indeksi & Space(30) & "0" & Space(30) & "--"
                     active_form.gridaKlasat.row = j
                     active_form.gridaKlasat.col = 0
                     active_form.gridaKlasat.CellAlignment = 1
                     active_form.gridaKlasat.Text = "  " & Klasa & "-" & indeksi
                     active_form.gridaKlasat.col = 1
                     active_form.gridaKlasat.Text = Space(10) & " 0"
                     active_form.gridaKlasat.col = 2
                     active_form.gridaKlasat.Text = Space(10) & " 0"
                     active_form.gridaKlasat.col = 3
                     active_form.gridaKlasat.Text = Space(10) & "---"
                  End If
                  j = j + 1
                  recset.MoveNext
               Loop
               recset.Close
            Else
               recset.Open q_res, conn, adOpenDynamic, adLockOptimistic
               j = 0
               Do While Not recset.EOF
                  Klasa = recset.Fields("KL_KLASA")
                  indeksi = recset.Fields("KL_INDEKSI")
                  sum = 0
                  nra = 0
                  nrpa = 0
                  nr = 0
                  tmp.MoveFirst

                  Do While Not tmp.EOF
                     If tmp.Fields("NT_KLASA") = Klasa And tmp.Fields("NT_INDEKSI") = indeksi Then
                        If tmp.Fields("NT_VLERESIMI") = "m" Then
                           If tmp.Fields("NT_MUNGESE_PA") = True Then
                              nrpa = nrpa + 1
                           Else
                              nra = nra + 1
                           End If
                        Else
                           If tmp.Fields("NT_SEMESTRI2") = True Then
                              sum = sum + Val(tmp.Fields("NT_VLERESIMI"))
                              nr = nr + 1
                           End If
                        End If
                     End If
                     tmp.MoveNext
                  Loop
                  If nr = 0 Then
                     strmes = "---"
                  Else
                     mes = sum / nr
                     strmes = str(mes)
                  End If
                  If Klasa = "10" Or Klasa = "11" Or Klasa = "12" Then
                     'active_form.lista1.AddItem Klasa & "-" & indeksi & Space(25) & str(nrpa) & Space(30) & str(nra) & Space(35) & strmes
                     active_form.gridaKlasat.row = j
                     active_form.gridaKlasat.col = 0
                     active_form.gridaKlasat.CellAlignment = 1
                     active_form.gridaKlasat.Text = Klasa & "-" & indeksi
                     active_form.gridaKlasat.col = 1
                     active_form.gridaKlasat.Text = Space(10) & str(nrpa)
                     active_form.gridaKlasat.col = 2
                     active_form.gridaKlasat.Text = Space(10) & str(nra)
                     active_form.gridaKlasat.col = 3
                     If strmes = "4" Or strmes = "5" Or strmes = "6" Or strmes = "7" Or strmes = "8" Or strmes = "9" Or strmes = "10" Then
                        active_form.gridaKlasat.Text = Space(10) & strmes
                     ElseIf IsNumeric(strmes) Then
                        active_form.gridaKlasat.Text = Space(10) & Format(strmes, ".0")
                     Else
                        active_form.gridaLendet.Text = strmes
                     End If
                  Else
                     ' active_form.lista1.AddItem "  " & Klasa & "-" & indeksi & Space(25) & str(nrpa) & Space(30) & str(nra) & Space(35) & strmes
                     active_form.gridaKlasat.row = j
                     active_form.gridaKlasat.col = 0
                     active_form.gridaKlasat.CellAlignment = 1
                     active_form.gridaKlasat.Text = "  " & Klasa & "-" & indeksi
                     active_form.gridaKlasat.col = 1
                     active_form.gridaKlasat.Text = Space(10) & str(nrpa)
                     active_form.gridaKlasat.col = 2
                     active_form.gridaKlasat.Text = Space(10) & str(nra)
                     active_form.gridaKlasat.col = 3
                     If strmes = "4" Or strmes = "5" Or strmes = "6" Or strmes = "7" Or strmes = "8" Or strmes = "9" Or strmes = "10" Then
                        active_form.gridaKlasat.Text = Space(10) & strmes
                     ElseIf IsNumeric(strmes) Then
                        active_form.gridaKlasat.Text = Space(10) & Format(strmes, ".0")
                     Else
                        active_form.gridaLendet.Text = strmes
                     End If
                  End If
                  recset.MoveNext
                  j = j + 1
               Loop
               recset.Close
               tmp.Close
            End If
         Else
            q_res1 = "SELECT NT_KLASA, NT_INDEKSI, NT_VLERESIMI, NT_MUNGESE_PA, NT_MUNGESE_ME, NT_VJETORE FROM TBL_NOTA WHERE NT_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND (NT_MUNGESE_PA = 1 OR NT_MUNGESE_ME = 1 OR NT_VJETORE = 1)  ORDER BY NT_KLASA, NT_INDEKSI "
            tmp.Open q_res1, conn, adOpenDynamic, adLockOptimistic
            If tmp.EOF Then
               tmp.Close
               recset.Open q_res, conn, adOpenDynamic, adLockOptimistic
               j = 0
               Do While Not recset.EOF
                  Klasa = recset.Fields("KL_KLASA")
                  indeksi = recset.Fields("KL_INDEKSI")
                  If Klasa = "10" Or Klasa = "11" Or Klasa = "12" Then
                     ' active_form.lista1.AddItem Klasa & "-" & indeksi & Space(30) & "0" & Space(30) & "0" & Space(30) & "--"
                     active_form.gridaKlasat.row = j
                     active_form.gridaKlasat.col = 0
                     active_form.gridaKlasat.CellAlignment = 1
                     active_form.gridaKlasat.Text = Klasa & "-" & indeksi
                     active_form.gridaKlasat.col = 1
                     active_form.gridaKlasat.Text = Space(10) & " 0"
                     active_form.gridaKlasat.col = 2
                     active_form.gridaKlasat.Text = Space(10) & " 0"
                     active_form.gridaKlasat.col = 3
                     active_form.gridaKlasat.Text = Space(10) & "---"
                  Else
                     ' active_form.lista1.AddItem "  " & Klasa & "-" & indeksi & Space(30) & "0" & Space(30) & "--"
                     active_form.gridaKlasat.row = j
                     active_form.gridaKlasat.col = 0
                     active_form.gridaKlasat.CellAlignment = 1
                     active_form.gridaKlasat.Text = "  " & Klasa & "-" & indeksi
                     active_form.gridaKlasat.col = 1
                     active_form.gridaKlasat.Text = Space(10) & " 0"
                     active_form.gridaKlasat.col = 2
                     active_form.gridaKlasat.Text = Space(10) & " 0"
                     active_form.gridaKlasat.col = 3
                     active_form.gridaKlasat.Text = Space(10) & "---"
                  End If
                  j = j + 1
                  recset.MoveNext
               Loop
               recset.Close
            Else

               recset.Open q_res, conn, adOpenDynamic, adLockOptimistic
               j = 0
               Do While Not recset.EOF
                  Klasa = recset.Fields("KL_KLASA")
                  indeksi = recset.Fields("KL_INDEKSI")
                  sum = 0
                  nra = 0
                  nrpa = 0
                  nr = 0
                  tmp.MoveFirst

                  Do While Not tmp.EOF
                     If tmp.Fields("NT_KLASA") = Klasa And tmp.Fields("NT_INDEKSI") = indeksi Then
                        If tmp.Fields("NT_VLERESIMI") = "m" Then
                           If tmp.Fields("NT_MUNGESE_PA") = True Then
                              nrpa = nrpa + 1
                           Else
                              nra = nra + 1
                           End If
                        Else
                           If tmp.Fields("NT_VJETORE") = True Then
                              sum = sum + Val(tmp.Fields("NT_VLERESIMI"))
                              nr = nr + 1
                           End If
                        End If
                     End If
                     tmp.MoveNext
                  Loop
                  If nr = 0 Then
                     strmes = "---"
                  Else
                     mes = sum / nr
                     strmes = str(mes)
                  End If
                  If Klasa = "10" Or Klasa = "11" Or Klasa = "12" Then
                     ' active_form.lista1.AddItem Klasa & "-" & indeksi & Space(25) & str(nrpa) & Space(30) & str(nra) & Space(35) & strmes
                     active_form.gridaKlasat.row = j
                     active_form.gridaKlasat.col = 0
                     active_form.gridaKlasat.CellAlignment = 1
                     active_form.gridaKlasat.Text = Klasa & "-" & indeksi
                     active_form.gridaKlasat.col = 1
                     active_form.gridaKlasat.Text = Space(10) & str(nrpa)
                     active_form.gridaKlasat.col = 2
                     active_form.gridaKlasat.Text = Space(10) & str(nra)
                     active_form.gridaKlasat.col = 3
                     If strmes = "4" Or strmes = "5" Or strmes = "6" Or strmes = "7" Or strmes = "8" Or strmes = "9" Or strmes = "10" Then
                        active_form.gridaKlasat.Text = Space(10) & strmes
                     ElseIf IsNumeric(strmes) Then
                        active_form.gridaKlasat.Text = Space(10) & Format(strmes, ".0")
                     Else
                        active_form.gridaLendet.Text = strmes
                     End If
                  Else
                     ' active_form.lista1.AddItem "  " & Klasa & "-" & indeksi & Space(25) & str(nrpa) & Space(30) & str(nra) & Space(35) & strmes
                     active_form.gridaKlasat.row = j
                     active_form.gridaKlasat.col = 0
                     active_form.gridaKlasat.CellAlignment = 1
                     active_form.gridaKlasat.Text = "  " & Klasa & "-" & indeksi
                     active_form.gridaKlasat.col = 1
                     active_form.gridaKlasat.Text = Space(10) & str(nrpa)
                     active_form.gridaKlasat.col = 2
                     active_form.gridaKlasat.Text = Space(10) & str(nra)
                     active_form.gridaKlasat.col = 3
                     If strmes = "4" Or strmes = "5" Or strmes = "6" Or strmes = "7" Or strmes = "8" Or strmes = "9" Or strmes = "10" Then
                        active_form.gridaKlasat.Text = Space(10) & strmes
                     ElseIf IsNumeric(strmes) Then
                        active_form.gridaKlasat.Text = Space(10) & Format(strmes, ".0")
                     Else
                        active_form.gridaLendet.Text = strmes
                     End If
                  End If
                  recset.MoveNext
                  j = j + 1
               Loop
               recset.Close
               tmp.Close
            End If
         End If

         i = 0
         'q_res = "SELECT a.KL_KLASA, a.KL_INDEKSI, COUNT(a.KL_KLASA) AS NR_NXENESISH  FROM" & _
         '"  (SELECT KL_NR_AMZA, KL_CIKLI, KL_KLASA, KL_VITI_SHKOLLOR, KL_INDEKSI , AMZA_LARGUAR" & _
         '" FROM TBL_KLASA, TBL_AMZA WHERE ( AMZA_NR_AMZA = KL_NR_AMZA AND AMZA_CIKLI = KL_CIKLI )" & _
         '" AND AMZA_LARGUAR = 0 AND KL_VITI_SHKOLLOR = '" & active_form.cboVitiShkollor.Text & "'" & _
         '" ) as a GROUP BY KL_KLASA, KL_INDEKSI ORDER BY KL_KLASA, KL_INDEKSI "
         
         
         q_res = "SELECT B.KL_KLASA, B.KL_INDEKSI, C.NR_NXENESISH FROM " & _
                " (SELECT a.KL_KLASA, a.KL_INDEKSI, COUNT(a.KL_KLASA) AS NR_NXENESISH  FROM " & _
                " (SELECT KL_NR_AMZA, KL_CIKLI, KL_KLASA, KL_VITI_SHKOLLOR, KL_INDEKSI , " & _
                " AMZA_LARGUAR FROM TBL_KLASA, TBL_AMZA WHERE ( AMZA_NR_AMZA = KL_NR_AMZA AND " & _
                "  AMZA_CIKLI = KL_CIKLI ) AND AMZA_LARGUAR = 0 AND " & _
                 " KL_VITI_SHKOLLOR = '" & active_form.cboVitiShkollor.Text & "' ) " & _
                 " as a GROUP BY KL_KLASA, KL_INDEKSI ) AS C " & _
                " Right Join (SELECT D.KL_KLASA, D.KL_INDEKSI, COUNT(D.KL_KLASA) AS NR_NXENESISH  FROM " & _
                " (SELECT KL_NR_AMZA, KL_CIKLI, KL_KLASA, KL_VITI_SHKOLLOR, KL_INDEKSI , " & _
                " AMZA_LARGUAR FROM TBL_KLASA, TBL_AMZA WHERE ( AMZA_NR_AMZA = KL_NR_AMZA AND " & _
                "  AMZA_CIKLI = KL_CIKLI )AND " & _
                "  KL_VITI_SHKOLLOR = '" & active_form.cboVitiShkollor.Text & "' ) " & _
                "  as D GROUP BY KL_KLASA, KL_INDEKSI ) AS B " & _
                " ON C.KL_KLASA = B.KL_KLASA AND C.KL_INDEKSI = B.KL_INDEKSI " & _
                " ORDER BY B.KL_KLASA, B.KL_INDEKSI "
         
         tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
        Do While Not tmp.EOF
            active_form.gridaNr.row = i
            active_form.gridaNr.col = 0
            active_form.gridaNr.Text = tmp.Fields("KL_KLASA") & "-" & tmp.Fields("KL_INDEKSI")

            active_form.gridaNr.row = i
            active_form.gridaNr.col = 1
            If IsNull(tmp.Fields("NR_NXENESISH")) Then
            active_form.gridaNr.Text = 0
            Else
            active_form.gridaNr.Text = tmp.Fields("NR_NXENESISH")
            End If
            tmp.MoveNext
            i = i + 1
        Loop
         tmp.Close

      End If

   Case SHFAQ_MESATARE_LENDET

      'Dim k                As Integer
      Dim numri_i_rreshtit As Integer
      Dim rreshti_i_klases As String
      opsioni = active_form.cboOptions.Text
      nrVitiShkollor = active_form.cboVitiShkollor.Text

      'rreshti_i_klases = active_form.lista1.List(active_form.lista1.ListIndex)
      numri_i_rreshtit = active_form.gridaKlasat.RowSel
      active_form.gridaKlasat.row = numri_i_rreshtit
      active_form.gridaKlasat.col = 0
      rreshti_i_klases = Trim(active_form.gridaKlasat.Text)
      Klasa = ""
      ugjet = False
      j = 1
      Do While Not ugjet
         If Mid(rreshti_i_klases, j, 1) <> "-" Then
            Klasa = Klasa & Mid(rreshti_i_klases, j, 1)
            j = j + 1
         Else
            ugjet = True
         End If
      Loop
      cikli = ktheCiklin(Klasa, CStr(nrVitiShkollor))
      If cikli Then
         strCikli = "1"
      Else
         strCikli = "0"
      End If
      Dim gjatesia         As Integer
      indeksi = Mid(rreshti_i_klases, j + 1)
      If strCikli = "1" Then
         q_res = "SELECT LN_EMRI FROM TBL_LENDA WHERE LN_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND LN_KLASA" & Klasa & " = 1  ORDER BY LN_NUMRI_MESME "
      Else
         q_res = "SELECT LN_EMRI FROM TBL_LENDA WHERE LN_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND LN_KLASA" & Klasa & " = 1  ORDER BY LN_NUMRI_ULET "
      End If
      gjatesia = 0
      tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
      Do While Not tmp.EOF
         gjatesia = gjatesia + 1
         tmp.MoveNext
      Loop
      tmp.Close
      inicializoGridaLendet active_form.gridaLendet, active_form.gridaLendetLabel, gjatesia
      '  active_form.gridaLendet.Rows = gjatesia + 1
      If opsioni = "Vjetore" Then
         q_res1 = "SELECT NT_VLERESIMI,NT_EMRI_LENDA  FROM TBL_NOTA WHERE (NT_KLASA = '" & Klasa & "' AND NT_INDEKSI = '" & indeksi & "') AND (NT_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND NT_VJETORE = 1)"
      ElseIf opsioni = "Semestri I" Then
         q_res1 = "SELECT NT_VLERESIMI,NT_EMRI_LENDA  FROM TBL_NOTA WHERE (NT_KLASA = '" & Klasa & "' AND NT_INDEKSI = '" & indeksi & "') AND (NT_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND NT_SEMESTRI1 = 1)"
      Else
         q_res1 = "SELECT NT_VLERESIMI,NT_EMRI_LENDA  FROM TBL_NOTA WHERE (NT_KLASA = '" & Klasa & "' AND NT_INDEKSI = '" & indeksi & "') AND (NT_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND NT_SEMESTRI2 = 1)"
      End If
      tmp.Open q_res1, conn, adOpenDynamic, adLockOptimistic
      i = 0
      Do While Not tmp.EOF
         i = i + 1
         tmp.MoveNext
      Loop
      tmp.Close

      If i = 0 Then
         recset.Open q_res, conn, adOpenDynamic, adLockOptimistic
         j = 0
         Do While Not recset.EOF
            'active_form.lista2.AddItem recset.Fields("LN_EMRI")
            'active_form.lista3.AddItem "---"
            active_form.gridaLendet.row = j
            active_form.gridaLendet.col = 0
            active_form.gridaLendet.Text = recset.Fields("LN_EMRI")
            active_form.gridaLendet.col = 1
            active_form.gridaLendet.CellAlignment = 1
            active_form.gridaLendet.Text = "---"
            recset.MoveNext
            j = j + 1
         Loop
         recset.Close
      Else
         recset.Open q_res, conn, adOpenDynamic, adLockOptimistic
         tmp.Open q_res1, conn, adOpenDynamic, adLockOptimistic
         k = 0
         Do While Not recset.EOF
            lenda = recset.Fields("LN_EMRI")
            sum = 0
            nr = 0
            tmp.MoveFirst
            Do While Not tmp.EOF
               If tmp.Fields("NT_EMRI_LENDA") = lenda Then
                  sum = sum + Val(tmp.Fields("NT_VLERESIMI"))
                  nr = nr + 1
               End If
               tmp.MoveNext
            Loop
            If nr = 0 Then
               strmes = "---"
            Else
               mes = sum / nr
               strmes = str(mes)
            End If

            ' active_form.lista2.AddItem lenda
            ' active_form.lista3.AddItem strmes
            active_form.gridaLendet.row = k
            active_form.gridaLendet.col = 0
            active_form.gridaLendet.Text = lenda
            active_form.gridaLendet.col = 1
            active_form.gridaLendet.CellAlignment = 1
            If strmes = "4" Or strmes = "5" Or strmes = "6" Or strmes = "7" Or strmes = "8" Or strmes = "9" Or strmes = "10" Then
               active_form.gridaLendet.Text = strmes
            ElseIf IsNumeric(strmes) Then
               active_form.gridaLendet.Text = CStr(Round(strmes, 2))
            Else
                active_form.gridaLendet.Text = strmes
            End If
            recset.MoveNext
            k = k + 1
         Loop
         recset.Close
         tmp.Close
      End If
   

  
  
End Select
End Sub
Public Sub andiNew()
   ' Dim nrAmza As String
   Dim vleresimi        As String
   Dim a                As Integer
   Dim b                As Integer
   Dim q_res2           As String
   Dim s1, s2, q_res, q_res1, q_res11, s3 As String
   Dim numerAmze        As String
   Dim cikli            As Integer
   Dim strCikli         As String
   Dim emri             As String
   Dim emriProvimi             As String
   Dim mbiemri          As String
   Dim nrVitiShkollor   As String
   Dim Klasa            As String
   Dim indeksi          As String
   Dim atesia           As String
   Dim memesia          As String
   Dim Vendlindja       As String
   Dim Datelindja       As Date
   Dim seksi            As String
   Dim Vrejtje          As String
   Dim mungeseMeS1, mungeseMeS2, mungesePaS1, MungesePaS2 As Integer
   Dim conn             As New ADODB.Connection
   
   With conn
      .Provider = "Microsoft.Access.OLEDB.10.0"
      .Properties("Data Provider").Value = "SQLOLEDB"
      .Properties("Data Source").Value = EmerServer + "\SHM"
      .Properties("User ID").Value = "sa"
      .Properties("Password").Value = "vision"
      .Properties("Initial Catalog").Value = "DProgNec"
      If (conn.state = 1) Then
        .Close
      End If
      .Open
   End With

   'conn.ConnectionString = getConnectionString()
   'conn.Open
   Dim tmp              As New ADODB.Recordset
   Dim tmp11            As New ADODB.Recordset
   Dim tmp111           As New ADODB.Recordset
   Select Case Me.actionName
      Case RUAJ_PERDORUES_I_RI
         Dim ndodhet_perdorues As Boolean
         Dim statuskuo        As String
         ndodhet_perdorues = False
         s1 = active_form.Text2.Text
         s2 = active_form.Text1.Text
         statuskuo = active_form.cboStatusi.Text
         q_res1 = "SELECT USER_NAME FROM TBL_USERS"
         tmp.Open q_res1, conn, adOpenDynamic, adLockOptimistic
         Do While Not tmp.EOF And Not ndodhet_perdorues
            If tmp.Fields("USER_NAME") = s1 Then
               ndodhet_perdorues = True
            End If
            tmp.MoveNext
         Loop
         tmp.Close
         'nuk mund te shtohet vizitor nqs nuk ka administrator
         If (Not active_form.cboStatusi.Text = "Administrator") Then
            q_res1 = "SELECT USER_NAME FROM TBL_USERS WHERE USER_STATUSI = 'Administrator'"
            tmp.Open q_res1, conn, adOpenDynamic, adLockOptimistic
            If (tmp.EOF) Then
                MsgBox "Nuk mund të krijoni asnjë lloj perdoruesi tjetër nqs nuk keni një administrator në program." + Chr(10) + "Krijoni më parë një përdorues administrator!", vbCritical, "Konfigurimi i perdoruesve."
                Exit Sub
            End If
         End If
         
         If Not ndodhet_perdorues Then
            If active_form.cboStatusi.Text = "Vizitor" Then
                q_res = " Insert Into TBL_USERS (USER_NAME, USER_PASSWORD, USER_STATUSI,USER_WORKSTATION) values " & "('" & s1 & "'," & "'" & s2 & "', '" & statuskuo & "','" + active_form.cboWorkstation.Text + "')"
            Else
                q_res = " Insert Into TBL_USERS (USER_NAME, USER_PASSWORD, USER_STATUSI) values " & "('" & s1 & "'," & "'" & s2 & "', '" & statuskuo & "')"
            End If
            ex (q_res)
            
            MsgBox "Perdoruesi i ri u regjistrua.", vbInformation, "Konfigurimi i perdoruesve."
            active_form.Text2.Text = ""
            active_form.Text1.Text = ""
            active_form.Text3.Text = ""
            active_form.cboStatusi.ListIndex = -1
            active_form.cboWorkstation.ListIndex = -1
         Else
            MsgBox "Ekziston nje perdorues me kete emer.Jepni nje emer te ri", vbInformation, "Konfigurimi i perdoruesve."
            active_form.Text2.Text = ""
         End If

      Case MODIFIKO_PERDORUES
         emri = active_form.txtUserName.List(active_form.txtUserName.ListIndex)
         s1 = active_form.txtPerdorues.Text
         s2 = active_form.txtPassword.Text
         If (active_form.cboStatusiModifiko.Text = "Vizitor") Then
            q_res = "UPDATE TBL_USERS SET USER_NAME = '" & s1 & "' , USER_PASSWORD = '" & s2 & "',USER_STATUSI = '" + active_form.cboStatusiModifiko.Text + "',USER_WORKSTATION = '" + active_form.cboWorkstationModifiko.Text + "' WHERE USER_NAME = '" & emri & "'"
         Else
            q_res = "UPDATE TBL_USERS SET USER_NAME = '" & s1 & "' , USER_PASSWORD = '" & s2 & "',USER_STATUSI = '" + active_form.cboStatusiModifiko.Text + "',USER_WORKSTATION = null WHERE USER_NAME = '" & emri & "'"
         End If
         conn.Execute q_res
         
      Case KERKO_PERDORUES
         s1 = active_form.txtUserName

         q_res = "Select  USER_NAME,USER_PASSWORD from TBL_USERS "

         'get_rec (q_res)
         tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
         Do While Not tmp.EOF
            active_form.txtUserName.AddItem tmp.Fields("USER_NAME")
            tmp.MoveNext
         Loop
         tmp.Close
         'active_form.txtPassword.Text = record1.Fields("USER_PASSWORD")
      Case VERIFIKO_PERDORUES
         Dim kodi             As String
         Dim eshte            As Boolean
         Dim perd_statusi, workstation     As String
         emri = active_form.txtUserName.List(active_form.txtUserName.ListIndex)
         kodi = active_form.txtPassword.Text
         eshte = False
         q_res = "Select * from TBL_USERS "
         tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
         Do While Not tmp.EOF And Not eshte

            If emri = tmp.Fields("USER_NAME") And tmp.Fields("USER_PASSWORD") = kodi Then
               eshte = True
               perd_statusi = tmp.Fields("USER_STATUSI")
               If (perd_statusi = "Vizitor") Then
                    If tmp.Fields("USER_WORKSTATION") & "" = "" Then
                    workstation = ""
                    Else
                    workstation = tmp.Fields("USER_WORKSTATION")
                    End If
               End If
            End If
            tmp.MoveNext
         Loop
         tmp.Close
         If Not eshte Then
            MsgBox "Fjalekalimi nuk eshte i sakte.Jepni perseri fjalekalimin. ", vbInformation, "Konfigurimi i perdoruesve."
            active_form.txtPassword.Text = ""
         Else
            active_form.txtPassword.Text = ""
            active_form.txtPerdorues.Visible = True
            active_form.txtPerdorues.Text = emri
            active_form.txtPassword1.Visible = True
            active_form.lbPassword1.Visible = True
            active_form.cmdVerifiko.Visible = False
            active_form.cmdModifikoUser.Visible = True
            active_form.cmdElimino.Visible = True

            Select Case perd_statusi

               Case "Administrator"
                  active_form.cboStatusiModifiko.ListIndex = 0
               Case "SupervizorEmesme"
                  active_form.cboStatusiModifiko.ListIndex = 1
               Case "SupervizorTetevjecare"
                  active_form.cboStatusiModifiko.ListIndex = 2
               Case "Vizitor"
                  active_form.cboStatusiModifiko.ListIndex = 3
               Case Else
                  active_form.cboStatusiModifiko.ListIndex = -1
            End Select

            If statusi = "Administrator" Then
               active_form.cboStatusiModifiko.Visible = True
               active_form.Label5.Visible = True
            End If

            If perd_statusi = "Vizitor" Then
                If (workstation = "") Then
                    active_form.cboWorkstationModifiko.ListIndex = -1
                Else
                    active_form.cboWorkstationModifiko.Text = workstation
                End If
                
                active_form.cboWorkstationModifiko.Visible = True
                active_form.lblWorkstationModifiko.Visible = True
            Else
                active_form.cboWorkstationModifiko.ListIndex = -1
            End If
         End If
      Case EKZISTON_PERDORUES
         q_res = "SELECT * FROM TBL_USERS"
         tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
         If tmp.EOF Then
            Unload frmLogin
            frmMDIMain.show
            statusi = "Administrator"
         Else
            frmLogin.show vbModal

         End If
         tmp.Close
      Case ELIMINO_PERDORUES

         emri = active_form.txtPerdorues.Text
         q_res = "DELETE FROM TBL_USERS WHERE USER_NAME = '" & emri & "'"
         conn.Execute q_res
         MsgBox "Perdoruesi i zgjedhur nga ju u fshi.", vbInformation, "Konfigurimi i perdoruesve."

      Case LOGIMI
         Dim ugjet            As Boolean
         emri = frmLogin.txtUserName.Text
         kodi = frmLogin.txtPassword.Text
         ugjet = False
         Dim emerWorkstation As String
         emerWorkstation = GjejEmerWorkstation
         q_res = "SET CONCAT_NULL_YIELDS_NULL OFF SELECT USER_NAME, USER_PASSWORD, USER_STATUSI,USER_WORKSTATION + '' as USER_WORKSTATION FROM TBL_USERS WHERE USER_NAME = '" & emri & "' AND USER_PASSWORD = '" & kodi & "' SET CONCAT_NULL_YIELDS_NULL ON"
         
         tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
         Dim workstationUser As String
         If Not tmp.EOF Then
            workstationUser = tmp.Fields("USER_WORKSTATION")
         End If
         If tmp.EOF Then
            MsgBox " Emri i perdoruesit ose fjalekalimi mund te mos jene te sakta." & Chr(10) & "Jepni te dhenat perseri.", vbInformation, "Hyrja ne program."
            frmLogin.txtUserName.Text = ""
            frmLogin.txtPassword.Text = ""
            frmLogin.txtUserName.SetFocus
         ElseIf tmp.Fields("USER_STATUSI") = "Vizitor" And UCase(emerWorkstation) <> UCase(workstationUser) Then
            MsgBox "Nuk mund të hyni në program nga ky kompjuter!", vbInformation, "Hyrja ne program."
            frmLogin.txtUserName.Text = ""
            frmLogin.txtPassword.Text = ""
            frmLogin.txtUserName.SetFocus
         Else
            statusi = tmp.Fields("USER_STATUSI")
            Unload frmLogin
            frmMDIMain.show
            frmMDIMain.Enabled = True

         End If
         tmp.Close
      Case MODIFIKO_GJENERALITETE_SHENIME

         If active_form.optMesme.Value Then
            strCikli = "1"
         Else
            strCikli = "0"
         End If
         nrVitiShkollor = active_form.cboVitiShkollor.Text
         If active_form.txtAmzaNo.Text <> "" Then
            numerAmze = active_form.txtAmzaNo.Text
            q_res = "SELECT AMZA_NR_AMZA FROM TBL_AMZA WHERE AMZA_NR_AMZA = '" & numerAmze & "' AND AMZA_CIKLI = " & strCikli
            tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
            If tmp.EOF Then
               MsgBox "Ky numer amze nuk ekziston ." & Chr(10) & "Jepni perseri numrin e amzes.", vbInformation, "Modifikimi i shenimeve"
               tmp.Close
               GoTo fund_shenime
            End If
            tmp.Close
         Else
            emri = active_form.txtEmri.Text
            mbiemri = active_form.txtMbiemri.Text
            q_res = "SELECT AMZA_NR_AMZA FROM TBL_AMZA WHERE AMZA_EMRI = '" & emri & "' AND AMZA_MBIEMRI = '" & mbiemri & "'  AND AMZA_CIKLI = '" & strCikli & "' "
            tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
            If tmp.EOF Then
               MsgBox "Nuk ekziston ndonje nxenes me keto te dhena " & Chr(10) & "ne ciklin e percaktuar prej jush.", vbInformation, "Modifikimi i shenimeve."
               tmp.Close
               GoTo fund_shenime
            End If
            numerAmze = tmp.Fields("AMZA_NR_AMZA")
         End If

         q_res = "SELECT  SJ_SJELLJE FROM TBL_SJELLJE WHERE SJ_NR_AMZA = '" & numerAmze & "' AND SJ_CIKLI = " & strCikli & " AND SJ_VITI_SHKOLLOR = '" & nrVitiShkollor & "'"
         tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
         If tmp.EOF Then
            MsgBox "Per kete vit shkollor nxenesi i percaktuar nga ju nuk ka qene nxenes ne kete shkolle ose ne ciklin e shkolles te zgjedhur nga ju " & Chr(10) & "ose shenimet mbi sjelljen nuk jane hedhur akoma per kete vit shkollor.", vbInformation, "Modifikimi i sjelljeve."
         Else
            active_form.rtbSjellja.Text = tmp.Fields("SJ_SJELLJE")
         End If
         tmp.Close

fund_shenime:

      Case MODIFIKO_SHENIME_SJELLJE
         Dim sjellje_e_modifikuar As String
         Klasa = active_form.cboKlasa.Text
         nrVitiShkollor = active_form.cboVitiShkollor.Text
         cikli = ktheCiklin(Klasa, CStr(nrVitiShkollor))
         If cikli Then
            strCikli = "1"
         Else
            strCikli = "0"
         End If
        ' Ndryshohet ngjyra e tekstit qe sherben per deftesen
        active_form.rtbSjellja.SelStart = 0
        If cikli = 0 Then
            active_form.rtbSjellja.SelLength = 206
        Else
            active_form.rtbSjellja.SelLength = 373
        End If
        active_form.rtbSjellja.SelColor = vbBlue
        active_form.rtbSjellja.SelLength = 0
         
         numerAmze = active_form.listAmza.List(active_form.lista.ListIndex)
         sjellje_e_modifikuar = active_form.rtbSjellja.Text
         q_res = "UPDATE TBL_SJELLJE SET SJ_SJELLJE = '" & sjellje_e_modifikuar & "' WHERE SJ_NR_AMZA = '" & numerAmze & "' AND SJ_CIKLI = " & strCikli & " AND SJ_VITI_SHKOLLOR = '" & nrVitiShkollor & "'"
         conn.Execute q_res
         MsgBox "Shenimet mbi sjelljen e nxenesve ne vitin shkollor te zgjedhur nga ju u modifikuan.", vbInformation, "Modifikimi i shenimeve."
      
      Case JEP_ADRESA_LOGO
         q_res = "SELECT * FROM TBL_PERDORUES"
         tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
         If tmp.EOF Then
            adresaLogo = ""
            emerShkolla = ""
            email = ""
            website = ""
         Else
            If Not IsNull(tmp.Fields("PERD_LOGO")) Then
               adresaLogo = tmp.Fields("PERD_LOGO")
            Else
               adresaLogo = ""
            End If
            If Not IsNull(tmp.Fields("PERD_EMRI_SHKOLLA")) Then
               emerShkolla = tmp.Fields("PERD_EMRI_SHKOLLA")
            Else
               emerShkolla = ""
            End If
            If Not IsNull(tmp.Fields("PERD_WEBSITE")) Then
               website = tmp.Fields("PERD_WEBSITE")
            Else
               website = ""
            End If
            If Not IsNull(tmp.Fields("PERD_EMAIL")) Then
               email = tmp.Fields("PERD_EMAIL")
            Else
               email = ""
            End If
            If Not IsNull(tmp.Fields("PERD_QYTETI")) Then
               qytetiShkolla = tmp.Fields("PERD_QYTETI")
            End If
            If Not IsNull(tmp.Fields("PERD_TELEFONI")) Then
               telefoniShkolla = tmp.Fields("PERD_TELEFONI")
            End If
            If Not IsNull(tmp.Fields("PERD_ADRESA")) Then
               adresaShkolla = tmp.Fields("PERD_ADRESA")
            End If

         End If
         tmp.Close
         Dim fileName         As String
         fileName = Dir(adresaLogo)
         If fileName = "" Then
            adresaLogo = ""
         End If
      Case SHENIME_BUTONI_KLIK

         nrVitiShkollor = active_form.cboVitiShkollor.Text

         If active_form.optMesme.Value Then
            strCikli = "1"
         Else
            strCikli = "0"
         End If
         If active_form.txtAmzaNo.Text <> "" Then
            numerAmze = active_form.txtAmzaNo.Text
            q_res = "SELECT AMZA_NR_AMZA FROM TBL_AMZA WHERE AMZA_NR_AMZA = '" & numerAmze & "' AND AMZA_CIKLI = " & strCikli
            tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
            If tmp.EOF Then
               tmp.Close
               GoTo dil3
            End If
            tmp.Close
         Else
            emri = active_form.txtEmri.Text
            mbiemri = active_form.txtMbiemri.Text
            q_res = "SELECT AMZA_NR_AMZA FROM TBL_AMZA WHERE AMZA_EMRI = '" & emri & "' AND AMZA_MBIEMRI = '" & mbiemri & "'"
            tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
            If tmp.EOF Then
               tmp.Close
               GoTo dil3
            Else

               numerAmze = tmp.Fields("AMZA_NR_AMZA")
               tmp.Close
            End If
         End If

         'Dim sjellje_e_modifikuar As String
         sjellje_e_modifikuar = active_form.rtbSjellja.Text
         q_res = "UPDATE TBL_SJELLJE SET SJ_SJELLJE = '" & sjellje_e_modifikuar & "'"
         conn.Execute q_res
         MsgBox "Shenimet mbi sjelljet per nxenesin e dhene ne vitin shkollor te zgjedhur nga ju u modifikua.", vbInformation, "Modifikimi i shenimeve."
         GoTo dil4
dil3:
         MsgBox "Nxenesi me te dhenat e mesiperme nuk  eshte nxenes i kesaj shkolle ose nuk eshte ne ciklin e dhene nga ju.", vbInformation, "Modifikimi i shenimeve."
dil4:
      Case SHENIME_LISTA_KLIK

         numerAmze = active_form.listAmza.List(active_form.lista.ListIndex)
         Klasa = active_form.cboKlasa.Text
         nrVitiShkollor = active_form.cboVitiShkollor.Text
         cikli = ktheCiklin(Klasa, CStr(nrVitiShkollor))
         If cikli Then
            strCikli = "1"
         Else
            strCikli = "0"
         End If

         q_res = "SELECT SJ_SJELLJE FROM TBL_SJELLJE WHERE SJ_NR_AMZA = '" & numerAmze & "' AND SJ_CIKLI = " & strCikli & " AND SJ_VITI_SHKOLLOR = '" & nrVitiShkollor & "'"
         tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
         If tmp.EOF Then
            MsgBox "Per kete vit shkollor ketij nxenesi nuk i eshte hedhur sjellja .", vbInformation, "Modifikimi i shenimeve."
         Else
            active_form.rtbSjellja.Text = tmp.Fields("SJ_SJELLJE")
            ' Ndryshohet ngjyra e tekstit qe sherben per deftesen
            active_form.rtbSjellja.SelStart = 0
            If cikli = 0 Then
               active_form.rtbSjellja.SelLength = 206
            Else
               active_form.rtbSjellja.SelLength = 373
            End If
            active_form.rtbSjellja.SelColor = vbBlue
            active_form.rtbSjellja.SelLength = 0
         End If
         tmp.Close
         GoTo dil1
dil:
         MsgBox "Nxenesi me te dhenat e mesiperme nuk  eshte nxenes i kesaj shkolle" & Chr(10) & " ose nuk eshte ne ciklin e dhene nga ju.", vbInformation, "Modifikimi i shenimeve."

dil1:
      Case HIDH_INFORMACIONE_NE_FORME

         q_res = "SELECT * FROM TBL_PERDORUES "
         tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
         If Not tmp.EOF Then

            If Not IsNull(tmp.Fields("PERD_EMRI_SHKOLLA")) Then
               active_form.txtEmri.Text = tmp.Fields("PERD_EMRI_SHKOLLA")
            End If
            active_form.txtAdresa.Text = tmp.Fields("PERD_ADRESA")
            active_form.txtWebsite.Text = tmp.Fields("PERD_WEBSITE")
            If Not IsNull(tmp.Fields("PERD_EMAIL")) Then
               active_form.txtEmail.Text = tmp.Fields("PERD_EMAIL")
            End If
            If Not IsNull(tmp.Fields("PERD_QYTETI")) Then
               active_form.txtQyteti.Text = tmp.Fields("PERD_QYTETI")
            End If
            If Not IsNull(tmp.Fields("PERD_RRETHI")) Then
               active_form.txtRrethi.Text = tmp.Fields("PERD_RRETHI")
            End If
            If Not IsNull(tmp.Fields("PERD_TELEFONI")) Then
               active_form.txtTelefoni.Text = tmp.Fields("PERD_TELEFONI")
            End If

         End If
      Case HIDH_INFORMACIONE_NE_DATABASE
         Dim emer_shkolle     As String
         Dim adresa_shkolle   As String
         Dim website_shkolle  As String
         Dim email_shkolle    As String
         Dim rrethi_shkolle   As String
         Dim telefoni_shkolle As String
         Dim qyteti_shkolle   As String

         emer_shkolle = active_form.txtEmri.Text
         adresa_shkolle = active_form.txtAdresa.Text
         website_shkolle = active_form.txtWebsite.Text
         email_shkolle = active_form.txtEmail.Text
         qyteti_shkolle = active_form.txtQyteti.Text
         rrethi_shkolle = active_form.txtRrethi.Text
         telefoni_shkolle = active_form.txtTelefoni.Text
         qyteti_shkolle = active_form.txtQyteti.Text

         q_res = "SELECT * FROM TBL_PERDORUES "
         tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
         If tmp.EOF Then
            tmp.Close
            q_res = "INSERT INTO TBL_PERDORUES (PERD_EMRI_SHKOLLA, PERD_WEBSITE, PERD_EMAIL, PERD_ADRESA, PERD_LOGO, PERD_QYTETI, PERD_RRETHI, PERD_TELEFONI ) VALUES ('" & emer_shkolle & "', '" & website_shkolle & "', '" & email_shkolle & "', '" & adresa_shkolle & "', '" & adresaLogo & "', '" & qyteti_shkolle & "', '" & rrethi_shkolle & "', '" & telefoni_shkolle & "')"

            conn.Execute q_res
            MsgBox "Informacionet mbi shkollen u hodhen.", vbInformation, "Informacion mbi shkollen."
            adresaShkolla = adresa_shkolle
            emerShkolla = emer_shkolle
            website = website_shkolle
            email = email_shkolle
            telefoniShkolla = telefoni_shkolle
            qytetiShkolla = qyteti_shkolle
            rrethiShkolla = rrethi_shkolle
         Else
            tmp.Close
            q_res = "UPDATE TBL_PERDORUES SET PERD_EMRI_SHKOLLA = '" & emer_shkolle & "', PERD_WEBSITE = '" & website_shkolle & "', PERD_EMAIL = '" & email_shkolle & "', PERD_ADRESA = '" & adresa_shkolle & "', PERD_LOGO = '" & adresaLogo & "', PERD_QYTETI = '" & qyteti_shkolle & "' , PERD_RRETHI = '" & rrethi_shkolle & "', PERD_TELEFONI = '" & telefoni_shkolle & "'"
            conn.Execute q_res
            MsgBox "Informacioni mbi shkollen u modifikua.", vbInformation, "Informacion mbi shkollen."
         End If
         adresaShkolla = adresa_shkolle
         emerShkolla = emer_shkolle
         website = website_shkolle
         email = email_shkolle
         qytetiShkolla = qyteti_shkolle
         rrethiShkolla = rrethi_shkolle
         telefoniShkolla = telefoni_shkolle

      Case KERKO_KLASA_E_RE
         active_form.lista2Amza.Clear
         active_form.lista2.Clear
         nrVitiShkollor = active_form.cboVitiShkollor2.Text
         Klasa = active_form.cboKlasa2.Text
         indeksi = active_form.cboIndeksi2.Text
         q_res = "SELECT AMZA_NR_AMZA,AMZA_EMRI, AMZA_MBIEMRI FROM TBL_AMZA,TBL_KLASA WHERE (AMZA_NR_AMZA = KL_NR_AMZA AND AMZA_CIKLI = KL_CIKLI ) AND ( KL_KLASA = '" & Klasa & "' AND KL_INDEKSI = '" & indeksi & "' AND KL_VITI_SHKOLLOR = '" & nrVitiShkollor & "') AND AMZA_LARGUAR = 0 ORDER BY AMZA_EMRI, AMZA_MBIEMRI"
         tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
         Do While Not tmp.EOF
            active_form.lista2.AddItem tmp.Fields("AMZA_EMRI") & "     " & tmp.Fields("AMZA_MBIEMRI")
            active_form.lista2Amza.AddItem tmp.Fields("AMZA_NR_AMZA")
            tmp.MoveNext
         Loop
         tmp.Close
      Case SHFAQ_KLASA_PROVIMET
         Dim i, j             As Integer
         Dim lirimi, mature, ck As Boolean
         lirimi = active_form.optLirimi.Value
         mature = active_form.optMature.Value

         If mature Then
            ck = True
            active_form.Label5.Visible = True
            active_form.Label4.Visible = True
            active_form.Label3.Visible = False
Else:
            ck = False
            active_form.Label3.Visible = True
            active_form.Label4.Visible = True
            active_form.Label5.Visible = False
         End If

         '*** nxjerr numrin e provimeve sipas ciklit
         Dim llojProvimi      As String
         nrVitiShkollor = active_form.cboVitiShkollor.Text
         Klasa = active_form.cboKlasa.Text
         indeksi = active_form.cboIndeksi.Text
         cikli = ktheCiklin(Klasa, CStr(nrVitiShkollor))
         Dim nr_provimesh     As Integer
         If cikli Then
            strCikli = "1"
            llojProvimi = " e matures"
         Else
            strCikli = "0"
            llojProvimi = " e lirimit"
         End If

         q_res = "SELECT NTA_EMRI_PROVIMI FROM TBL_NOTAAMZ WHERE NTA_CIKLI = " & strCikli & " AND NTA_VITI_SHKOLLOR = '" & nrVitiShkollor & "'"
         tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
         If tmp.EOF Then
            scrbar = False
            scrbarver = False
            MsgBox "Per kete vit shkollor nuk jane percaktuar provimet " & llojProvimi & ".", vbExclamation, "Hedhja e notave te provimeve."
            active_form.Label3.Visible = False
            active_form.Label4.Visible = False
            active_form.Label5.Visible = False
            active_form.provimet.Visible = False

            active_form.cboKlasa.Enabled = True
            active_form.cboIndeksi.Enabled = True
            active_form.cboVitiShkollor.Enabled = True

            active_form.optMomentale.Enabled = True
            active_form.optMomentale.Value = False
            active_form.optLirimi.Enabled = True
            active_form.optLirimi.Value = False
            active_form.optMature.Enabled = True
            active_form.optMature.Value = False
            active_form.optDetyreKontrolli.Enabled = True
            active_form.optDetyreKontrolli.Value = False
            active_form.optSemestrale.Enabled = True
            active_form.optSemestrale.Value = False
            active_form.optMungeseMe.Enabled = True
            active_form.optMungeseMe.Value = False
            active_form.optMungesePa.Enabled = True
            active_form.optMungesePa.Value = False
            active_form.optSemestri1.Enabled = True
            active_form.optSemestri1.Value = False
            active_form.optSemestri2.Enabled = True
            active_form.optSemestri2.Value = False
            active_form.optVjetore.Enabled = True
            active_form.optVjetore.Value = False
            active_form.cboKlasa.ListIndex = -1
            active_form.cboIndeksi.ListIndex = -1
            active_form.cmdOK.Enabled = False
            Exit Sub
         Else
            tmp.MoveFirst
            j = 0
            Do While Not tmp.EOF
               j = j + 1
               tmp.MoveNext
            Loop
         End If
         tmp.Close
         nr_provimesh = j

         ' Kontrollon nese kjo klase ekziston
         nrVitiShkollor = active_form.cboVitiShkollor.Text
         Dim Kl               As String
         s2 = active_form.cboIndeksi.Text
         Kl = active_form.cboKlasa.Text
         q_res1 = "SELECT KL_KLASA FROM TBL_KLASA WHERE (KL_KLASA = '" & Kl & "' AND KL_INDEKSI = '" & s2 & "' )"
         tmp.Open q_res1, conn, adOpenDynamic, adLockOptimistic
         If tmp.EOF Then
            scrbar = False
            scrbarver = False
            active_form.Label3.Visible = False
            active_form.Label4.Visible = False
            active_form.Label5.Visible = False
            active_form.provimet.Visible = False
            MsgBox "Klasa e zgjedhur nga ju nuk ekziston ose nuk eshte e rregjistruar ne kete vit shkollor", vbInformation, "Hedhja e notave te provimeve"
            active_form.Label3.Visible = False
            active_form.Label4.Visible = False
            active_form.Label5.Visible = False
            active_form.provimet.Visible = False

            active_form.cboKlasa.Enabled = True
            active_form.cboIndeksi.Enabled = True
            active_form.cboVitiShkollor.Enabled = True

            active_form.optMomentale.Enabled = True
            active_form.optMomentale.Value = False
            active_form.optLirimi.Enabled = True
            active_form.optLirimi.Value = False
            active_form.optMature.Enabled = True
            active_form.optMature.Value = False
            active_form.optDetyreKontrolli.Enabled = True
            active_form.optDetyreKontrolli.Value = False
            active_form.optSemestrale.Enabled = True
            active_form.optSemestrale.Value = False
            active_form.optMungeseMe.Enabled = True
            active_form.optMungeseMe.Value = False
            active_form.optMungesePa.Enabled = True
            active_form.optMungesePa.Value = False
            active_form.optSemestri1.Enabled = True
            active_form.optSemestri1.Value = False
            active_form.optSemestri2.Enabled = True
            active_form.optSemestri2.Value = False
            active_form.optVjetore.Enabled = True
            active_form.optVjetore.Value = False
            active_form.cboKlasa.ListIndex = -1
            active_form.cboIndeksi.ListIndex = -1
            scrbar = False
            scrbarver = False
            active_form.cmdOK.Enabled = False
            Exit Sub
         End If
         tmp.Close

         'nxjerr nxenesit qe ka nje klase
         nrVitiShkollor = active_form.cboVitiShkollor.Text
         s2 = active_form.cboIndeksi.Text
         q_res1 = "SELECT AMZA_EMRI , AMZA_MBIEMRI, AMZA_NR_AMZA, KL_KLASA, KL_INDEKSI, KL_VITI_SHKOLLOR FROM TBL_AMZA, TBL_KLASA  WHERE (AMZA_NR_AMZA = KL_NR_AMZA AND AMZA_CIKLI = KL_CIKLI) AND (KL_KLASA = '" & Klasa & "') AND (KL_INDEKSI = '" & indeksi & "') AND (KL_VITI_SHKOLLOR = '" & nrVitiShkollor & "') AND (AMZA_LARGUAR = 0) ORDER BY AMZA_EMRI, AMZA_MBIEMRI  "
         tmp.Open q_res1, conn, adOpenDynamic, adLockOptimistic
         If tmp.EOF Then
            scrbar = False
            scrbarver = False
            MsgBox "Klasa e zgjedhur nga ju nuk ka asnje nxenes.", vbInformation, "Hedhja e notave te provimeve."
            active_form.Label3.Visible = False
            active_form.Label4.Visible = False
            active_form.Label5.Visible = False
            active_form.provimet.Visible = False

            active_form.cboKlasa.Enabled = True
            active_form.cboIndeksi.Enabled = True
            active_form.cboVitiShkollor.Enabled = True

            active_form.optMomentale.Enabled = True
            active_form.optMomentale.Value = False
            active_form.optLirimi.Enabled = True
            active_form.optLirimi.Value = False
            active_form.optMature.Enabled = True
            active_form.optMature.Value = False
            active_form.optDetyreKontrolli.Enabled = True
            active_form.optDetyreKontrolli.Value = False
            active_form.optSemestrale.Enabled = True
            active_form.optSemestrale.Value = False
            active_form.optMungeseMe.Enabled = True
            active_form.optMungeseMe.Value = False
            active_form.optMungesePa.Enabled = True
            active_form.optMungesePa.Value = False
            active_form.optSemestri1.Enabled = True
            active_form.optSemestri1.Value = False
            active_form.optSemestri2.Enabled = True
            active_form.optSemestri2.Value = False
            active_form.optVjetore.Enabled = True
            active_form.optVjetore.Value = False
            active_form.cboKlasa.ListIndex = -1
            active_form.cboIndeksi.ListIndex = -1
            scrbar = False
            scrbarver = False
            Exit Sub
            active_form.cmdOK.Enabled = False
         Else
            tmp.MoveFirst
            j = 0
            Do While Not tmp.EOF
               j = j + 1
               tmp.MoveNext
            Loop
         End If
         tmp.Close

         nr_nxenesit = j

         Call shfaq_grile2(nr_nxenesit + 1, nr_provimesh + 2)
         q_res = "SELECT NTA_EMRI_PROVIMI FROM TBL_NOTAAMZ WHERE NTA_CIKLI = " & strCikli & " AND NTA_VITI_SHKOLLOR = '" & active_form.cboVitiShkollor.Text & "'"
         tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
         i = 2
         Do While Not tmp.EOF
            i = i + 1
            active_form.provimet.Text(1, i) = tmp.Fields("NTA_EMRI_PROVIMI")
            tmp.MoveNext
         Loop
         tmp.Close
         j = 1
         q_res1 = "SELECT AMZA_EMRI ,AMZA_MBIEMRI, AMZA_NR_AMZA, KL_KLASA, KL_INDEKSI, KL_VITI_SHKOLLOR FROM TBL_AMZA, TBL_KLASA  WHERE (AMZA_NR_AMZA = KL_NR_AMZA AND AMZA_CIKLI = KL_CIKLI) AND (KL_KLASA = '" & Klasa & "') AND (KL_INDEKSI = '" & indeksi & "') AND (KL_VITI_SHKOLLOR = '" & nrVitiShkollor & "') AND (AMZA_LARGUAR = 0) ORDER BY AMZA_EMRI, AMZA_MBIEMRI "
         tmp.Open q_res1, conn, adOpenDynamic, adLockOptimistic
         Do While Not tmp.EOF
            j = j + 1
            active_form.provimet.Text(j, 1) = j
            active_form.provimet.Text(j, 2) = tmp.Fields("AMZA_EMRI") & "     " & tmp.Fields("AMZA_MBIEMRI")
            active_form.Amza.Text(j - 1, 1) = tmp.Fields("AMZA_NR_AMZA")
            'active_form.Mbiemrat.Text(j, 1) = record2.Fields("AMZA_MBIEMRI")
            tmp.MoveNext

         Loop
         tmp.Close
         active_form.cmdOK.Enabled = True
         active_form.Text1.Visible = False
         Dim lenda            As String

         nr_lenda_provimet = nr_provimesh
         nr_nxenesit_provimet = nr_nxenesit

      Case GET_INFO_HEDHJE_NOTA    'gjen lendet qe ben nje klase dhe nxenesit qe ka kjo klase

         'nxjerr lendet qe ben nje klase
         Klasa = active_form.cboKlasa.Text
         nrVitiShkollor = active_form.cboVitiShkollor.Text
         cikli = ktheCiklin(Klasa, CStr(nrVitiShkollor))
         If cikli Then
            strCikli = "1"
         Else
            strCikli = "0"
         End If
         indeksi = active_form.cboIndeksi.Text
         If strCikli = "1" Then
            q_res = "SELECT LN_EMRI FROM TBL_LENDA WHERE LN_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND LN_KLASA" & Klasa & " = 1  ORDER BY LN_NUMRI_MESME "
         Else
            q_res = "SELECT LN_EMRI FROM TBL_LENDA WHERE LN_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND LN_KLASA" & Klasa & " = 1  ORDER BY LN_NUMRI_ULET "
         End If
         tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
         If tmp.EOF Then
            scrbar = False
            scrbarver = False
            MsgBox "Klases se zgjedhur nga ju nuk i eshte hedhur asnje lende" & Chr(10) & "per vitin shkollor perkates", vbInformation, "Hedhja e notave."
            active_form.Label3.Visible = False
            active_form.Label4.Visible = False
            active_form.Label5.Visible = False

            active_form.Cgrid1.Visible = False
            active_form.emrat.Visible = False
            active_form.Text1.Visible = False
            active_form.Cgrid2.Visible = False
            active_form.provimet.Visible = False

            active_form.cboKlasa.ListIndex = -1
            active_form.cboIndeksi.ListIndex = -1
            'active_form.Amze.Visible = False
            active_form.optMomentale.Enabled = True
            active_form.optLirimi.Enabled = True
            active_form.optMature.Enabled = True
            active_form.optDetyreKontrolli.Enabled = True
            active_form.optMomentale.Enabled = True
            active_form.optMungeseMe.Enabled = True
            active_form.optMungesePa.Enabled = True
            active_form.optSemestri1.Enabled = True
            active_form.optSemestri2.Enabled = True
            active_form.optSemestrale.Enabled = True
            active_form.optRiprovim.Enabled = True
            active_form.optVjetore.Enabled = True
            active_form.cmdOK.Enabled = False
            Exit Sub
         Else
            'nr i lendeve

            i = 0
            Do While Not tmp.EOF
               i = i + 1
               tmp.MoveNext

            Loop
            nr_lenda = i
         End If
         tmp.Close

         'nxjerr nxenesit qe ka nje klase
         nrVitiShkollor = active_form.cboVitiShkollor.Text

         q_res1 = "SELECT AMZA_EMRI , AMZA_MBIEMRI, AMZA_NR_AMZA, KL_KLASA, KL_INDEKSI, KL_VITI_SHKOLLOR FROM TBL_AMZA, TBL_KLASA  WHERE (AMZA_NR_AMZA = KL_NR_AMZA AND AMZA_CIKLI = KL_CIKLI) AND (KL_KLASA = '" & Klasa & "') AND (KL_INDEKSI = '" & indeksi & "') AND (KL_VITI_SHKOLLOR = '" & nrVitiShkollor & "') AND (AMZA_LARGUAR = 0) ORDER BY AMZA_EMRI, AMZA_MBIEMRI  "
         tmp.Open q_res1, conn, adOpenDynamic, adLockOptimistic
         If tmp.EOF Then
            scrbar = False
            scrbarver = False
            MsgBox "Klasa e zgjedhur nga ju nuk ka asnje nxenes per vitin shkollor perkates.", vbInformation, "Hedhja e notave."
            active_form.Label3.Visible = False
            active_form.Label4.Visible = False
            active_form.Label5.Visible = False

            active_form.Cgrid1.Visible = False
            active_form.emrat.Visible = False
            active_form.Text1.Visible = False
            active_form.Cgrid2.Visible = False
            active_form.provimet.Visible = False
            active_form.cboKlasa.ListIndex = -1
            active_form.cboIndeksi.ListIndex = -1
            'active_form.Amze.Visible = False
            active_form.optMomentale.Enabled = True
            active_form.optLirimi.Enabled = True
            active_form.optMature.Enabled = True
            active_form.optDetyreKontrolli.Enabled = True
            active_form.optMungeseMe.Enabled = True
            active_form.optMungesePa.Enabled = True
            active_form.optSemestri1.Enabled = True
            active_form.optSemestri2.Enabled = True
            active_form.optVjetore.Enabled = True
            active_form.optRiprovim.Enabled = True
            active_form.optSemestrale.Enabled = True
            active_form.cmdOK.Enabled = False
            Exit Sub
         Else

            j = 0
            Do While Not tmp.EOF
               j = j + 1
               tmp.MoveNext
            Loop
            tmp.Close
         End If
         nr_nxenesit = j
         'nr_nxenesit = record2.RecordCount

         Call shfaq_grile(nr_lenda, nr_nxenesit)
         i = 0
         If strCikli = "1" Then
            q_res = "SELECT LN_EMRI FROM TBL_LENDA WHERE LN_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND LN_KLASA" & Klasa & " = 1  ORDER BY LN_NUMRI_MESME "
         Else
            q_res = "SELECT LN_EMRI FROM TBL_LENDA WHERE LN_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND LN_KLASA" & Klasa & " = 1  ORDER BY LN_NUMRI_ULET "
         End If
         tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic

         Do While Not tmp.EOF
            i = i + 1
            active_form.Cgrid2.Text(1, i) = tmp.Fields("LN_EMRI")
            tmp.MoveNext

         Loop
         tmp.Close
         j = 0
         q_res1 = "SELECT AMZA_EMRI , AMZA_MBIEMRI, AMZA_NR_AMZA, KL_KLASA, KL_INDEKSI, KL_VITI_SHKOLLOR FROM TBL_AMZA, TBL_KLASA  WHERE (AMZA_NR_AMZA = KL_NR_AMZA AND AMZA_CIKLI = KL_CIKLI) AND (KL_KLASA = '" & Klasa & "') AND (KL_INDEKSI = '" & indeksi & "') AND (KL_VITI_SHKOLLOR = '" & nrVitiShkollor & "') AND (AMZA_LARGUAR = 0) ORDER BY AMZA_EMRI, AMZA_MBIEMRI  "
         tmp.Open q_res1, conn, adOpenDynamic, adLockOptimistic
         Do While Not tmp.EOF
            j = j + 1
            active_form.emrat.Text(j, 1) = j
            active_form.emrat.Text(j, 2) = tmp.Fields("AMZA_EMRI") & "     " & tmp.Fields("AMZA_MBIEMRI")
            active_form.Amza.Text(j, 1) = tmp.Fields("AMZA_NR_AMZA")
            'active_form.Mbiemrat.Text(j, 1) = record2.Fields("AMZA_MBIEMRI")
            tmp.MoveNext

         Loop
         tmp.Close
         active_form.cmdOK.Enabled = True
         nr_lenda_momentale = nr_lenda
         nr_nxenesit_momentale = nr_nxenesit
         active_form.emrat.Visible = True
         active_form.Cgrid2.Visible = True
         active_form.Cgrid1.Visible = True
         active_form.Text1.Visible = True

      Case HIDH_PROVIME_NOTA

         Dim nr_hedhje        As Integer
         nr_hedhje = 0
         Dim nota             As String
         'Dim lenda            As String
         Dim data             As String
         Klasa = active_form.cboKlasa.Text
         indeksi = active_form.cboIndeksi.Text
         nrVitiShkollor = active_form.cboVitiShkollor.Text

         If active_form.optMature.Value Then
            strCikli = "1"
         Else
            strCikli = "0"
         End If

         '*** gjen numrin e lendeve
         q_res = "SELECT NTA_EMRI_PROVIMI FROM TBL_NOTAAMZ WHERE NTA_CIKLI = " & strCikli & " AND NTA_VITI_SHKOLLOR = '" & nrVitiShkollor & "'"
         tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
         i = 0
         Do While Not tmp.EOF
            i = i + 1
            tmp.MoveNext
         Loop
         tmp.Close
         nr_lenda_provimet = i

         '**** gjen numrin e nxenesve

         q_res1 = "SELECT AMZA_EMRI , AMZA_MBIEMRI, AMZA_NR_AMZA, KL_KLASA, KL_INDEKSI, KL_VITI_SHKOLLOR FROM TBL_AMZA, TBL_KLASA  WHERE (AMZA_NR_AMZA = KL_NR_AMZA AND AMZA_CIKLI = KL_CIKLI) AND (KL_KLASA = '" & Klasa & "') AND (KL_INDEKSI = '" & indeksi & "') AND (KL_VITI_SHKOLLOR = '" & nrVitiShkollor & "') AND (AMZA_LARGUAR = 0) ORDER BY AMZA_EMRI, AMZA_MBIEMRI  "
         tmp.Open q_res1, conn, adOpenDynamic, adLockOptimistic

         j = 0
         Do While Not tmp.EOF
            j = j + 1
            tmp.MoveNext
         Loop
         tmp.Close
         nr_nxenesit_provimet = j

         '**** KONTROLLON FORMATIN E NOTAVE

         For j = 3 To nr_lenda_provimet + 1
            For i = 2 To nr_nxenesit_provimet + 1
               nota = active_form.provimet.Text(i, j)
               If Not (nota = "" Or nota = "4" Or nota = "5" Or nota = "6" Or nota = "7" Or nota = "8" Or nota = "9" Or nota = "10") Then
                  MsgBox "Nota e hedhur nuk eshte ne formatin e duhur.", vbInformation, "Hedhja e notave."
                  'active_form.provimet.CellSetFocus i, j
                  Exit Sub
               End If
            Next i
         Next j

         q_res = "SELECT NT_NR_AMZA, NT_VLERESIMI, NT_EMRI_LENDA FROM TBL_NOTA WHERE NT_CIKLI = " & strCikli & " AND NT_KLASA = '" & Klasa & "' AND NT_INDEKSI = '" & indeksi & "' AND NT_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND NT_PROVIM = 1"
         tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
         If Not tmp.EOF Then
            For j = 3 To nr_lenda_provimet + 2
               For i = 2 To nr_nxenesit_provimet + 1
                  emri = Trim(active_form.provimet.Text(i, 2))
                  nota = active_form.provimet.Text(i, j)
                  If nota <> "" Then
                     lenda = active_form.provimet.Text(1, j)
                     numerAmze = active_form.Amza.Text(i - 1, 1)
                     data = active_form.Calendar1.Value
                     
                     ugjet = False
                     tmp.MoveFirst
                     Do While Not tmp.EOF And Not ugjet
                        If tmp.Fields("NT_NR_AMZA") = numerAmze And tmp.Fields("NT_EMRI_LENDA") = lenda Then
                           ugjet = True
                        End If
                        tmp.MoveNext
                     Loop
                     If ugjet Then
                        MsgBox "Nota e provimit '" & lenda & "' per nxenesin '" & emri & "' eshte hedhur nje here.", vbInformation, "Hedhja e provimeve."
                        active_form.provimet.Text(i, j) = ""
                     Else
                        q_res1 = "INSERT INTO TBL_NOTA( NT_NR_AMZA, NT_CIKLI, NT_EMRI_LENDA,NT_VLERESIMI, NT_DATA, NT_KLASA, NT_INDEKSI, NT_VITI_SHKOLLOR, NT_PROVIM,NT_TEKST_DATA ) VALUES ('" & numerAmze & "', " & strCikli & ", '" & lenda & "','" & nota & "', '" & konvertoDaten(data) & "', '" & Klasa & "', '" & indeksi & "', '" & nrVitiShkollor & "', 1,'" + DataTekst(data) + "')"
                        conn.Execute q_res1
                        nr_hedhje = nr_hedhje + 1
                     End If
                  End If
               Next i
            Next j
         Else
            For j = 3 To nr_lenda_provimet + 2
               For i = 2 To nr_nxenesit + 1
                  emri = Trim(active_form.provimet.Text(i, 2))
                  nota = active_form.provimet.Text(i, j)
                  If nota <> "" Then
                     lenda = active_form.provimet.Text(1, j)
                     numerAmze = active_form.Amza.Text(i - 1, 1)
                     data = active_form.Calendar1.Value
                     q_res1 = "INSERT INTO TBL_NOTA( NT_NR_AMZA, NT_CIKLI, NT_EMRI_LENDA,NT_VLERESIMI, NT_DATA, NT_KLASA, NT_INDEKSI, NT_VITI_SHKOLLOR, NT_PROVIM ,NT_TEKST_DATA) VALUES ('" & numerAmze & "', " & strCikli & ", '" & lenda & "', '" & nota & "', '" & konvertoDaten(data) & "', '" & Klasa & "', '" & indeksi & "', '" & nrVitiShkollor & "', 1,'" + DataTekst(data) + "')"
                     conn.Execute q_res1
                     nr_hedhje = nr_hedhje + 1
                  End If

               Next i
            Next j
         End If
         tmp.Close
         If nr_hedhje >= 1 Then

            MsgBox "Notat e provimeve u hodhen.", vbInformation, "Hedhja e provimeve."
         End If

         For j = 3 To nr_lenda_provimet + 2
            For i = 2 To nr_nxenesit + 1
               active_form.provimet.Text(i, j) = ""
            Next i
         Next j
         active_form.provimet.Visible = False
         active_form.Label3.Visible = False
         active_form.Label4.Visible = False
         active_form.Label5.Visible = False
      Case PASTRO_KLASE

         Dim klasaEre         As String
         Dim indeksiIri       As String
         Dim nrVitiShkollorIri As String
         nrVitiShkollorIri = active_form.cboVitiShkollor2.Text
         klasaEre = active_form.cboKlasa2.Text
         indeksiIri = active_form.cboIndeksi2.Text

         If active_form.optMesme.Value Then
            strCikli = "1"
         Else
            strCikli = "0"
         End If
         Dim nxenesi          As String
         Dim klasa_nxenesi    As String
         Dim indeksi_nxenesi  As String

         ' q_res1 = "DELETE FROM TBL_KLASA WHERE  KL_KLASA = '" & klasaEre & "' AND KL_INDEKSI = '" & indeksiIri & "' AND KL_VITI_SHKOLLOR = '" & nrVitiShkollorIri & "'"
         ' conn.Execute q_res1

         ' For I = 0 To active_form.lista2Amza.ListCount - 1
         'nxenesi = Trim(active_form.lista2.List(I))

         '   numerAmze = active_form.lista2Amza.List(I)

         '   q_res = "INSERT INTO TBL_KLASA (KL_NR_AMZA, KL_CIKLI, KL_VITI_SHKOLLOR,KL_KLASA, KL_INDEKSI) VALUES ('" & numerAmze & "', " & strcikli & " , '" & nrVitiShkollorIri & "', '" & klasaEre & "', '" & indeksiIri & "')"
         '   conn.Execute q_res
         '   q_res1 = "UPDATE TBL_AMZA SET AMZA_KLASA = '" & klasaEre & "', AMZA_INDEKSI = '" & indeksiIri & "', AMZA_VITI_SHKOLLOR = '" & nrVitiShkollorIri & "' WHERE AMZA_NR_AMZA = '" & numerAmze & "' AND AMZA_CIKLI = " & strcikli
         '  conn.Execute q_res1

         'ex (q_res)
         ' Next I
         i = 0
         a = active_form.lista2.ListCount - 1
         For i = 0 To a
            If active_form.lista2.Selected(i) Then
               numerAmze = active_form.lista2Amza.List(i)
               j = MsgBox("Ju jeni duke fshire nje nxenesin e selektuar nga klasa perkatese." & Chr(10) & "Fshirja e tij do te shoqerohetme fshirjen e te gjitha vleresimeve te marra ne kete klase." & Chr(10) & "Jeni te sigurte qe doni te vazhdoni ?", vbInformation + vbOKCancel, "Konfigurimi  klasave.")
               If j = vbCancel Then
                  Exit Sub
               End If
               q_res = "DELETE FROM TBL_KLASA"
               MsgBox "Klasa u rikonfigurua.", vbInformation, "Konfigurimi i klasave."
               active_form.lista1.Clear
               active_form.lista2.Clear
               active_form.cboKlasa.ListIndex = -1
               active_form.cboIndeksi.ListIndex = -1
               active_form.cboVitiShkollor.ListIndex = -1
               active_form.cboKlasa2.ListIndex = -1
               active_form.cboIndeksi2.ListIndex = -1
               active_form.cboVitiShkollor2.ListIndex = -1
            End If
         Next i
      
      
      Case HIDH_NOTA_MOMENTALE

         Dim momentales1      As String
         Dim momentales2      As String
         Dim semestrale11      As String
         Dim semestrale21      As String
         Dim detyrekontrolli  As String
         Dim vjetore1          As String
         Dim riprovim1         As String
         Dim mungeseMe        As String
         Dim mungesePa        As String
         nr_hedhje = 0
         Klasa = active_form.cboKlasa.Text
         indeksi = active_form.cboIndeksi.Text
         nrVitiShkollor = active_form.cboVitiShkollor.Text
         cikli = ktheCiklin(Klasa, CStr(nrVitiShkollor))
         If cikli Then
            strCikli = "1"
         Else
            strCikli = "0"
         End If

         '*** kontrollon formatin e notave

         ugjet = False
         For j = 1 To nr_lenda_momentale
            For i = 1 To nr_nxenesit_momentale
               nota = active_form.Cgrid1.Text(i, j)
               If Not (nota = "" Or nota = "4" Or nota = "5" Or nota = "6" Or nota = "7" Or nota = "8" Or nota = "9" Or nota = "10" Or nota = "m" Or nota = "M") Then
                  MsgBox "Notat nuk jane te gjitha ne formatin e duhur.", vbCritical, "Hedhja e notave."
                  active_form.Cgrid1.CellSetFocus i, j
                  Exit Sub
               End If
            Next i
         Next j

         Dim data_tekst       As String
         Dim data_konvertuar  As String

         data_tekst = DataTekst(active_form.Calendar1.Value)
         data_konvertuar = konvertoDaten(active_form.Calendar1.Value)
         data = active_form.Calendar1.Value
         'q_res = "SELECT NT_NR_AMZA,NT_EMRI_LENDA, NT_VLERESIMI,NT_TEKST_DATA FROM TBL_NOTA WHERE (NT_KLASA = '" & Klasa & "' AND NT_INDEKSI = '" & indeksi & "' AND NT_VITI_SHKOLLOR = '" & nrVitiShkollor & "' ) AND ( NT_MOMENTALES1 = TRUE OR NT_MOMENTALES2 = TRUE OR NT_DETYREKONTROLLI = TRUE OR NT_MUNGESE_PA = TRUE OR NT_MUNGESE_ME = TRUE ) AND NT_TEKST_DATA = '" & data_tekst & "'"
         'tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
        For j = 1 To nr_lenda_momentale
            For i = 1 To nr_nxenesit_momentale
                If vektor(i, j).llojNote = "M" And vektor(i, j).semester = "S1" Then
                    momentales1 = "1"
                    momentales2 = "0"
                    detyrekontrolli = "0"
                    semestrale11 = "0"
                    semestrale21 = "0"
                    mungeseMe = "0"
                    mungesePa = "0"
                    vjetore1 = "0"
                    riprovim1 = "0"
                ElseIf vektor(i, j).llojNote = "M" And vektor(i, j).semester = "S2" Then
                    momentales1 = "0"
                    momentales2 = "1"
                    detyrekontrolli = "0"
                    semestrale11 = "0"
                    semestrale21 = "0"
                    mungeseMe = "0"
                    mungesePa = "0"
                    vjetore1 = "0"
                    riprovim1 = "0"
                ElseIf vektor(i, j).llojNote = "D" And vektor(i, j).semester = "S1" Then
                    momentales1 = "1"
                    momentales2 = "0"
                    detyrekontrolli = "1"
                    semestrale11 = "0"
                    semestrale21 = "0"
                    mungeseMe = "0"
                    mungesePa = "0"
                    vjetore1 = "0"
                    riprovim1 = "0"
                ElseIf vektor(i, j).llojNote = "D" And vektor(i, j).semester = "S2" Then
                    momentales1 = "0"
                    momentales2 = "1"
                    detyrekontrolli = "1"
                    semestrale11 = "0"
                    semestrale21 = "0"
                    mungeseMe = "0"
                    mungesePa = "0"
                    vjetore1 = "0"
                    riprovim1 = "0"
                ElseIf vektor(i, j).llojNote = "Mm" And vektor(i, j).semester = "S1" Then
                    momentales1 = "1"
                    momentales2 = "0"
                    detyrekontrolli = "0"
                    semestrale11 = "0"
                    semestrale21 = "0"
                    mungeseMe = "1"
                    mungesePa = "0"
                    vjetore1 = "0"
                    riprovim1 = "0"
                ElseIf vektor(i, j).llojNote = "Mp" And vektor(i, j).semester = "S1" Then
                    momentales1 = "1"
                    momentales2 = "0"
                    detyrekontrolli = "0"
                    semestrale11 = "0"
                    semestrale21 = "0"
                    mungeseMe = "0"
                    mungesePa = "1"
                    vjetore1 = "0"
                    riprovim1 = "0"
                ElseIf vektor(i, j).llojNote = "Mm" And vektor(i, j).semester = "S2" Then
                    momentales1 = "0"
                    momentales2 = "1"
                    mungeseMe = "1"
                    mungesePa = "0"
                    detyrekontrolli = "0"
                    semestrale11 = "0"
                    semestrale21 = "0"
                    vjetore1 = "0"
                    riprovim1 = "0"
                ElseIf vektor(i, j).llojNote = "Mp" And vektor(i, j).semester = "S2" Then
                    momentales1 = "0"
                    momentales2 = "1"
                    semestrale11 = "0"
                    semestrale21 = "0"
                    mungesePa = "1"
                    detyrekontrolli = "0"
                    mungeseMe = "0"
                    vjetore1 = "0"
                    riprovim1 = "0"
                ElseIf vektor(i, j).llojNote = "S" And vektor(i, j).semester = "S1" Then
                    momentales1 = "0"
                    momentales2 = "0"
                    detyrekontrolli = "0"
                    semestrale11 = "1"
                    semestrale21 = "0"
                    vjetore1 = "0"
                    mungeseMe = "0"
                    mungesePa = "0"
                    riprovim1 = "0"
                ElseIf vektor(i, j).llojNote = "S" And vektor(i, j).semester = "S2" Then
                    momentales1 = "0"
                    momentales2 = "0"
                    detyrekontrolli = "0"
                    semestrale11 = "0"
                    semestrale21 = "1"
                    vjetore1 = "0"
                    mungeseMe = "0"
                    mungesePa = "0"
                    riprovim1 = "0"
                ElseIf vektor(i, j).semester = "V" Then
                    momentales1 = "0"
                    momentales2 = "0"
                    detyrekontrolli = "0"
                    semestrale11 = "0"
                    semestrale21 = "0"
                    vjetore1 = "1"
                    mungeseMe = "0"
                    mungesePa = "0"
                    riprovim1 = "0"
                ElseIf vektor(i, j).semester = "R" Then
                    momentales1 = "0"
                    momentales2 = "0"
                    detyrekontrolli = "0"
                    semestrale11 = "0"
                    semestrale21 = "0"
                    vjetore1 = "0"
                    mungeseMe = "0"
                    mungesePa = "0"
                    riprovim1 = "1"
                End If
                data_tekst = DataToTekst(vektor(i, j).date)
                data_konvertuar = konvertoDaten(CStr(vektor(i, j).date))
                ' Nese nota eshte perfundimtare: vjetore, semestrale ose riprovim
                If vjetore1 = "1" Or semestrale11 = "1" Or semestrale21 = "1" Or riprovim1 = "1" Then
                    If vjetore1 = "1" Then
                        q_res = "SELECT NT_NR_AMZA, NT_EMRI_LENDA, NT_VLERESIMI, NT_TEKST_DATA FROM TBL_NOTA WHERE (NT_KLASA = '" & Klasa & "' AND NT_INDEKSI = '" & indeksi & "' AND NT_VITI_SHKOLLOR = '" & nrVitiShkollor & "') AND (NT_VJETORE = 1)"
                    ElseIf semestrale11 = "1" Then
                        q_res = "SELECT NT_NR_AMZA, NT_EMRI_LENDA, NT_VLERESIMI, NT_TEKST_DATA FROM TBL_NOTA WHERE (NT_KLASA = '" & Klasa & "' AND NT_INDEKSI = '" & indeksi & "' AND NT_VITI_SHKOLLOR = '" & nrVitiShkollor & "') AND (NT_SEMESTRI1 = 1)"
                    ElseIf semestrale21 = "1" Then
                        q_res = "SELECT NT_NR_AMZA, NT_EMRI_LENDA, NT_VLERESIMI, NT_TEKST_DATA FROM TBL_NOTA WHERE (NT_KLASA = '" & Klasa & "' AND NT_INDEKSI = '" & indeksi & "' AND NT_VITI_SHKOLLOR = '" & nrVitiShkollor & "') AND (NT_SEMESTRI2 = 1)"
                    ElseIf riprovim1 = "1" Then
                        q_res = "SELECT NT_NR_AMZA, NT_EMRI_LENDA, NT_VLERESIMI, NT_TEKST_DATA FROM TBL_NOTA WHERE (NT_KLASA = '" & Klasa & "' AND NT_INDEKSI = '" & indeksi & "' AND NT_VITI_SHKOLLOR = '" & nrVitiShkollor & "') AND (NT_RIPROVIM = 1)"
                    End If
                    tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
                    numerAmze = active_form.Amza.Text(i, 1)
                    lenda = vektor(i, j).lende
                    emri = Trim(active_form.emrat.Text(i, 2))
                    nota = vektor(i, j).note
                    If Not tmp.EOF Then
                        nota = vektor(i, j).note
                        If nota <> "" Then
                            numerAmze = active_form.Amza.Text(i, 1)
                            lenda = vektor(i, j).lende
                            emri = Trim(active_form.emrat.Text(i, 2))
                            ugjet = False
                            tmp.MoveFirst
                            Do While Not ugjet And Not tmp.EOF
                                If tmp.Fields("NT_NR_AMZA") = numerAmze And tmp.Fields("NT_EMRI_LENDA") = vektor(i, j).lende Then
                                    ugjet = True
                                End If
                                tmp.MoveNext
                            Loop
                            If Not ugjet Then
                            q_res1 = "INSERT INTO TBL_NOTA (NT_NR_AMZA, NT_CIKLI,NT_EMRI_LENDA,NT_VLERESIMI, NT_DATA,NT_TEKST_DATA,NT_KLASA, NT_INDEKSI, NT_VITI_SHKOLLOR, NT_SEMESTRI1, NT_SEMESTRI2, NT_VJETORE, NT_RIPROVIM) VALUES ( '" & numerAmze & "', " & strCikli & ", '" & lenda & "', '" & nota & "', " & " '" & data_konvertuar & "', '" & data_tekst & "', '" & Klasa & "', '" & indeksi & "', '" & nrVitiShkollor & "', " & semestrale11 & ", " & semestrale21 & ", " & vjetore1 & ", " & riprovim1 & ")"
                                conn.Execute q_res1
                                nr_hedhje = nr_hedhje + 1
                            Else
                                active_form.Cgrid1.CellSetFocus i, j
                                MsgBox "Nxenesit '" & emri & "' i eshte hedhur nje note perfundimtare ne lenden '" & lenda & Chr(10) & "Ju nuk mund te hidhni dy nota perfundimtare per te njejtin nxenes ne te njejten lende.", vbCritical, "Hedhja e notave."
                                active_form.Cgrid1.Text(i, j) = ""
                            End If
                        End If
                        tmp.Close
                    Else
                        If nota <> "" Then
                            q_res1 = "INSERT INTO TBL_NOTA (NT_NR_AMZA, NT_CIKLI,NT_EMRI_LENDA,NT_VLERESIMI, NT_DATA,NT_TEKST_DATA,NT_KLASA, NT_INDEKSI, NT_VITI_SHKOLLOR, NT_SEMESTRI1, NT_SEMESTRI2, NT_VJETORE, NT_RIPROVIM) VALUES ( '" & numerAmze & "', " & strCikli & ", '" & lenda & "', '" & nota & "', " & " '" & data_konvertuar & "', '" & data_tekst & "', '" & Klasa & "', '" & indeksi & "', '" & nrVitiShkollor & "', " & semestrale11 & ", " & semestrale21 & ", " & vjetore1 & ", " & riprovim1 & ")"
                            conn.Execute q_res1
                            nr_hedhje = nr_hedhje + 1
                        End If
                        tmp.Close
                    End If
                ' Nese nota e hedhur eshte momentale e semestrit te pare ose te dyte
                ElseIf mungeseMe = "1" Or mungesePa = "1" Then
                    If momentales1 = "1" Or momentales2 = "1" Then
                        q_res = "SELECT NT_NR_AMZA, NT_EMRI_LENDA, NT_VLERESIMI, NT_TEKST_DATA FROM TBL_NOTA WHERE (NT_KLASA = '" & Klasa & "' AND NT_INDEKSI = '" & indeksi & "' AND NT_VITI_SHKOLLOR = '" & nrVitiShkollor & "') AND (NT_MOMENTALES1 = 1 OR NT_MOMENTALES2 = 1 OR NT_DETYREKONTROLLI = 1) AND (NT_MUNGESE_PA = 1 OR NT_MUNGESE_ME = 1) AND NT_TEKST_DATA = '" & data_tekst & "'"
                        tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
                        If Not tmp.EOF Then
                            nota = vektor(i, j).note
                            If nota <> "M" And nota <> "m" Then
                                MsgBox "Mungesa e hedhur nga ju nuk eshte ne formatin e sakte", vbExclamation, "Hedhja e notave"
                                active_form.Cgrid1.CellSetFocus i, j
                            Else
                                numerAmze = active_form.Amza.Text(i, 1)
                                lenda = vektor(i, j).lende
                                emri = Trim(active_form.emrat.Text(i, 2))
                                ugjet = False
                                Do While Not tmp.EOF And Not ugjet
                                    If tmp.Fields("NT_NR_AMZA") = numerAmze And tmp.Fields("NT_EMRI_LENDA") = lenda Then
                                        ugjet = True
                                    End If
                                    tmp.MoveNext
                                Loop
                                If Not ugjet Then
                                    'q_res = "INSERT INTO TBL_NOTA (NT_NR_AMZA, NT_CIKLI,NT_EMRI_LENDA,NT_VLERESIMI,NT_MOMENTALES1, NT_MOMENTALES2, NT_MUNGESE_PA,NT_MUNGESE_ME, NT_DATA,NT_TEKST_DATA,NT_KLASA, NT_INDEKSI, NT_VITI_SHKOLLOR) VALUES ( '"
                                    q_res1 = "INSERT INTO TBL_NOTA (NT_NR_AMZA, NT_CIKLI,NT_EMRI_LENDA,NT_VLERESIMI,NT_MOMENTALES1, NT_MOMENTALES2, NT_MUNGESE_PA,NT_MUNGESE_ME, NT_DATA,NT_TEKST_DATA,NT_KLASA, NT_INDEKSI, NT_VITI_SHKOLLOR) VALUES ( '" & numerAmze & "', " & strCikli & ", '" & lenda & "', '" & nota & "', " & momentales1 & "," & momentales2 & "," & mungesePa & ", " & mungeseMe & ", '" & data_konvertuar & "', '" & data_tekst & "', '" & Klasa & "', '" & indeksi & "', '" & nrVitiShkollor & "')"
                                    'q_res1 = numerAmze & "', " & strcikli & ", '" & lenda & "', '" & nota & "', " & momentales1 & "," & momentales2 & "," & mungeseMe & ", " & mungesePa & ", '" & data_konvertuar & "', '" & data_tekst & "', '" & Klasa & "', '" & indeksi & "', '" & nrVitiShkollor & "')"
                                    conn.Execute q_res1
                                    nr_hedhje = nr_hedhje + 1
                                Else
                                    active_form.Cgrid1.CellSetFocus i, j
                                    MsgBox "Nxenesit '" & emri & "' i eshte hedhur nje mungese ose note ne lenden '" & lenda & "' ne daten e percaktuar nga ju." & Chr(10) & "Ju nuk mund t'i hidhni dy vleresime ne te njejten date.", vbCritical, "Hedhja e mungesave."
                                    active_form.Cgrid1.Text(i, j) = ""
                                End If
                            End If
                            tmp.Close
                        Else
                            nota = vektor(i, j).note
                            If nota <> "" Then
                                If nota <> "m" And nota <> "M" Then
                                    MsgBox "Mungesa e hedhur nga ju nuk eshte ne format te sakte", vbExclamation, "Hedhja e notave"
                                Else
                                    nota = "m"
                                    numerAmze = active_form.Amza.Text(i, 1)
                                    lenda = active_form.Cgrid2.Text(1, j)
                                    q_res1 = "INSERT INTO TBL_NOTA (NT_NR_AMZA, NT_CIKLI,NT_EMRI_LENDA,NT_VLERESIMI,NT_MOMENTALES1, NT_MOMENTALES2, NT_MUNGESE_PA,NT_MUNGESE_ME, NT_DATA,NT_TEKST_DATA,NT_KLASA, NT_INDEKSI, NT_VITI_SHKOLLOR) VALUES ( '" & numerAmze & "', " & strCikli & ", '" & lenda & "', '" & nota & "', " & momentales1 & "," & momentales2 & "," & mungesePa & ", " & mungeseMe & ", '" & data_konvertuar & "', '" & data_tekst & "', '" & Klasa & "', '" & indeksi & "', '" & nrVitiShkollor & "')"
                                    conn.Execute q_res1
                                    nr_hedhje = nr_hedhje + 1
                                End If
                            End If
                        tmp.Close
                        End If
                    End If
                ElseIf momentales1 = "1" Or momentales2 = "1" Then
                    q_res = "SELECT NT_NR_AMZA,NT_EMRI_LENDA, NT_VLERESIMI,NT_TEKST_DATA FROM TBL_NOTA WHERE (NT_KLASA = '" & Klasa & "' AND NT_INDEKSI = '" & indeksi & "' AND NT_VITI_SHKOLLOR = '" & nrVitiShkollor & "' ) AND ( NT_MOMENTALES1 = 1 OR NT_MOMENTALES2 = 1 OR NT_DETYREKONTROLLI = 1 OR NT_MUNGESE_PA = 1 OR NT_MUNGESE_ME = 1) AND NT_TEKST_DATA = '" & data_tekst & "'"
                    tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
                    If Not tmp.EOF Then
                        nota = active_form.Cgrid1.Text(i, j)
                        nota = vektor(i, j).note
                      
                        If nota <> "" Then
                            If nota = "m" Or nota = "M" Then
                                MsgBox "Nota e hedhur nga ju nuk eshte ne formatin e sakte", vbExclamation, "Hedhja e notave"
                                active_form.Cgrid1.CellSetFocus i, j
                            Else
                                numerAmze = active_form.Amza.Text(i, 1)
                                lenda = active_form.Cgrid2.Text(1, j)
                                lenda = vektor(i, j).lende
                                emri = Trim(active_form.emrat.Text(i, 2))
                                ugjet = False
                                tmp.MoveFirst
                                Do While Not tmp.EOF And Not ugjet
                                    If tmp.Fields("NT_NR_AMZA") = numerAmze And tmp.Fields("NT_EMRI_LENDA") = lenda Then
                                        ugjet = True
                                    End If
                                        tmp.MoveNext
                                Loop
                                If Not ugjet Then
                                    q_res1 = "INSERT INTO TBL_NOTA (NT_NR_AMZA, NT_CIKLI,NT_EMRI_LENDA,NT_VLERESIMI,NT_MOMENTALES1,NT_MOMENTALES2, NT_DETYREKONTROLLI, NT_DATA,NT_TEKST_DATA,NT_KLASA, NT_INDEKSI, NT_VITI_SHKOLLOR) VALUES ( '" & numerAmze & "', " & strCikli & ", '" & lenda & "', '" & nota & "', " & momentales1 & ", " & momentales2 & ", " & detyrekontrolli & ", '" & data_konvertuar & "', '" & data_tekst & "', '" & Klasa & "', '" & indeksi & "', '" & nrVitiShkollor & "')"
                                    conn.Execute q_res1
                                    nr_hedhje = nr_hedhje + 1
                                Else
                                    active_form.Cgrid1.CellSetFocus i, j
                                    MsgBox "Nxenesit '" & emri & "' i eshte hedhur nje note ne lenden '" & lenda & "' ne daten e percaktuar nga ju." & Chr(10) & "Ju nuk mund t'i hidhni dy nota ne te njejten date.", vbCritical, "Hedhja e notave."
                                    active_form.Cgrid1.Text(i, j) = ""
                                End If
                            End If
                        End If
                        tmp.Close
                    Else
                        nota = active_form.Cgrid1.Text(i, j)
                        nota = vektor(i, j).note
                        If nota <> "" Then
                            If nota = "m" Or nota = "M" Then
                                MsgBox "Nota e hedhur nga ju nuk eshte ne formatin e sakte", vbExclamation, "Hedhja e notave"
                                active_form.Cgrid1.CellSetFocus i, j
                            Else
                                numerAmze = active_form.Amza.Text(i, 1)
                                lenda = active_form.Cgrid2.Text(1, j)
                                data_konvertuar = konvertoDaten(CStr(vektor(i, j).date))
                                q_res1 = "INSERT INTO TBL_NOTA (NT_NR_AMZA, NT_CIKLI,NT_EMRI_LENDA,NT_VLERESIMI,NT_MOMENTALES1,NT_MOMENTALES2, NT_DETYREKONTROLLI, NT_DATA, NT_TEKST_DATA, NT_KLASA, NT_INDEKSI, NT_VITI_SHKOLLOR) VALUES ( '" & numerAmze & "', " & strCikli & ", '" & vektor(i, j).lende & "', '" & vektor(i, j).note & "', " & momentales1 & ", " & momentales2 & ", " & detyrekontrolli & ", '" & data_konvertuar & "', '" & data_tekst & "', '" & Klasa & "', '" & indeksi & "', '" & nrVitiShkollor & "')"
                                conn.Execute q_res1
                                nr_hedhje = nr_hedhje + 1
                            End If
                        End If
                        tmp.Close
                    End If
                ElseIf mungeseMe = "1" Or mungesePa = "1" Then
                    nota = vektor(i, j).note
                    If Not (nota <> "" Or nota <> "M" Or nota <> "m") Then
                        MsgBox "Mungesat qe ju keni hedhur nuk jane te gjitha ne formatin e duhur", vbExclamation, "Hedhja e notave"
                    Else
                        data_tekst = DataToTekst(vektor(i, j).date)
                        data_konvertuar = konvertoDaten(CStr(data_tekst))
                        numerAmze = active_form.Amza.Text(i, 1)
                        lenda = vektor(i, j).lende
                        emri = Trim(active_form.emrat.Text(i, 2))
                        q_res = "SELECT NT_NR_AMZA,NT_EMRI_LENDA, NT_VLERESIMI,NT_TEKST_DATA FROM TBL_NOTA WHERE (NT_KLASA = '" & Klasa & "' AND NT_INDEKSI = '" & indeksi & "' AND NT_VITI_SHKOLLOR = '" & nrVitiShkollor & "' ) "
                        q_res1 = "AND ( NT_MUNGESE_PA = 1 OR NT_MUNGESE_ME = 1 OR NT_MOMENTALES1 = 1 OR NT_MOMENTALES2 = 1 OR NT_DETYREKONTROLLI = 1 ) AND NT_TEKST_DATA = '" & data_tekst & "'"
                        tmp.Open q_res + q_res1, conn, adOpenDynamic, adLockOptimistic
                        ugjet = False
                        If Not tmp.EOF Then
                            If nota <> "" Then
                                nota = "m"
                                tmp.MoveFirst
                                Do While Not tmp.EOF And Not ugjet
                                    If tmp.Fields("NT_NR_AMZA") = numerAmze And tmp.Fields("NT_EMRI_LENDA") = lenda Then
                                        ugjet = True
                                    End If
                                    tmp.MoveNext
                                Loop
                                If Not ugjet Then
                                    q_res1 = "INSERT INTO TBL_NOTA (NT_NR_AMZA, NT_CIKLI,NT_EMRI_LENDA,NT_VLERESIMI,NT_SEMESTRI1,NT_SEMESTRI2, NT_VJETORE,NT_RIPROVIM, NT_DATA,NT_TEKST_DATA,NT_KLASA, NT_INDEKSI, NT_VITI_SHKOLLOR) VALUES ( '" & numerAmze & "', " & strCikli & ", '" & lenda & "', '" & nota & "', " & semestrale11 & ", " & semestrale21 & ", " & vjetore1 & ", " & riprovim1 & ", '" & data_konvertuar & "', '" & data_tekst & "', '" & Klasa & "', '" & indeksi & "', '" & nrVitiShkollor & "')"
                                    conn.Execute q_res1
                                    nr_hedhje = nr_hedhje + 1
                                Else
                                    active_form.Cgrid1.CellSetFocus i, j
                                    MsgBox "Nxenesit '" & emri & "' i eshte hedhur nje note ne lenden '" & lenda & "' ." & Chr(10) & "Ju nuk mund t'i hidhni dy nota semestrale, vjetore ose per provimet e vjeshtes .", vbCritical, "Hedhja e notave."
                                    active_form.Cgrid1.Text(i, j) = ""
                                End If
                            End If
                            tmp.Close
                        Else
                            If nota <> "" Then
                                nota = "m"
                                q_res = "INSERT INTO TBL_NOTA (NT_NR_AMZA, NT_CIKLI,NT_EMRI_LENDA,NT_VLERESIMI,NT_SEMESTRI1,NT_SEMESTRI2, NT_VJETORE, NT_MUNGESE_ME, NT_MUNGESE_PA, NT_RIPROVIM, NT_DATA,NT_TEKST_DATA,NT_KLASA, NT_INDEKSI, NT_VITI_SHKOLLOR)"
                                q_res1 = " VALUES ( '" & numerAmze & "', " & strCikli & ", '" & lenda & "', '" & nota & "', " & semestrale11 & ", " & semestrale21 & ", " & vjetore1 & ", " & mungeseMe & ", " & mungesePa & ", " & riprovim1 & ", '" & data_konvertuar & "', '" & data_tekst & "', '" & Klasa & "', '" & indeksi & "', '" & nrVitiShkollor & "')"
                                conn.Execute q_res + q_res1
                                nr_hedhje = nr_hedhje + 1
                            End If
                            tmp.Close
                        End If
                    End If
                End If
            Next i
        Next j
         If nr_hedhje > 0 Then
            MsgBox "Notat u hodhen.", vbInformation, "Hedhja e notave."
         End If
         For j = 1 To nr_lenda_momentale
            For i = 1 To nr_nxenesit_momentale
               active_form.Cgrid1.Text(i, j) = ""

            Next i
         Next j

      Case MODIFIKO_PROVIME

         nrVitiShkollor = active_form.cboVitiShkollor.Text

         '*** ketu percaktohet cikli i nxenesit

         If active_form.optMesme.Value Then
            cikli = 1
            strCikli = "1"
            emriProvimi = "matures"
         Else
            cikli = 0
            strCikli = "0"
            emriProvimi = "lirimit"
         End If

         '***
         '*** ketu kryhet hedhja e te dhenave plotesuese te nxenesit
         If active_form.txtAmzaNo.Text <> "" Then
            numerAmze = active_form.txtAmzaNo.Text

            q_res = " SELECT AMZA_NR_AMZA, AMZA_EMRI, AMZA_MBIEMRI, AMZA_ATESIA, AMZA_DATELINDJA, AMZA_VREJTJE, AMZA_VENDLINDJA,KL_KLASA,KL_INDEKSI, KL_VITI_SHKOLLOR FROM TBL_AMZA, TBL_KLASA WHERE ( AMZA_NR_AMZA = KL_NR_AMZA AND AMZA_CIKLI = KL_CIKLI) AND (AMZA_NR_AMZA = '" & numerAmze & "' AND AMZA_CIKLI = " & strCikli & ") AND KL_VITI_SHKOLLOR = '" & nrVitiShkollor & "' ORDER BY AMZA_EMRI, AMZA_MBIEMRI"
            tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic

            'get_rec (q_res)
            If Not tmp.EOF Then
               active_form.txtEmri.Text = tmp.Fields("AMZA_EMRI")
               active_form.txtMbiemri.Text = tmp.Fields("AMZA_MBIEMRI")
               active_form.txtAtesia = tmp.Fields("AMZA_ATESIA")
               active_form.txtKlasa.Text = tmp.Fields("KL_KLASA")
               active_form.txtIndeksi.Text = tmp.Fields("KL_INDEKSI")
               tmp.Close
Else:
               tmp.Close

               active_form.txtEmri.Text = ""
               active_form.txtMbiemri.Text = ""
               active_form.txtAtesia.Text = ""
               active_form.txtKlasa.Text = ""
               active_form.txtIndeksi.Text = ""
               ' Atesia, Vendlindja, Datelindja, Vrejtje,  sherbejne vetem per
               ' printimin e defteses

               active_form.Label3.Visible = False
               active_form.txtData.Visible = False
               active_form.Cgrid1.Visible = False
               active_form.Lendet.Visible = False
               active_form.Label2.Visible = False
               active_form.lblNotatDheMungesat.Visible = False
               MsgBox "   Nxenesi me numrin e amzes se dhene nuk eshte i regjistruar ne vitin shkollor qe eshte dhene" & vbCrLf & "   dhe ne ciklin qe eshte zgjedhur", vbInformation, "  Konsultime Nota"
               active_form.txtAmzaNo.Text = ""
               Exit Sub
            End If
         Else
            emri = active_form.txtEmri.Text
            mbiemri = active_form.txtMbiemri.Text
            atesia = active_form.txtAtesia.Text
            q_res = " SELECT AMZA_NR_AMZA, AMZA_CIKLI,AMZA_EMRI, AMZA_MBIEMRI, AMZA_ATESIA, AMZA_DATELINDJA, AMZA_VREJTJE, AMZA_VENDLINDJA,KL_KLASA,KL_INDEKSI, KL_VITI_SHKOLLOR FROM TBL_AMZA, TBL_KLASA  WHERE (AMZA_NR_AMZA = KL_NR_AMZA AND AMZA_CIKLI = KL_CIKLI ) AND ( AMZA_EMRI LIKE '" & emri & "%' AND AMZA_MBIEMRI LIKE '" & mbiemri & "%' AND AMZA_ATESIA LIKE '" & atesia & "%' AND AMZA_CIKLI = " & strCikli & ")  AND KL_VITI_SHKOLLOR = '" & nrVitiShkollor & "' ORDER BY AMZA_EMRI, AMZA_MBIEMRI"
            tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
            If Not tmp.EOF Then
                If tmp.RecordCount = 1 Then
                    active_form.txtAmzaNo.Text = tmp.Fields("AMZA_NR_AMZA")
                    active_form.txtKlasa.Text = tmp.Fields("KL_KLASA")
                    active_form.txtIndeksi.Text = tmp.Fields("KL_INDEKSI")
                    active_form.txtEmri.Text = tmp.Fields("AMZA_EMRI")
                    active_form.txtMbiemri.Text = tmp.Fields("AMZA_MBIEMRI")
                    active_form.txtAtesia = tmp.Fields("AMZA_ATESIA")
                    tmp.Close
                Else
                    'kerko
                    Dim nrRreshta As Integer
                    nrRreshta = tmp.RecordCount
                    Call inicializoGridaZgjidhNxenes(nrRreshta)
                    'mbush griden me te dhena
                    Dim Index As Integer
                    Index = 0
                    Do While Not tmp.EOF
                        frmZgjidhNxenes.gridaNxenesit.row = Index
                        frmZgjidhNxenes.gridaNxenesit.col = 0
                        frmZgjidhNxenes.gridaNxenesit.Text = tmp.Fields("AMZA_NR_AMZA").Value
                        frmZgjidhNxenes.gridaNxenesit.col = 1
                        frmZgjidhNxenes.gridaNxenesit.Text = tmp.Fields("AMZA_EMRI").Value + " " + tmp.Fields("AMZA_ATESIA").Value + " " + tmp.Fields("AMZA_MBIEMRI").Value
                        frmZgjidhNxenes.gridaNxenesit.col = 2
                        frmZgjidhNxenes.gridaNxenesit.Text = tmp.Fields("AMZA_CIKLI").Value
                        tmp.MoveNext
                        Index = Index + 1
                    Loop
                    tmp.Close
                    ShkollaManager.frmZgjidhNxenes.show vbModal
                    numerAmze = ShkollaManager.frmZgjidhNxenes.Amza
                    strCikli = ShkollaManager.frmZgjidhNxenes.cikli
                    If (numerAmze = "" Or strCikli = "") Then
                    'pastro
                    Exit Sub
                    End If
                    q_res = " SELECT AMZA_NR_AMZA, AMZA_EMRI, AMZA_MBIEMRI, AMZA_ATESIA, AMZA_DATELINDJA, AMZA_VREJTJE, AMZA_VENDLINDJA,KL_KLASA,KL_INDEKSI, KL_VITI_SHKOLLOR FROM TBL_AMZA, TBL_KLASA WHERE ( AMZA_NR_AMZA = KL_NR_AMZA AND AMZA_CIKLI = KL_CIKLI) AND (AMZA_NR_AMZA = '" & numerAmze & "' AND AMZA_CIKLI = " & strCikli & ") AND KL_VITI_SHKOLLOR = '" & nrVitiShkollor & "' ORDER BY AMZA_EMRI, AMZA_MBIEMRI"
                    tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
                     If Not tmp.EOF Then
                        active_form.txtEmri.Text = tmp.Fields("AMZA_EMRI")
                        active_form.txtMbiemri.Text = tmp.Fields("AMZA_MBIEMRI")
                        active_form.txtAtesia = tmp.Fields("AMZA_ATESIA")
                        active_form.txtKlasa.Text = tmp.Fields("KL_KLASA")
                        active_form.txtIndeksi.Text = tmp.Fields("KL_INDEKSI")
                        active_form.txtAmzaNo.Text = tmp.Fields("AMZA_NR_AMZA")
                        tmp.Close
                    End If
                End If
               
Else:
               tmp.Close
               active_form.txtEmri.Text = ""
               active_form.txtMbiemri.Text = ""
               active_form.txtAtesia.Text = ""
               active_form.Label3.Visible = False
               active_form.txtData.Visible = False
               active_form.Cgrid1.Visible = False
               active_form.Lendet.Visible = False
               active_form.Label2.Visible = False
               active_form.lblNotatDheMungesat.Visible = False
               MsgBox "   Nxenesi me kete emer dhe mbiemer  nuk eshte i regjistruar ne vitin shkollor qe eshte dhene" & vbCrLf & "   dhe ne ciklin qe eshte zgjedhur", vbInformation, "  Konsultime Nota"
               Exit Sub
            End If
         End If
         numerAmze = active_form.txtAmzaNo.Text

         '*** Ketu inicialozohet grida
         q_res = "SELECT NT_EMRI_LENDA, NT_VLERESIMI FROM TBL_NOTA WHERE NT_CIKLI = " & strCikli & " AND NT_NR_AMZA = '" & numerAmze & "' AND NT_PROVIM = 1 "
         tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
         If tmp.EOF Then
            MsgBox "Nxenesi nuk ka marre akoma asnje note ne provimet e " & emriProvimi & ".", vbInformation, "Modifikimi i provimeve."
            active_form.txtAmzaNo.Text = ""
            active_form.txtEmri.Text = ""
            active_form.txtMbiemri.Text = ""
            active_form.txtAtesia.Text = ""
            active_form.txtKlasa.Text = ""
            active_form.txtIndeksi.Text = ""
            active_form.cboVitiShkollor.Text = gjej_vitin(DateTime.Month(DateTime.Now), DateTime.Year(DateTime.Now))
            Exit Sub
         End If
         i = 0
         tmp.MoveFirst
         Do While Not tmp.EOF
            i = i + 1
            tmp.MoveNext
         Loop

         nr_provimesh = i
         shfaq_grilaProvimet nr_provimesh, 2

         tmp.MoveFirst
         i = 0
         Do While Not tmp.EOF
            i = i + 1
            active_form.gridaProvime.Text(1, i) = tmp.Fields("NT_EMRI_LENDA")
            active_form.gridaProvime.Text(2, i) = tmp.Fields("NT_VLERESIMI")
            tmp.MoveNext
         Loop
         tmp.Close
         active_form.gridaProvime.Visible = True
         If strCikli = "1" Then
            active_form.Label5.Caption = "Provimet e matures :"
         Else
            active_form.Label5.Caption = "Provimet e lirimit :"
         End If

      Case NDRYSHO_PROVIME

         numerAmze = active_form.txtAmzaNo.Text
         If active_form.optMesme.Value Then
            strCikli = "1"
         Else
            strCikli = "0"
         End If
         lenda = active_form.gridaProvime.Text(1, shtyllaModifiko)
         Dim nota1, nota2     As String
         If modifikoNota = 1 Then
            q_res = "DELETE FROM TBL_NOTA WHERE NT_NR_AMZA = '" & numerAmze & "' AND NT_CIKLI = " & strCikli & " AND NT_EMRI_LENDA = '" & lenda & "'"
            conn.Execute q_res
            active_form.gridaProvime.Text(rreshtiModifiko, shtyllaModifiko) = ""
            MsgBox "Nota e provimit '" & lenda & "' u fshi.", vbInformation, "Modifikimi i provimeve."
         Else
            nota1 = active_form.gridaProvime.Text(rreshtiModifiko, shtyllaModifiko)
            nota2 = InputBox("Jepni noten e provimit :", "Modifikimi i provimeve.")
            If nota1 = nota2 Or nota2 = "" Then
               Exit Sub
            End If
            If Not (nota2 = "4" Or nota2 = "5" Or nota2 = "6" Or nota2 = "7" Or nota2 = "8" Or nota2 = "9" Or nota2 = "10") Then
               MsgBox "Nota e dhene nga ju nuk eshte ne formatin e duhur.", vbInformation, "Modifikimi i provimeve."
               Exit Sub
            End If

            q_res = "UPDATE TBL_NOTA SET NT_VLERESIMI = '" & nota2 & "' WHERE NT_NR_AMZA = '" & numerAmze & "' AND NT_CIKLI = " & strCikli & " AND NT_EMRI_LENDA = '" & lenda & "'"
            conn.Execute q_res
            active_form.gridaProvime.Text(rreshtiModifiko, shtyllaModifiko) = nota2
            MsgBox "Nota e provimit '" & lenda & "' u modifikua.", vbInformation, "Modifikimi i provimeve."
         End If
      
      Case HIDH_NE_EKSEL
         Dim rekset           As New ADODB.Recordset
         Dim FileExcel        As Workbook
         Dim FoglioExcel      As Worksheet
         'Dim CellaFoglioExcel As Range
         Dim skedari          As String
         On Error GoTo GabimExcel
KontrollDiskete:
         If Not IsDriveReady("A:") Then
            MsgBox "Ju lutemi fusni nje diskete dhe pastaj shtypni OK", vbExclamation + vbOKCancel, "Fusni Disketen"
            If vbCancel Then

               Exit Sub
            Else

               GoTo KontrollDiskete
            End If
         End If

         KopjoSkedar App.Path + "\amza.xls", "A:\"
         'skedari = Dir("A:\amza.xls")
         'If skedari = "" Then
         'cop
         'MsgBox "Ne disketen tuaj nuk ekziston nje skedar ekseli i domosdoshem per hedhjen e te dhenave.", vbInformation, "Hedhja ne eksel."
         'Exit Sub
         'End If
         'skedari = Dir("A:\amza.xls")
         'If skedari = "" Then
         'MsgBox "Ju duhet te riemeroni skedarin Excel .Ju duhet t'i jepni emrin 'amza.cls'.", vbExclamation, "Hedhja e te dhenave ne eksel."
         'Exit Sub
         'End If
         Set FileExcel = Excel.Workbooks.Open("A:" & "\amza.xls")
         'Set FoglioExcel = FileExcel.Worksheets("Shkolla")
         'FileExcel.Worksheets.Add "Shkolla"

         Set FoglioExcel = FileExcel.Worksheets("Shkolla")
         On Error GoTo GabimExcel
         FoglioExcel.Cells(1, 2) = "AMZA E SHKOLLES SE MESME E BASHKUAR : " & "   " & emerShkolla
         FoglioExcel.Cells(3, 1) = "Nr"
         FoglioExcel.Cells(3, 2) = "Rrethi"
         FoglioExcel.Cells(3, 3) = "Qytet-Fshat"
         FoglioExcel.Cells(3, 4) = "Shkolla"
         FoglioExcel.Cells(3, 5) = "Emri"
         FoglioExcel.Cells(3, 6) = "Atesia"
         FoglioExcel.Cells(3, 7) = "Mbiemri"
         FoglioExcel.Cells(3, 8) = "Amza"
         FoglioExcel.Cells(3, 9) = "Datelindja"
         FoglioExcel.Cells(3, 10) = "Vendlindja"
         FoglioExcel.Cells(3, 11) = "Letersi1"
         FoglioExcel.Cells(3, 12) = "Anglisht1"
         FoglioExcel.Cells(3, 13) = "Histori1"
         FoglioExcel.Cells(3, 14) = "Gjeografi1"
         FoglioExcel.Cells(3, 15) = "NjShoqerore1"
         FoglioExcel.Cells(3, 16) = "Matemat1"
         FoglioExcel.Cells(3, 17) = "Fizike1"
         FoglioExcel.Cells(3, 18) = "Kimi1"
         FoglioExcel.Cells(3, 19) = "Biologji1"
         FoglioExcel.Cells(3, 20) = "Teknolo1"
         FoglioExcel.Cells(3, 21) = "EdFizik1"
         FoglioExcel.Cells(3, 22) = "Letersi2"
         FoglioExcel.Cells(3, 23) = "Anglisht2"
         FoglioExcel.Cells(3, 24) = "Histori2"
         FoglioExcel.Cells(3, 25) = "Gjeografi2"
         FoglioExcel.Cells(3, 26) = "NjShoqerore2"
         FoglioExcel.Cells(3, 27) = "Matemat2"
         FoglioExcel.Cells(3, 28) = "Fizike2"
         FoglioExcel.Cells(3, 29) = "Kimi2"
         FoglioExcel.Cells(3, 30) = "Biologji2"
         FoglioExcel.Cells(3, 31) = "Teknolo2"
         FoglioExcel.Cells(3, 32) = "EdFizik2"
         FoglioExcel.Cells(3, 33) = "Letersi3"
         FoglioExcel.Cells(3, 34) = "Anglisht3"
         FoglioExcel.Cells(3, 35) = "Histori3"
         FoglioExcel.Cells(3, 36) = "Gjeografi3"
         FoglioExcel.Cells(3, 37) = "NjShoqerore3"
         FoglioExcel.Cells(3, 38) = "Matemat3"
         FoglioExcel.Cells(3, 39) = "Fizike3"
         FoglioExcel.Cells(3, 40) = "Kimi3"
         FoglioExcel.Cells(3, 41) = "Biologji3"
         FoglioExcel.Cells(3, 42) = "Informat3"
         FoglioExcel.Cells(3, 43) = "EdFizik3"
         FoglioExcel.Cells(3, 44) = "Letersi4"
         FoglioExcel.Cells(3, 45) = "Anglisht4"
         FoglioExcel.Cells(3, 46) = "Histori4"
         FoglioExcel.Cells(3, 47) = "Gjeograf4"
         FoglioExcel.Cells(3, 48) = "NjShoqerore4"
         FoglioExcel.Cells(3, 49) = "Matemat4"
         FoglioExcel.Cells(3, 50) = "Fizike4"
         FoglioExcel.Cells(3, 51) = "Kimi4"
         FoglioExcel.Cells(3, 52) = "Biologji4"
         FoglioExcel.Cells(3, 53) = "Informat4"
         FoglioExcel.Cells(3, 54) = "EdFizik4"
         FoglioExcel.Cells(3, 55) = "Astron4"
         FoglioExcel.Cells(3, 56) = "LetersiPM"
         FoglioExcel.Cells(3, 57) = "MatPM"
         FoglioExcel.Cells(3, 58) = "FizikPM"
         Dim M, v As Integer
         M = Val(DateTime.Month(DateTime.Now))
         v = Val(DateTime.Year(DateTime.date))
         nrVitiShkollor = gjej_vitin(M, v)

         'nrVitiShkollor = "2004-2005"
         '  ;Set CellaFoglioExcel = FoglioExcel.Cells(2, 1)
         'Label1 = FoglioExcel.Cells(2, 1)
         q_res1 = "SELECT * FROM TBL_AMZA WHERE AMZA_CIKLI = 1 AND AMZA_VITI_SHKOLLOR = '" & nrVitiShkollor & "' ORDER BY AMZA_KLASA, AMZA_NR_AMZA "
         rekset.Open q_res1, conn, adOpenDynamic, adLockOptimistic
         If rekset.EOF Then
            MsgBox "Per momentin nuk ka asnje nxenes ne amze.", vbInformation, "Hedhja ne eksel."
            FileExcel.Close (True)
            Set FileExcel = Nothing
            Set FoglioExcel = Nothing
            Exit Sub
         End If

         i = 4

         Do While Not rekset.EOF
            i = i + 1
            FoglioExcel.Cells(i, 1) = str(i - 4)
            FoglioExcel.Cells(i, 4) = emerShkolla
            FoglioExcel.Cells(i, 5) = rekset.Fields("AMZA_EMRI")
            FoglioExcel.Cells(i, 6) = rekset.Fields("AMZA_ATESIA")
            FoglioExcel.Cells(i, 7) = rekset.Fields("AMZA_MBIEMRI")
            FoglioExcel.Cells(i, 8) = rekset.Fields("AMZA_NR_AMZA")
            FoglioExcel.Cells(i, 9) = rekset.Fields("AMZA_DATELINDJA")
            FoglioExcel.Cells(i, 10) = rekset.Fields("AMZA_VENDLINDJA")
            rekset.MoveNext
         Loop
         rekset.MoveFirst

         q_res = "SELECT AMZA_NR_AMZA,AMZA_EMRI, AMZA_MBIEMRI, AMZA_ATESIA, AMZA_VENDLINDJA, AMZA_DATELINDJA, NT_VLERESIMI, NT_EMRI_LENDA, NT_VITI_SHKOLLOR, NT_KLASA, NT_VITI_SHKOLLOR, NT_VJETORE, NT_PROVIM FROM TBL_AMZA, TBL_NOTA WHERE (AMZA_NR_AMZA = NT_NR_AMZA AND AMZA_CIKLI = NT_CIKLI AND AMZA_CIKLI = 1 AND NT_CIKLI = 1) AND (NT_VJETORE = 1 OR NT_PROVIM = 1) ORDER BY AMZA_KLASA, AMZA_NR_AMZA "
         tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
         If tmp.EOF Then
            FileExcel.Close (True)
            Set FileExcel = Nothing
            Set FoglioExcel = Nothing
            Exit Sub
         End If

         rekset.MoveFirst
         i = 4
         Do While Not rekset.EOF
            i = i + 1
            numerAmze = rekset.Fields("AMZA_NR_AMZA")
            tmp.MoveFirst
            Do While Not tmp.EOF
               If tmp.Fields("AMZA_NR_AMZA") = numerAmze Then
                  lenda = tmp.Fields("NT_EMRI_LENDA")
                  Klasa = tmp.Fields("NT_KLASA")
                  If tmp.Fields("NT_VJETORE") = True Then
                     If Klasa = "9" Then
                        Select Case lenda
                           Case "Letersi"
                              FoglioExcel.Cells(i, 11) = tmp.Fields("NT_VLERESIMI")
                           Case "Anglisht"
                              FoglioExcel.Cells(i, 12) = tmp.Fields("NT_VLERESIMI")
                           Case "Histori"
                              FoglioExcel.Cells(i, 13) = tmp.Fields("NT_VLERESIMI")
                           Case "GJeografi"
                              FoglioExcel.Cells(i, 14) = tmp.Fields("NT_VLERESIMI")
                           Case "Njohuri per shoqerine"
                              FoglioExcel.Cells(i, 15) = tmp.Fields("NT_VLERESIMI")
                           Case "Matematike"
                              FoglioExcel.Cells(i, 16) = tmp.Fields("NT_VLERESIMI")
                           Case "Fizike"
                              FoglioExcel.Cells(i, 17) = tmp.Fields("NT_VLERESIMI")
                           Case "Kimi"
                              FoglioExcel.Cells(i, 18) = tmp.Fields("NT_VLERESIMI")
                           Case "Biologji"
                              FoglioExcel.Cells(i, 19) = tmp.Fields("NT_VLERESIMI")
                           Case "Teknologji"
                              FoglioExcel.Cells(i, 20) = tmp.Fields("NT_VLERESIMI")
                           Case "Edukim fizik"
                              FoglioExcel.Cells(i, 21) = tmp.Fields("NT_VLERESIMI")
                           Case Else
                        End Select

                     ElseIf Klasa = "10" Then

                        Select Case lenda
                           Case "Letersi"
                              FoglioExcel.Cells(i, 22) = tmp.Fields("NT_VLERESIMI")
                           Case "Anglisht"
                              FoglioExcel.Cells(i, 23) = tmp.Fields("NT_VLERESIMI")
                           Case "Histori"
                              FoglioExcel.Cells(i, 24) = tmp.Fields("NT_VLERESIMI")
                           Case "GJeografi"
                              FoglioExcel.Cells(i, 25) = tmp.Fields("NT_VLERESIMI")
                           Case "Njohuri per shoqerine"
                              FoglioExcel.Cells(i, 26) = tmp.Fields("NT_VLERESIMI")
                           Case "Matematike"
                              FoglioExcel.Cells(i, 27) = tmp.Fields("NT_VLERESIMI")
                           Case "Fizike"
                              FoglioExcel.Cells(i, 28) = tmp.Fields("NT_VLERESIMI")
                           Case "Kimi"
                              FoglioExcel.Cells(i, 29) = tmp.Fields("NT_VLERESIMI")
                           Case "Biologji"
                              FoglioExcel.Cells(i, 30) = tmp.Fields("NT_VLERESIMI")
                           Case "Teknologji"
                              FoglioExcel.Cells(i, 31) = tmp.Fields("NT_VLERESIMI")
                           Case "Edukim fizik"
                              FoglioExcel.Cells(i, 32) = tmp.Fields("NT_VLERESIMI")
                           Case Else
                        End Select

                     ElseIf Klasa = "11" Then

                        Select Case lenda
                           Case "Letersi"
                              FoglioExcel.Cells(i, 33) = tmp.Fields("NT_VLERESIMI")
                           Case "Anglisht"
                              FoglioExcel.Cells(i, 34) = tmp.Fields("NT_VLERESIMI")
                           Case "Histori"
                              FoglioExcel.Cells(i, 35) = tmp.Fields("NT_VLERESIMI")
                           Case "GJeografi"
                              FoglioExcel.Cells(i, 36) = tmp.Fields("NT_VLERESIMI")
                           Case "Njohuri per shoqerine"
                              FoglioExcel.Cells(i, 37) = tmp.Fields("NT_VLERESIMI")
                           Case "Matematike"
                              FoglioExcel.Cells(i, 38) = tmp.Fields("NT_VLERESIMI")
                           Case "Fizike"
                              FoglioExcel.Cells(i, 39) = tmp.Fields("NT_VLERESIMI")
                           Case "Kimi"
                              FoglioExcel.Cells(i, 40) = tmp.Fields("NT_VLERESIMI")
                           Case "Biologji"
                              FoglioExcel.Cells(i, 41) = tmp.Fields("NT_VLERESIMI")
                           Case "Informatike"
                              FoglioExcel.Cells(i, 42) = tmp.Fields("NT_VLERESIMI")
                           Case "Edukim fizik"
                              FoglioExcel.Cells(i, 43) = tmp.Fields("NT_VLERESIMI")
                           Case Else
                        End Select

                     ElseIf Klasa = "12" Then

                        Select Case lenda
                           Case "Letersi"
                              FoglioExcel.Cells(i, 44) = tmp.Fields("NT_VLERESIMI")
                           Case "Anglisht"
                              FoglioExcel.Cells(i, 45) = tmp.Fields("NT_VLERESIMI")
                           Case "Histori"
                              FoglioExcel.Cells(i, 46) = tmp.Fields("NT_VLERESIMI")
                           Case "GJeografi"
                              FoglioExcel.Cells(i, 47) = tmp.Fields("NT_VLERESIMI")
                           Case "Njohuri per shoqerine"
                              FoglioExcel.Cells(i, 48) = tmp.Fields("NT_VLERESIMI")
                           Case "Matematike"
                              FoglioExcel.Cells(i, 49) = tmp.Fields("NT_VLERESIMI")
                           Case "Fizike"
                              FoglioExcel.Cells(i, 50) = tmp.Fields("NT_VLERESIMI")
                           Case "Kimi"
                              FoglioExcel.Cells(i, 51) = tmp.Fields("NT_VLERESIMI")
                           Case "Biologji"
                              FoglioExcel.Cells(i, 52) = tmp.Fields("NT_VLERESIMI")
                           Case "Informatike"
                              FoglioExcel.Cells(i, 53) = tmp.Fields("NT_VLERESIMI")
                           Case "Edukim fizik"
                              FoglioExcel.Cells(i, 54) = tmp.Fields("NT_VLERESIMI")
                           Case "Astronomi"
                              FoglioExcel.Cells(i, 55) = tmp.Fields("NT_VLERESIMI")
                           Case Else
                        End Select
                     End If
                  End If
                  If tmp.Fields("NT_PROVIM") = True Then
                     Select Case lenda
                        Case "Letersi"
                           FoglioExcel.Cells(i, 56) = tmp.Fields("NT_VLERESIMI")
                        Case "Matematike"
                           FoglioExcel.Cells(i, 57) = tmp.Fields("NT_VLERESIMI")
                        Case "Fizike"
                           FoglioExcel.Cells(i, 58) = tmp.Fields("NT_VLERESIMI")
                        Case Else
                     End Select
                  End If

               End If
               tmp.MoveNext
            Loop
            rekset.MoveNext
         Loop
GabimExcel:
        If (Not (rekset.ActiveConnection Is Nothing)) Then
         rekset.Close
         End If
        If (Not (tmp.ActiveConnection Is Nothing)) Then
         tmp.Close
        End If
          FileExcel.Close (True)
         Set FileExcel = Nothing
         Set FoglioExcel = Nothing
         ' Set CellaFoglioExcel = Nothing
         MsgBox "Te dhenat e amzes bashke me notat vjetore u hodhen ne skedari"
        
      

      Case KONTROLLO_REGJISTRIMIN
         Dim viti_i_shkolles  As String
         If active_form.optMesme.Value Then
            strCikli = "1"
         Else
            strCikli = "0"
         End If
         nrVitiShkollor = active_form.cboVitiShkollor2.Text
         indeksi = active_form.cboIndeksi2.Text
         Klasa = active_form.cboKlasa2.Text
         q_res = "SELECT * FROM TBL_KLASA WHERE KL_CIKLI = " & strCikli
         tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
         If tmp.EOF Then
            tmp.Close
            Exit Sub
         End If
         i = 0
         Do While (i <= active_form.lista2.ListCount - 1)
            If active_form.lista2.Selected(i) Then
               numerAmze = active_form.lista2Amza.List(i)
               emri = active_form.lista2.List(i)
               j = 0
               tmp.MoveFirst
               Do While Not tmp.EOF
                  If tmp.Fields("KL_NR_AMZA") = numerAmze Then
                     j = j + 1
                  End If
                  tmp.MoveNext
               Loop
               If j > 1 Then
                  a = MsgBox("Ju jeni duke fshire nxenesin e zgjedhur nga klasa e shfaqur ne ekran." & Chr(10) & "Fshirja e nxenesit nga klasa  do te shoqerohet me fshirjen  te gjitha vleresimeve te marra ne kte klase." & Chr(10) & "Jeni te sigurte se doni te fshini nxenesin ' " & emri & " ' ?", vbExclamation + vbOKCancel, "Konfigurimi i klasave.")
                  If a = vbOK Then
                     numerAmze = active_form.lista2Amza.List(i)
                     q_res1 = "DELETE FROM TBL_KLASA WHERE KL_NR_AMZA = '" & numerAmze & "' AND KL_CIKLI = " & strCikli & " AND KL_VITI_SHKOLLOR = '" & nrVitiShkollor & "'"
                     conn.Execute q_res1
                     q_res1 = "SELECT * FROM TBL_KLASA WHERE KL_NR_AMZA = '" & numerAmze & "' AND KL_CIKLI = " & strCikli & " ORDER BY KL_VITI_SHKOLLOR DESC"
                     rekset.Open q_res1, conn, adOpenDynamic, adLockOptimistic
                     Klasa = rekset.Fields("KL_KLASA")
                     indeksi = rekset.Fields("KL_INDEKSI")
                     viti_i_shkolles = rekset.Fields("KL_VITI_SHKOLLOR")
                     rekset.Close
                     q_res1 = "UPDATE TBL_AMZA SET AMZA_KLASA = '" & Klasa & "' , AMZA_INDEKSI = '" & indeksi & "', AMZA_VITI_SHKOLLOR = '" & viti_i_shkolles & "' WHERE AMZA_NR_AMZA = '" & numerAmze & "' AND AMZA_CIKLI = " & strCikli
                     conn.Execute q_res1
                     q_res1 = "DELETE FROM TBL_NOTA WHERE NT_NR_AMZA = '" & numerAmze & "' AND NT_CIKLI = " & strCikli & " AND NT_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND NT_KLASA = '" & Klasa & "' AND NT_INDEKSI = '" & indeksi & "'"
                     conn.Execute q_res1
                     emri = active_form.lista2.List(i)
                     active_form.lista2.RemoveItem (i)
                     active_form.lista2Amza.RemoveItem (i)
                     MsgBox "Nxenesi ' " & emri & " ' u fshi nga klasa.", vbExclamation, "Konfigurimi i klasave."
                     i = i - 1
                  End If

               Else
                  MsgBox "Nxenesi  " & active_form.lista2.List(i) & "  me numer amze  " & numerAmze & _
                     "  eshte i regjistruar vetem ne nje klase." & Chr(10) & "Heqja e tij nga klasa e dhene do te bente qe nxenesi mos te figuronte ne asnje klase." & Chr(10) & "Pra ju nuk mund ta hiqni ate nga kjo klase." & Chr(10) & "Per t'i ndryshuar klasen nxenesit shkoni te modifikimet e gjeneraliteteve" & Chr(10) & "dhe modifikoni klasen e nxenesit,pastaj fshijeni nga klasa ekzistuese.", vbExclamation, "Konfigurimi i klasave."
               End If
            End If
            i = i + 1
         Loop
         tmp.Close

      Case PRINTO_DEFTESA_SERI
         ' Vektor qe do te permbaje vektoret e te dhenave dhe notave si elemente
         Dim q_res3, q_res4   As String
         Dim TeDhenatVecanta(60) As VektorTeDhenash
         Dim objPrintim       As Object
         Dim NumriNxenesve    As Integer
         Dim numriAmzes       As String
         Dim LendaParaardhese As String
         '*** q_res1 nxjerr numrat e amzes, emrat dhe mbiemrat e nje klase te caktuar.
         q_res1 = "SELECT AMZA_EMRI, AMZA_LARGUAR, AMZA_CIKLI, AMZA_MBIEMRI, AMZA_DATELINDJA, AMZA_NR_AMZA, AMZA_ATESIA, AMZA_VENDLINDJA FROM TBL_AMZA, TBL_KLASA  WHERE (AMZA_NR_AMZA = KL_NR_AMZA AND AMZA_CIKLI = KL_CIKLI) AND (KL_KLASA = '" & active_form.cboKlasa.Text & "') AND (KL_INDEKSI = '" & active_form.cboIndeksi & "') AND (KL_VITI_SHKOLLOR = '" & active_form.cboVitiShkollor.Text & "') AND (AMZA_LARGUAR = 0) ORDER BY AMZA_EMRI, AMZA_MBIEMRI "
         '& active_form.cboVitiShkollor.Text & "') AND (AMZA_INDEKSI = '" & active_form.cboIndeksi & _
         '"') AND (AMZA_KLASA = '" & active_form.cboKlasa.Text & "') ORDER BY AMZA_EMRI, AMZA_MBIEMRI"
         tmp.Open q_res1, conn, adOpenDynamic, adLockOptimistic
         NumriNxenesve = active_form.emrat.Rows
         i = 0
         Do While Not tmp.EOF
            'Do merren te dhenat e nevojshme per printimin e defteses te cilat
            'do i hedhe ne nje tabele sipas kesaj renditjeje:
            ' Rreshti 0: Emri i nxenesit
            ' Rreshti 1: Mbiemri i nxenesit
            ' Rreshti 2: Atesia
            ' Rreshti 3: Numri i amzes
            ' Rreshti 4: Vrejtje
            ' rreshti 5: emri i shkolles
            ' rreshti 6: vendlindja
            ' rreshti 7: rrethi, qyteti ku ndodhet shkolla
            ' rreshti 8: datelindja e nxenesit
            ' rreshti 9: viti shkollor
            ' rreshti 10: klasa
            numriAmzes = tmp.Fields("AMZA_NR_AMZA")
            cikli = tmp.Fields("AMZA_CIKLI")
            Dim Larguar          As Integer
            If (tmp.Fields("AMZA_LARGUAR") = False) Then
            Larguar = 0
            Else
            Larguar = 1
            End If
            
           
            Dim strC As Integer
           If (ktheCiklin(active_form.cboKlasa, CStr(active_form.cboVitiShkollor.Text)) = False) Then
           strC = 0
           Else
           strC = 1
           End If
            
            q_res2 = "SELECT NT_NR_AMZA, NT_EMRI_LENDA, NT_VLERESIMI, NT_SEMESTRI1, NT_SEMESTRI2, NT_VJETORE, NT_RIPROVIM" & _
               " From TBL_NOTA,TBL_LENDA WHERE (NT_VITI_SHKOLLOR = '" & active_form.cboVitiShkollor & _
               "') AND (NT_KLASA = '" & active_form.cboKlasa & "') AND (NT_INDEKSI = '" & active_form.cboIndeksi & _
               "') AND (NT_NR_AMZA = '" & numriAmzes & "' AND NT_CIKLI = " & cikli & ")"
            q_res3 = " AND (NT_EMRI_LENDA = LN_EMRI) AND (LN_KLASA" & active_form.cboKlasa & " = 1) AND (" & Larguar & " = 0) AND NT_CIKLI = " & strC & " AND (NT_SEMESTRI1 = 1 OR NT_SEMESTRI2 = 1 OR NT_VJETORE  = 1 OR NT_RIPROVIM = 1) ORDER BY NT_EMRI_LENDA"
            TeDhenatVecanta(i).TeDhenatSpecifike(0) = tmp.Fields("AMZA_EMRI")
            TeDhenatVecanta(i).TeDhenatSpecifike(1) = tmp.Fields("AMZA_MBIEMRI")
            TeDhenatVecanta(i).TeDhenatSpecifike(2) = tmp.Fields("AMZA_ATESIA")
            TeDhenatVecanta(i).TeDhenatSpecifike(3) = tmp.Fields("AMZA_NR_AMZA")
            q_res4 = "SELECT SJ_SJELLJE FROM TBL_SJELLJE, TBL_AMZA WHERE (SJ_NR_AMZA = '" & numriAmzes & "') AND (SJ_VITI_SHKOLLOR = '" & active_form.cboVitiShkollor & "') AND (SJ_CIKLI = " & cikli & ")"
            tmp111.Open q_res4, conn, adOpenDynamic, adLockOptimistic
            If Not tmp111.EOF Then
               If Not IsNull(tmp111.Fields("SJ_SJELLJE")) Then
                  TeDhenatVecanta(i).TeDhenatSpecifike(4) = tmp111.Fields("SJ_SJELLJE")
               End If
            End If
            If Not tmp111.EOF Then
               tmp111.MoveNext
            End If
            tmp111.Close
            TeDhenatVecanta(i).TeDhenatSpecifike(5) = emerShkolla
            TeDhenatVecanta(i).TeDhenatSpecifike(6) = tmp.Fields("AMZA_VENDLINDJA")
            TeDhenatVecanta(i).TeDhenatSpecifike(7) = adresaShkolla
            TeDhenatVecanta(i).TeDhenatSpecifike(8) = tmp.Fields("AMZA_DATELINDJA")
            TeDhenatVecanta(i).TeDhenatSpecifike(9) = active_form.cboVitiShkollor.Text
            TeDhenatVecanta(i).TeDhenatSpecifike(10) = active_form.cboKlasa.Text

            ' Hidhen mungesat e nxenesve
            q_res = "SELECT NT_MUNGESE_ME, NT_MUNGESE_PA, NT_MOMENTALES1, NT_MOMENTALES2 FROM TBL_NOTA WHERE (NT_NR_AMZA = '" _
               & numriAmzes & "') AND (NT_CIKLI = " & cikli & ") AND (NT_MUNGESE_ME = 1 OR NT_MUNGESE_PA = 1)"
            mungeseMeS1 = 0
            mungeseMeS2 = 0
            mungesePaS1 = 0
            MungesePaS2 = 0
            rekset.Open q_res, conn, adOpenDynamic, adLockOptimistic
            If Not rekset.EOF Then
               Do While Not rekset.EOF
                  If rekset.Fields("NT_MUNGESE_ME") = True And rekset.Fields("NT_MOMENTALES1") Then
                     mungeseMeS1 = mungeseMeS1 + 1
                  ElseIf rekset.Fields("NT_MUNGESE_ME") = True And rekset.Fields("NT_MOMENTALES2") Then
                     mungeseMeS2 = mungeseMeS2 + 1
                  ElseIf rekset.Fields("NT_MUNGESE_PA") = True And rekset.Fields("NT_MOMENTALES1") Then
                     mungesePaS1 = mungesePaS1 + 1
                  ElseIf rekset.Fields("NT_MUNGESE_PA") = True And rekset.Fields("NT_MOMENTALES2") Then
                     MungesePaS2 = MungesePaS2 + 1
                  End If
                  rekset.MoveNext
               Loop
            End If
            rekset.Close
            TeDhenatVecanta(i).TeDhenatSpecifike(11) = mungeseMeS1
            TeDhenatVecanta(i).TeDhenatSpecifike(12) = mungeseMeS2
            TeDhenatVecanta(i).TeDhenatSpecifike(13) = mungesePaS1
            TeDhenatVecanta(i).TeDhenatSpecifike(14) = MungesePaS2

            ' Pas hedhjes se gjeneraliteteve hedhim notat per cdo nxenes
            tmp11.Open q_res2 + q_res3, conn, adOpenDynamic, adLockOptimistic
            j = 0
            If Not tmp11.EOF Then
               LendaParaardhese = tmp11.Fields("NT_EMRI_LENDA")
            End If
            Do While Not tmp11.EOF
               TeDhenatVecanta(i).TabelNotave(j, 0) = tmp11.Fields("NT_EMRI_LENDA")
               If tmp11.Fields("NT_SEMESTRI1") = True Then
                  TeDhenatVecanta(i).TabelNotave(j, 1) = tmp11.Fields("NT_VLERESIMI")
               ElseIf tmp11.Fields("NT_SEMESTRI2") = True Then
                  TeDhenatVecanta(i).TabelNotave(j, 2) = tmp11.Fields("NT_VLERESIMI")
               ElseIf tmp11.Fields("NT_VJETORE") = True Then
                  TeDhenatVecanta(i).TabelNotave(j, 3) = tmp11.Fields("NT_VLERESIMI")
               ElseIf tmp11.Fields("NT_RIPROVIM") = True Then
                  TeDhenatVecanta(i).TabelNotave(j, 4) = tmp11.Fields("NT_VLERESIMI")
               End If
               tmp11.MoveNext
               If Not tmp11.EOF Then
                  If tmp11.Fields("NT_EMRI_LENDA") <> LendaParaardhese Then
                     j = j + 1
                  End If
                  LendaParaardhese = tmp11.Fields("NT_EMRI_LENDA")
               End If
            Loop
            tmp11.Close
            tmp.MoveNext
            i = i + 1
         Loop
         tmp.Close

         ' Klasa eshte ne shkollen e mesme
         Set objPrintim = CreateObject("PrintimComponent.clsFormatDeftesatSeri")
         If CInt(active_form.cboKlasa.Text) >= 8 Then
            nResult = MsgBox("Futini deftesat ne menyre qe te printohen gjeneralitetet", vbInformation + vbOKCancel, "Printimi i deftesave")
            If nResult = vbOK Then
               objPrintim.OrientimFaqe False
               If objPrintim.Gabim = True Then
                  MsgBox "Printeri juaj nuk lejon qe formati i printerit tuaj te nderrohet ne " _
                     & "menyre automatike nga Portrait ne Landscape. Ju lutemi nderrojeni formatin " _
                     & "ne menyre manuale tek preferencat e printerit tuaj.", vbOKOnly + vbCritical, "Gabim ne printim"
                  Exit Sub
               End If
               objPrintim.PaperSizeA4 False
               objPrintim.vitishkollor = active_form.cboVitiShkollor.Text
               If objPrintim.Gabim = True Then
                  MsgBox "Printeri juaj nuk mund te printoje dot ne format A3, i " _
                     & "nevojshem per printimin e deftesave te shkolles se mesme", _
                     vbOKOnly + vbCritical, "Gabim ne printim"
               End If
               For i = 0 To NumriNxenesve - 1
                  objPrintim.DeftesaMesmeSeriFPa TeDhenatVecanta(i).TeDhenatSpecifike
                  ' Nese perdoruesi anullon printimin
                  If objPrintim.Gabim = True Then
                     Exit Sub
                  End If
                  objPrintim.NewPage
               Next
               objPrintim.EndDoc
               nResult = MsgBox("Futini deftesat ne menyre qe te printohen notat", vbInformation + vbOKCancel, "Printimi i deftesave")
               If nResult = vbOK Then
                  objPrintim.OrientimFaqe False
                  If objPrintim.Gabim = True Then
                     MsgBox "Printeri juaj nuk lejon qe formati i printerit tuaj te nderrohet ne " _
                        & "menyre automatike nga Portrait ne Landscape. Ju lutemi nderrojeni formatin " _
                        & "ne menyre manuale tek preferencat e printerit tuaj.", vbOKOnly + vbCritical, "Gabim ne printim"
                     Exit Sub
                  End If
                  objPrintim.PaperSizeA4 False
                  If objPrintim.Gabim = True Then
                     MsgBox "Printeri juaj nuk mund te printoje dot ne format A3, i " _
                        & "nevojshem per printimin e deftesave te shkolles se mesme", _
                        vbOKOnly + vbCritical, "Gabim ne printim"
                  End If
                  For i = 0 To NumriNxenesve - 1
                     objPrintim.DeftesaMesmeSeriFPr TeDhenatVecanta(i).TabelNotave, TeDhenatVecanta(i).TeDhenatSpecifike
                     ' Nese perdoruesi anullon printimin
                     If objPrintim.Gabim = True Then
                        Exit Sub
                     End If
                     objPrintim.NewPage
                  Next
                  objPrintim.EndDoc
               Else
                  Exit Sub
               End If
            Else
               Exit Sub
            End If
            ' Klasa eshte ne shkollen tetevjecare
         Else
            ' Printimi i gjeneraliteteve
            nResult = MsgBox("Futini deftesat ne menyre qe te printohen gjeneralitetet", vbInformation + vbOKCancel, "Printimi i deftesave")
            If nResult = vbOK Then
               objPrintim.OrientimFaqe True
               objPrintim.PaperSizeA4 True
               objPrintim.vitishkollor = active_form.cboVitiShkollor.Text
               For i = 0 To NumriNxenesve - 1
                  objPrintim.DeftesaUletSeriFPa TeDhenatVecanta(i).TeDhenatSpecifike
                  ' Nese perdoruesi anullon printimin
                  If objPrintim.Gabim = True Then
                     Exit Sub
                  End If
                  objPrintim.NewPage
               Next
               objPrintim.EndDoc
            Else
               Exit Sub
            End If
            ' Printimi i notave
            nResult = MsgBox("Futini deftesat ne menyre qe te printohen notat", vbInformation + vbOKCancel, "Printimi i deftesave")
            If nResult = vbOK Then
               objPrintim.OrientimFaqe True
               objPrintim.PaperSizeA4 True
               For i = 0 To NumriNxenesve - 1
                  objPrintim.DeftesaUletSeriFPr TeDhenatVecanta(i).TabelNotave, TeDhenatVecanta(i).TeDhenatSpecifike
                  ' Nese perdoruesi anullon printimin
                  If objPrintim.Gabim = True Then
                     Exit Sub
                  End If
                  objPrintim.NewPage
               Next
               objPrintim.EndDoc
            Else
               Exit Sub
            End If
         End If

      Case HIQ_LENDE

         Klasa = active_form.cboKlasa.Text
         nrVitiShkollor = active_form.cboVitiShkollor.Text
         If active_form.optMesme.Value Then
            strCikli = "1"
         Else
            strCikli = "0"
         End If
         Dim gjatesia_liste   As Integer
         q_res = "SELECT NT_EMRI_LENDA FROM TBL_NOTA WHERE NT_KLASA= '" & Klasa & "' AND NT_VITI_SHKOLLOR = '" & nrVitiShkollor & "'"
         tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
         gjatesia_liste = active_form.lista2.ListCount
         If tmp.EOF Then
            i = 0
            Do While (i <= gjatesia_liste - 1)
               If active_form.lista2.Selected(i) Then

                  active_form.lista2.RemoveItem (i)
                  gjatesia_liste = gjatesia_liste - 1
                  active_form.listaNumri.RemoveItem (i)
                  i = i - 1
               End If
               i = i + 1
            Loop
         Else
            i = 0
            Do While (i <= gjatesia_liste - 1)
               If active_form.lista2.Selected(i) Then
                  lenda = active_form.lista2.List(i)
                  ugjet = False
                  If Not tmp.EOF Then
                     tmp.MoveFirst
                     Do While Not tmp.EOF And Not ugjet
                        If tmp.Fields("NT_EMRI_LENDA") = lenda Then
                           ugjet = True
                        End If
                        tmp.MoveNext
                     Loop
                  Else
                     ugjet = False
                  End If
                  If ugjet Then
                     j = MsgBox("Per lenden ' " & lenda & " ' jane hedhur nota ose mungesa ne regjister." & Chr(10) & "Fshirja e lendes do te fshije edhe vleresimet e marra ne kete lende." & Chr(10) & "Doni te vazhdoni ?", vbExclamation + vbOKCancel, "Konfigurimi i lendeve.")
                     If j = vbOK Then
                        active_form.lista2.RemoveItem (i)
                        gjatesia_liste = gjatesia_liste - 1
                        active_form.listaNumri.RemoveItem (i)
                        i = i - 1
                        q_res1 = "DELETE FROM TBL_NOTA WHERE NT_EMRI_LENDA = '" & lenda & "' AND NT_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND NT_KLASA = '" & Klasa & "'"
                        conn.Execute q_res1
                        q_res1 = "UPDATE TBL_LENDA SET LN_KLASA" & Klasa & " = 0 WHERE LN_EMRI = '" & lenda & "' AND LN_VITI_SHKOLLOR = '" & nrVitiShkollor & "'"
                        conn.Execute q_res1
                     End If

                  Else
                     q_res1 = "UPDATE TBL_LENDA SET LN_KLASA" & Klasa & " = 0 WHERE LN_EMRI = '" & lenda & "' AND LN_VITI_SHKOLLOR = '" & nrVitiShkollor & "'"
                     conn.Execute q_res1
                     active_form.lista2.RemoveItem (i)
                     active_form.listaNumri.RemoveItem (i)
                     gjatesia_liste = gjatesia_liste - 1
                     i = i - 1
                  End If

               End If
               i = i + 1
            Loop
         End If
         tmp.Close

      Case MERR_MUNGESA
         'Dim mungeseMeS1, mungeseMeS2, mungesePaS1, MungesePaS2 As Integer
         If active_form.optMesme = True Then
            cikli = 1
         Else
            cikli = 0
         End If
         q_res = "SELECT NT_MUNGESE_ME, NT_MUNGESE_PA, NT_MOMENTALES1, NT_MOMENTALES2 FROM TBL_NOTA WHERE (NT_NR_AMZA = '" _
            & active_form.txtAmzaNo.Text & "') AND (NT_CIKLI = " & cikli & ") AND (NT_MUNGESE_ME = 1 OR NT_MUNGESE_PA = 1)"
         mungeseMeS1 = 0
         mungeseMeS2 = 0
         mungesePaS1 = 0
         MungesePaS2 = 0
         tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
         If Not tmp.EOF Then
            Do While Not tmp.EOF
               If tmp.Fields("NT_MUNGESE_ME") = True And tmp.Fields("NT_MOMENTALES1") Then
                  mungeseMeS1 = mungeseMeS1 + 1
               ElseIf tmp.Fields("NT_MUNGESE_ME") = True And tmp.Fields("NT_MOMENTALES2") Then
                  mungeseMeS2 = mungeseMeS2 + 1
               ElseIf tmp.Fields("NT_MUNGESE_PA") = True And tmp.Fields("NT_MOMENTALES1") Then
                  mungesePaS1 = mungesePaS1 + 1
               ElseIf tmp.Fields("NT_MUNGESE_PA") = True And tmp.Fields("NT_MOMENTALES2") Then
                  MungesePaS2 = MungesePaS2 + 1
               End If
               tmp.MoveNext
            Loop
         End If
         If active_form.Name = "frmKonsultimeNota" Then
            active_form.MungesaS1Me = mungeseMeS1
            active_form.MungesaS1Pa = mungesePaS1
            active_form.MungesaS2Me = mungeseMeS2
            active_form.MungesaS2Pa = mungeseMeS2
         End If

   End Select
End Sub
Public Sub andiTjeter()
   Dim vleresimi        As String
   Dim a                As Integer
   Dim b                As Integer
   Dim q_res2           As String
   Dim s1, s2, q_res, q_res1, q_res11, s3 As String
   Dim numerAmze        As String
   Dim cikli            As Integer
   Dim strCikli         As String
   Dim emri             As String
   Dim mbiemri          As String
   Dim nrVitiShkollor   As String
   Dim Klasa            As String
   Dim indeksi          As String
   Dim atesia           As String
   Dim memesia          As String
   Dim Vendlindja       As String
   Dim Datelindja       As Date
   Dim seksi            As String
   Dim j                As Integer
   Dim nr_lendesh       As Integer
   Dim nr_max_notash    As Integer
   Dim Vrejtje          As String
   Dim mungeseMeS1, mungeseMeS2, mungesePaS1, MungesePaS2 As Integer
   Dim conn             As New ADODB.Connection
   
   With conn
      .Provider = "Microsoft.Access.OLEDB.10.0"
      .Properties("Data Provider").Value = "SQLOLEDB"
      .Properties("Data Source").Value = EmerServer + "\SHM"
      .Properties("User ID").Value = "sa"
      .Properties("Password").Value = "vision"
      .Properties("Initial Catalog").Value = "DProgNec"
      If (conn.state = 1) Then
        .Close
      End If
      .Open
   End With
   
  ' conn.ConnectionString = getConnectionString()
   'conn.Open
   Dim tmp              As New ADODB.Recordset
   Dim tmp11            As New ADODB.Recordset
   Dim tmp111           As New ADODB.Recordset
   Dim rekset           As New ADODB.Recordset
   Select Case Me.actionName
  
        Case KALO_NE_LISTE_NXENES
      'Dim gjatesia         As Integer
      Dim vjetore          As Boolean
      Dim riprovim         As Boolean
      Dim rreshti_emri     As String
      Dim ngelet           As Boolean
      Dim klasa_e_re       As String
      Dim nr_viti_shkollor_i_ri As String
      b = active_form.lista1.ListCount - 1
      a = 0
      If active_form.optMesme.Value Then
         strCikli = "1"
      Else
         strCikli = "0"
      End If
      Dim tmp1 As New ADODB.Recordset
      Dim ugjet As Boolean
      Dim i As Integer
      Dim lenda As String
      Klasa = active_form.cboKlasa.Text
      indeksi = active_form.cboIndeksi.Text
      nrVitiShkollor = active_form.cboVitiShkollor.Text
      klasa_e_re = active_form.cboKlasa2.Text
      nr_viti_shkollor_i_ri = active_form.cboVitiShkollor2.Text
      q_res1 = "SELECT LN_EMRI FROM TBL_LENDA WHERE  LN_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND LN_KLASA" & Klasa & " =  1 "
      tmp1.Open q_res1, conn, adOpenDynamic, adLockOptimistic
      If tmp1.EOF Then
         b = active_form.lista1.ListCount - 1
         a = 0
         Do While a <= b
            If active_form.lista1.Selected(a) Then
               numerAmze = active_form.listaAmza.List(a)
               emri = active_form.lista1.List(a)

               ugjet = False
               i = 0
               Do While i <= active_form.lista2Amza.ListCount And Not ugjet
                  If active_form.lista2Amza.List(i) = numerAmze Then
                     ugjet = True
                  End If
                  i = i + 1
               Loop
               If Not ugjet Then
                  active_form.lista1.RemoveItem (a)
                  active_form.listaAmza.RemoveItem (a)
                  active_form.lista2Amza.AddItem numerAmze
                  active_form.lista2.AddItem emri
                  b = b - 1
                  a = a - 1
               End If
            End If
            a = a + 1
         Loop
         Exit Sub
      End If
      If Klasa = klasa_e_re Then
         b = active_form.lista1.ListCount - 1
         a = 0
         Do While a <= b
            If active_form.lista1.Selected(a) Then
               numerAmze = active_form.listaAmza.List(a)
               emri = active_form.lista1.List(a)

               ugjet = False
               i = 0
               Do While i <= active_form.lista2Amza.ListCount And Not ugjet
                  If active_form.lista2Amza.List(i) = numerAmze Then
                     ugjet = True
                  End If
                  i = i + 1
               Loop
               If Not ugjet Then
                  active_form.lista1.RemoveItem (a)
                  active_form.listaAmza.RemoveItem (a)
                  active_form.lista2Amza.AddItem numerAmze
                  active_form.lista2.AddItem emri
                  b = b - 1
                  a = a - 1
               End If
            End If
            a = a + 1
         Loop
         Exit Sub
      End If
      q_res = "SELECT NT_NR_AMZA, NT_VLERESIMI, NT_EMRI_LENDA, NT_VJETORE, NT_RIPROVIM FROM TBL_NOTA WHERE (NT_CIKLI = " & strCikli & " AND NT_KLASA = '" & Klasa & "' AND NT_INDEKSI = '" & indeksi & "' AND NT_VITI_SHKOLLOR = '" & nrVitiShkollor & "') AND ( NT_VJETORE = 1 OR NT_RIPROVIM = 1 )  ORDER BY NT_VJETORE DESC, NT_RIPROVIM DESC"
      tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
      If tmp.EOF Then
         b = active_form.lista1.ListCount - 1
         a = 0
         Do While a <= b
            If active_form.lista1.Selected(a) Then
               numerAmze = active_form.listaAmza.List(a)
               emri = active_form.lista1.List(a)

               ugjet = False
               i = 0
               Do While i <= active_form.lista2Amza.ListCount And Not ugjet
                  If active_form.lista2Amza.List(i) = numerAmze Then
                     ugjet = True
                  End If
                  i = i + 1
               Loop
               If Not ugjet Then
                  active_form.lista1.RemoveItem (a)
                  active_form.listaAmza.RemoveItem (a)
                  active_form.lista2Amza.AddItem numerAmze
                  active_form.lista2.AddItem emri
                  b = b - 1
                  a = a - 1
               End If
            End If
            a = a + 1
         Loop
         Exit Sub
      End If

      b = active_form.lista1.ListCount - 1
      a = 0
      Do While a <= b
         If active_form.lista1.Selected(a) Then
            numerAmze = active_form.listaAmza.List(a)
            emri = active_form.lista1.List(a)
            ngelet = False
            tmp1.MoveFirst

            Do While Not tmp1.EOF And Not ngelet
               lenda = tmp1.Fields("LN_EMRI")
               vjetore = True
               riprovim = True
               tmp.MoveFirst
               ugjet = False
               tmp.MoveFirst
               Do While Not tmp.EOF And Not ngelet
                  If numerAmze = tmp.Fields("NT_NR_AMZA") And lenda = tmp.Fields("NT_EMRI_LENDA") Then
                     If tmp.Fields("NT_VJETORE") Then
                        If tmp.Fields("NT_VLERESIMI") = "4" Then
                           vjetore = False
                           riprovim = False
                        End If
                     End If
                     If tmp.Fields("NT_RIPROVIM") Then
                        If tmp.Fields("NT_VLERESIMI") = "4" Then

                           riprovim = False
                        Else
                           riprovim = True
                        End If
                     End If
                  End If
                  tmp.MoveNext

               Loop
               If Not riprovim Then
                  ngelet = True
               End If
               tmp1.MoveNext
            Loop
            If Not ngelet Then
               i = 0
               ugjet = False
               Do While i <= active_form.lista2Amza.ListCount And Not ugjet
                  If active_form.lista2Amza.List(i) = numerAmze Then
                     ugjet = True
                  End If
                  i = i + 1
               Loop
               If Not ugjet Then
                  active_form.lista1.RemoveItem (a)
                  active_form.listaAmza.RemoveItem (a)
                  active_form.lista2Amza.AddItem numerAmze
                  active_form.lista2.AddItem emri
                  b = b - 1
                  a = a - 1
               End If
            Else
               MsgBox "Nxenesi ' " & emri & "' nuk mund te kaloje ne klasen pasardhese," & Chr(10) & "sepse nuk eshte kalues .", vbExclamation, "Konfigurimi i klasave."
            End If

         End If
         a = a + 1
      Loop
       

      Case LARGO_NXENES

         numerAmze = active_form.txtAmzaNo.Text
         If active_form.optMesme.Value = True Then
            strCikli = "1"
         Else
            strCikli = "0"
         End If
         nrVitiShkollor = active_form.txtVitiShkollor.Text
         Dim data_e_largimit
         data_e_largimit = active_form.Calendar1.Value

         q_res = "UPDATE TBL_AMZA SET AMZA_LARGUAR = 1 WHERE AMZA_NR_AMZA = '" & numerAmze & "' AND AMZA_CIKLI = " & strCikli
         conn.Execute q_res
         Vrejtje = "Larguar ne daten " & data_e_largimit & " ."
         q_res = "UPDATE TBL_AMZA SET AMZA_VREJTJE = '" & Vrejtje & "',AMZA_VITI_LARGIMIT = '" & nrVitiShkollor & "'  WHERE AMZA_NR_AMZA = '" & numerAmze & "' AND AMZA_CIKLI = " & strCikli
         conn.Execute q_res
         MsgBox "Nxenesi me te dhenat e mesiperme u largua nga shkolla.", vbInformation, "Elimino nxenes."

         active_form.txtAmzaNo.Text = ""
         active_form.txtEmri.Text = ""
         active_form.txtMbiemri.Text = ""
         active_form.txtKlasa.Text = ""
         active_form.txtIndeksi.Text = ""
         active_form.txtVitiShkollor.Text = ""
         active_form.txtAtesia.Text = ""
         active_form.txtMemesia.Text = ""
         active_form.txtSeksi.Text = ""
         active_form.txtVendlindja.Text = ""
         active_form.txtDatelindja.Text = ""

         active_form.optMesme.Value = False
         active_form.optTetevjecare.Value = False
      Case SHFAQ_NOTA

         'Dim nota             As String
         '*** hedh notat e nje nxenesi ne gride
         nrVitiShkollor = active_form.cboVitiShkollor.Text

         '*** ketu percaktohet cikli i nxenesit

         If active_form.optMesme.Value Then
            cikli = 1
            strCikli = "1"
         Else
            cikli = 0
            strCikli = "0"
         End If

         '***
         '*** ketu kryhet hedhja e te dhenave plotesuese te nxenesit
         If active_form.txtAmzaNo.Text <> "" Then
            numerAmze = active_form.txtAmzaNo.Text

            q_res = " SELECT AMZA_NR_AMZA, AMZA_EMRI, AMZA_MBIEMRI, AMZA_ATESIA, AMZA_DATELINDJA, AMZA_VREJTJE, AMZA_VENDLINDJA,KL_KLASA,KL_INDEKSI, KL_VITI_SHKOLLOR FROM TBL_AMZA, TBL_KLASA WHERE ( AMZA_NR_AMZA = KL_NR_AMZA AND AMZA_CIKLI = KL_CIKLI) AND (AMZA_NR_AMZA = '" & numerAmze & "' AND AMZA_CIKLI = " & strCikli & ") AND KL_VITI_SHKOLLOR = '" & nrVitiShkollor & "' ORDER BY AMZA_EMRI, AMZA_MBIEMRI"
            tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic

            q_res1 = "SELECT SJ_SJELLJE FROM TBL_SJELLJE WHERE (SJ_NR_AMZA = '" & _
               numerAmze & "') AND (SJ_CIKLI = " & cikli & ") AND (SJ_VITI_SHKOLLOR = '" & nrVitiShkollor & "')"
            tmp11.Open q_res1, conn, adOpenDynamic, adLockOptimistic

            'get_rec (q_res)
            If Not tmp.EOF Then
               active_form.txtEmri.Text = tmp.Fields("AMZA_EMRI")
               active_form.txtMbiemri.Text = tmp.Fields("AMZA_MBIEMRI")
               active_form.txtAtesia = tmp.Fields("AMZA_ATESIA")
               active_form.txtKlasa.Text = tmp.Fields("KL_KLASA")
               active_form.txtIndeksi.Text = tmp.Fields("KL_INDEKSI")

               If active_form.Name = "frmKonsultimeNota" Then
                  If Not IsNull(tmp.Fields("PERD_QYTETI")) Then
                     active_form.Qyteti = tmp.Fields("PERD_QYTETI")
                  End If
                  If Not IsNull(tmp.Fields("PERD_ADRESA")) Then
                     active_form.Adresa = tmp.Fields("PERD_ADRESA")
                  End If
                  active_form.Vendlindja = tmp.Fields("AMZA_VENDLINDJA")
                  active_form.Datelindja = tmp.Fields("AMZA_DATELINDJA")
                  If Not IsNull(tmp11.Fields("SJ_SJELLJE")) Then

                     active_form.Vrejtje = tmp11.Fields("SJ_SJELLJE")
                  End If
               End If
               tmp.Close
               tmp11.Close
Else:
               tmp.Close
               tmp11.Close
               active_form.txtEmri.Text = ""
               active_form.txtMbiemri.Text = ""
               active_form.txtAtesia.Text = ""
               active_form.txtKlasa.Text = ""
               active_form.txtIndeksi.Text = ""
               active_form.lblLargimi.Text = ""
               ' Atesia, Vendlindja, Datelindja, Vrejtje,  sherbejne vetem per
               ' printimin e defteses

               active_form.Label3.Visible = False
               active_form.txtData.Visible = False
               active_form.Cgrid1.Visible = False
               active_form.Lendet.Visible = False
               active_form.Label2.Visible = False
               active_form.fsbHorizontali.Visible = False
               active_form.lblNotatDheMungesat.Visible = False
               scrbar = False
               scrbarver = False
               MsgBox "   Nxenesi me numrin e amzes se dhene nuk eshte i regjistruar ne vitin shkollor qe eshte dhene" & vbCrLf & "   dhe ne ciklin qe eshte zgjedhur", vbInformation, "  Konsultime Nota"
               active_form.txtAmzaNo.Text = ""
               Exit Sub
            End If
         Else
            emri = active_form.txtEmri.Text
            mbiemri = active_form.txtMbiemri.Text
            atesia = active_form.txtAtesia.Text
            q_res = " SELECT AMZA_NR_AMZA, AMZA_EMRI, AMZA_MBIEMRI, AMZA_ATESIA, AMZA_DATELINDJA, AMZA_VREJTJE, AMZA_VENDLINDJA,KL_KLASA,KL_INDEKSI, AMZA_CIKLI, KL_VITI_SHKOLLOR FROM TBL_AMZA, TBL_KLASA WHERE (AMZA_NR_AMZA = KL_NR_AMZA AND AMZA_CIKLI = KL_CIKLI ) AND ( AMZA_EMRI LIKE '" & emri & "%' AND AMZA_MBIEMRI LIKE '" & mbiemri & "%' AND AMZA_ATESIA LIKE '" & atesia & "%' AND AMZA_CIKLI = " & strCikli & ")  AND KL_VITI_SHKOLLOR = '" & nrVitiShkollor & "' ORDER BY AMZA_EMRI, AMZA_MBIEMRI"
            'get_rec (q_res)
            tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
            If Not tmp.EOF Then
                If tmp.RecordCount = 1 Then
                    active_form.txtEmri.Text = tmp.Fields("AMZA_EMRI")
                    active_form.txtMbiemri.Text = tmp.Fields("AMZA_MBIEMRI")
                    active_form.txtAtesia = tmp.Fields("AMZA_ATESIA")
                    active_form.txtAmzaNo.Text = tmp.Fields("AMZA_NR_AMZA")
                    active_form.txtKlasa.Text = tmp.Fields("KL_KLASA")
                    active_form.txtIndeksi.Text = tmp.Fields("KL_INDEKSI")
                    tmp.Close
                    active_form.Label3.Visible = True
                    active_form.txtData.Visible = True
                Else
                    'nqs ka me teper se nje nxenes shfaq listen
                    'inicializo gridat ne forme
                    Dim nrRreshta As Integer
                    nrRreshta = tmp.RecordCount
                    Call inicializoGridaZgjidhNxenes(nrRreshta)
                    'mbush griden me te dhena
                    Dim Index As Integer
                    Index = 0
                    Do While Not tmp.EOF
                        frmZgjidhNxenes.gridaNxenesit.row = Index
                        frmZgjidhNxenes.gridaNxenesit.col = 0
                        frmZgjidhNxenes.gridaNxenesit.Text = tmp.Fields("AMZA_NR_AMZA").Value
                        frmZgjidhNxenes.gridaNxenesit.col = 1
                        frmZgjidhNxenes.gridaNxenesit.Text = tmp.Fields("AMZA_EMRI").Value + " " + tmp.Fields("AMZA_ATESIA").Value + " " + tmp.Fields("AMZA_MBIEMRI").Value
                        frmZgjidhNxenes.gridaNxenesit.col = 2
                        frmZgjidhNxenes.gridaNxenesit.Text = tmp.Fields("AMZA_CIKLI").Value
                        tmp.MoveNext
                        Index = Index + 1
                    Loop
                    tmp.Close
                    ShkollaManager.frmZgjidhNxenes.show vbModal
                    numerAmze = ShkollaManager.frmZgjidhNxenes.Amza
                    strCikli = ShkollaManager.frmZgjidhNxenes.cikli
                    If (numerAmze = "" Or strCikli = "") Then
                        'pastro
                         active_form.txtEmri.Text = ""
                        active_form.txtMbiemri.Text = ""
                        active_form.txtAtesia.Text = ""
                        active_form.lblLargimi.Text = ""
                        active_form.Label3.Visible = False
                        active_form.txtData.Visible = False
                        active_form.Cgrid1.Visible = False
                        active_form.Lendet.Visible = False
                        active_form.Label2.Visible = False
                        active_form.fsbHorizontali.Visible = False
                        active_form.lblNotatDheMungesat.Visible = False
                        Exit Sub
                    End If
                    q_res = " SELECT AMZA_NR_AMZA, AMZA_EMRI, AMZA_MBIEMRI, AMZA_ATESIA, AMZA_DATELINDJA, AMZA_VREJTJE, AMZA_VENDLINDJA,KL_KLASA,KL_INDEKSI, KL_VITI_SHKOLLOR FROM TBL_AMZA, TBL_KLASA WHERE ( AMZA_NR_AMZA = KL_NR_AMZA AND AMZA_CIKLI = KL_CIKLI) AND (AMZA_NR_AMZA = '" & numerAmze & "' AND AMZA_CIKLI = " & strCikli & ") AND KL_VITI_SHKOLLOR = '" & nrVitiShkollor & "' ORDER BY AMZA_EMRI, AMZA_MBIEMRI"
                    tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
                    If Not tmp.EOF Then
                        active_form.txtAmzaNo.Text = tmp.Fields("AMZA_NR_AMZA")
                        active_form.txtKlasa.Text = tmp.Fields("KL_KLASA")
                        active_form.txtIndeksi.Text = tmp.Fields("KL_INDEKSI")
                        active_form.txtEmri.Text = tmp.Fields("AMZA_EMRI")
                        active_form.txtMbiemri.Text = tmp.Fields("AMZA_MBIEMRI")
                        active_form.txtAtesia = tmp.Fields("AMZA_ATESIA")
                        active_form.Label3.Visible = True
                        active_form.txtData.Visible = True
                        tmp.Close
                    End If
                End If
               
Else:
               tmp.Close
               active_form.txtEmri.Text = ""
               active_form.txtMbiemri.Text = ""
               active_form.txtAtesia.Text = ""
               active_form.lblLargimi.Text = ""
               active_form.Label3.Visible = False
               active_form.txtData.Visible = False
               active_form.Cgrid1.Visible = False
               active_form.Lendet.Visible = False
               active_form.Label2.Visible = False
               active_form.fsbHorizontali.Visible = False
               active_form.lblNotatDheMungesat.Visible = False
               MsgBox "   Nxenesi me kete emer dhe mbiemer  nuk eshte i regjistruar ne vitin shkollor qe eshte dhene" & vbCrLf & "   dhe ne ciklin qe eshte zgjedhur", vbInformation, "  Konsultime Nota"
               Exit Sub
            End If
         End If
         numerAmze = active_form.txtAmzaNo.Text

         '***** Merren emri i shkolles dhe rrethi **********
         If active_form.Name = "frmKonsultimeNota" Then
            q_res = "SELECT PERD_EMRI_SHKOLLA, PERD_RRETHI FROM TBL_PERDORUES"
            tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
            If Not tmp.EOF Then
               If Not IsNull(tmp.Fields("PERD_ADRESA")) Then
                  active_form.Rrethi = tmp.Fields("PERD_ADRESA")
               End If
               If Not IsNull(tmp.Fields("PERD_EMRI_SHKOLLA")) Then
                  active_form.ShkEmri = tmp.Fields("PERD_EMRI_SHKOLLA")
               End If
            End If
            tmp.Close
         End If

         '***** Ketu percaktohen numri i lendeve
         'Dim nr_lendesh       As Integer
         Klasa = active_form.txtKlasa.Text
         q_res = "SELECT COUNT(LN_EMRI) AS NR_LENDESH FROM TBL_LENDA WHERE LN_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND LN_KLASA" & Klasa & " = 1 "
         tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
         nr_lendesh = tmp.Fields("NR_LENDESH")
         tmp.Close
         If nr_lendesh = 0 Then
            MsgBox "Klases se frekuentuar  nga nxenesi ne kete vit shkollor " & Chr(10) & "nuk i eshte regjistruar asnje lende.", vbExclamation, "Konsultimi i notave."
            active_form.Label3.Visible = False
            active_form.txtData.Visible = False

            Exit Sub
         End If

         '**** ketu gjendet numri maksimal i notave ne nje lende
         ' Dim nr_max_notash    As Integer
         Dim maksi1           As Integer, maksi2 As Integer

         q_res = "SELECT NT_VLERESIMI FROM TBL_NOTA WHERE NT_NR_AMZA = '" & numerAmze & "' AND NT_CIKLI = " & strCikli & " AND NT_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND NT_MOMENTALES1 = 1  "
         tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
         If tmp.EOF Then
            maksi1 = 0

            tmp.Close
         Else
            tmp.Close
            q_res = "SELECT MAX(a.NR_NOTASH) AS MAKSI FROM (SELECT  COUNT(NT_VLERESIMI) AS NR_NOTASH FROM TBL_NOTA WHERE NT_NR_AMZA = '" & numerAmze & "' AND NT_CIKLI = " & strCikli & " AND NT_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND NT_MOMENTALES1 = 1  GROUP BY NT_EMRI_LENDA) as a"
            tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
            If tmp.EOF Then
               maksi1 = 0

            Else
               maksi1 = tmp.Fields("MAKSI")
               ' nr_max_notash = nr_max_notash - 1
            End If
            tmp.Close
         End If

         q_res = "SELECT NT_VLERESIMI FROM TBL_NOTA WHERE NT_NR_AMZA = '" & numerAmze & "' AND NT_CIKLI = " & strCikli & " AND NT_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND NT_MOMENTALES2 = 1  "
         tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
         If tmp.EOF Then
            maksi2 = 0

            tmp.Close
         Else
            tmp.Close
            q_res = "SELECT MAX(A.NR_NOTASH) AS MAKSI FROM (SELECT  COUNT(NT_VLERESIMI) AS NR_NOTASH FROM TBL_NOTA WHERE NT_NR_AMZA = '" & numerAmze & "' AND NT_CIKLI = " & strCikli & " AND NT_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND NT_MOMENTALES2 = 1  GROUP BY NT_EMRI_LENDA) as A "
            tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
            If tmp.EOF Then
               maksi2 = 0

            Else
               maksi2 = tmp.Fields("MAKSI")
               'nr_max_notash = nr_max_notash - 1
            End If
            tmp.Close
         End If
         If maksi1 + maksi2 = 0 Then
            nr_max_notash = 1
         Else
            nr_max_notash = maksi1 + maksi2 + 3
         End If
         Call shfaq_grile1(nr_max_notash + 1, nr_lendesh)

         '*** ketu hidhen emrat e lendeve ne gride
         If strCikli = "1" Then
            q_res = "SELECT LN_EMRI FROM TBL_LENDA WHERE LN_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND LN_KLASA" & Klasa & " = 1  ORDER BY LN_NUMRI_MESME "
         Else
            q_res = "SELECT LN_EMRI FROM TBL_LENDA WHERE LN_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND LN_KLASA" & Klasa & " = 1  ORDER BY LN_NUMRI_ULET "
         End If
         tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
         
         i = 0
         Do While Not tmp.EOF
            i = i + 1
            active_form.Lendet.Text(i, 1) = tmp.Fields("LN_EMRI")
            tmp.MoveNext
         Loop
         If nr_max_notash < 1 Then
            Exit Sub
         End If
         Dim j1               As Integer, j2 As Integer
         q_res1 = "SELECT * FROM TBL_NOTA WHERE NT_NR_AMZA = '" & numerAmze & "' AND NT_CIKLI = " & strCikli & " AND NT_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND NT_PROVIM = 0 ORDER BY  NT_MOMENTALES1 DESC, NT_SEMESTRI1 DESC,NT_MOMENTALES2 DESC, NT_SEMESTRI2 DESC, NT_VJETORE DESC,NT_RIPROVIM DESC, NT_DATA "
         rekset.Open q_res1, conn, adOpenDynamic, adLockOptimistic
         If rekset.EOF Then
            nr_max_notash = 1
            GoTo fund_modifiko
         End If
         tmp.MoveFirst
         i = 0
        
         Do While Not tmp.EOF
            i = i + 1
            lenda = tmp.Fields("LN_EMRI")
            rekset.MoveFirst
            j1 = 0
            j2 = maksi1 + 1
            Do While Not rekset.EOF
               If lenda = rekset("NT_EMRI_LENDA") Then

                  If rekset.Fields("NT_MOMENTALES1") Then
                     active_form.Cgrid1.Text(i, j1 + 1) = rekset.Fields("NT_VLERESIMI")
                     j1 = j1 + 1
                  End If
                  If rekset.Fields("NT_SEMESTRI1") Then
                     active_form.Cgrid1.Text(i, maksi1 + 1) = rekset.Fields("NT_VLERESIMI")

                  End If

                  If rekset.Fields("NT_MOMENTALES2") Then
                     active_form.Cgrid1.Text(i, j2 + 1) = rekset.Fields("NT_VLERESIMI")
                     j2 = j2 + 1
                  End If
                  If rekset.Fields("NT_SEMESTRI2") Then
                     active_form.Cgrid1.Text(i, maksi1 + maksi2 + 2) = rekset.Fields("NT_VLERESIMI")

                  End If

                  If rekset.Fields("NT_VJETORE") Then
                     active_form.Cgrid1.Text(i, nr_max_notash) = rekset.Fields("NT_VLERESIMI")

                  End If
                  If rekset.Fields("NT_RIPROVIM") Then
                     active_form.Cgrid1.Text(i, nr_max_notash + 1) = rekset.Fields("NT_VLERESIMI")
                  End If

                  If rekset.Fields("NT_MOMENTALES1") Then
                     active_form.Cgrid1.CellForeColor(i, j1) = &H80000008
                     active_form.Cgrid1.CellFontBold(i, j1) = True
                     modifikoNotat(i, j1) = 1
                     matricaNotat(i, j1) = rekset.Fields("NT_TEKST_DATA")

                  End If

                  If rekset.Fields("NT_MOMENTALES2") Then
                     active_form.Cgrid1.CellForeColor(i, j2) = &H80000008
                     active_form.Cgrid1.CellFontBold(i, j2) = True
                     modifikoNotat(i, j2) = 1
                     matricaNotat(i, j2) = rekset.Fields("NT_TEKST_DATA")

                  End If

                  If rekset.Fields("NT_SEMESTRI1") Then
                     active_form.Cgrid1.CellForeColor(i, maksi1 + 1) = vbBlue
                     active_form.Cgrid1.CellFontBold(i, maksi1 + 1) = True
                     modifikoNotat(i, maksi1 + 1) = 2
                     matricaNotat(i, maksi1 + 1) = rekset.Fields("NT_TEKST_DATA")

                  End If

                  If rekset.Fields("NT_SEMESTRI2") Then
                     active_form.Cgrid1.CellForeColor(i, maksi1 + maksi2 + 2) = vbBlue
                     active_form.Cgrid1.CellFontBold(i, maksi1 + maksi2 + 2) = True
                     modifikoNotat(i, maksi1 + maksi2 + 2) = 3
                     matricaNotat(i, maksi1 + maksi2 + 2) = rekset.Fields("NT_TEKST_DATA")

                  End If

                  If rekset.Fields("NT_MOMENTALES1") And rekset.Fields("NT_DETYREKONTROLLI") = True Then
                     active_form.Cgrid1.CellForeColor(i, j1) = &H8000&
                     active_form.Cgrid1.CellFontBold(i, j1) = True
                     modifikoNotat(i, j1) = 1
                     matricaNotat(i, j1) = rekset.Fields("NT_TEKST_DATA")

                  End If

                  If rekset.Fields("NT_MOMENTALES2") And rekset.Fields("NT_DETYREKONTROLLI") = True Then
                     active_form.Cgrid1.CellForeColor(i, j2) = &H8000&
                     active_form.Cgrid1.CellFontBold(i, j2) = True
                     modifikoNotat(i, j2) = 1
                     matricaNotat(i, j2) = rekset.Fields("NT_TEKST_DATA")

                  End If
                  If rekset.Fields("NT_MOMENTALES1") And rekset.Fields("NT_MUNGESE_ME") = True Then
                     active_form.Cgrid1.CellForeColor(i, j1) = vbBlack
                     active_form.Cgrid1.CellFontBold(i, j1) = True
                     modifikoNotat(i, j1) = 1
                     matricaNotat(i, j1) = rekset.Fields("NT_TEKST_DATA")

                  End If

                  If rekset.Fields("NT_MOMENTALES2") And rekset.Fields("NT_MUNGESE_ME") = True Then
                     active_form.Cgrid1.CellForeColor(i, j2) = vbBlack
                     active_form.Cgrid1.CellFontBold(i, j2) = True
                     modifikoNotat(i, j2) = 1
                     matricaNotat(i, j2) = rekset.Fields("NT_TEKST_DATA")

                  End If
                  If rekset.Fields("NT_MOMENTALES1") And rekset.Fields("NT_MUNGESE_PA") = True Then
                     active_form.Cgrid1.CellForeColor(i, j1) = vbRed
                     active_form.Cgrid1.CellFontBold(i, j1) = True
                     modifikoNotat(i, j1) = 1
                     matricaNotat(i, j1) = rekset.Fields("NT_TEKST_DATA")

                  End If

                  If rekset.Fields("NT_MOMENTALES2") And rekset.Fields("NT_MUNGESE_PA") = True Then
                     active_form.Cgrid1.CellForeColor(i, j2) = vbRed
                     active_form.Cgrid1.CellFontBold(i, j2) = True
                     modifikoNotat(i, j2) = 1
                     matricaNotat(i, j2) = rekset.Fields("NT_TEKST_DATA")

                  End If

                  If rekset.Fields("NT_VJETORE") Then
                     active_form.Cgrid1.CellForeColor(i, nr_max_notash) = &HFF&
                     active_form.Cgrid1.CellFontBold(i, nr_max_notash) = True
                     modifikoNotat(i, nr_max_notash) = 4
                     matricaNotat(i, nr_max_notash) = rekset.Fields("NT_TEKST_DATA")

                  End If

                  If rekset.Fields("NT_RIPROVIM") Then
                     active_form.Cgrid1.CellForeColor(i, nr_max_notash + 1) = &HC000C0
                     active_form.Cgrid1.CellFontBold(i, nr_max_notash + 1) = True
                     modifikoNotat(i, nr_max_notash + 1) = 5
                     matricaNotat(i, nr_max_notash + 1) = rekset.Fields("NT_TEKST_DATA")

                  End If
               End If
               rekset.MoveNext
            Loop
            tmp.MoveNext
         Loop
fund_modifiko:
         rekset.Close
         tmp.Close

         q_res = "SELECT AMZA_LARGUAR FROM TBL_AMZA WHERE AMZA_NR_AMZA = '" & numerAmze & "' AND AMZA_CIKLI = " & strCikli
         tmp.Open q_res
         If Not tmp.EOF Then
            If tmp.Fields("AMZA_LARGUAR") = True Then
               active_form.lblLargimi.Text = "Nxenesi eshte larguar nga shkolla"
            End If
         End If
         tmp.Close

         For i = 1 To nr_max_notash + 1
            For j = 1 To nr_lendesh
               If Trim(active_form.Cgrid1.Text(j, i)) <> "" Then
                  If Len(Trim(active_form.Cgrid1.Text(j, i))) <> 2 Then
                     active_form.Cgrid1.Text(j, i) = Space(1) & active_form.Cgrid1.Text(j, i)
                  End If
               End If
            Next j
         Next i

         active_form.Label3.Visible = True
         active_form.txtData.Visible = True
         active_form.Lendet.Visible = True
         active_form.Cgrid1.Visible = True
         active_form.txtData.Visible = True
         active_form.Label3.Visible = True
         active_form.lblNotatDheMungesat.Visible = True

         If nr_max_notash > 42 Then
            active_form.fsbHorizontali.Visible = True
         End If
      Case KONSULTO_DATEN

         q_res = "SELECT * FROM TBL_MODIFIKIMI"
         tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
         If tmp.EOF Then
            data_modifikimi = ""
         Else
            data_modifikimi = tmp.Fields("MD_DATA")
         End If
         tmp.Close

      Case MODIFIKO_DATEN

         Dim data             As Date
         data = date
         q_res = "SELECT * FROM TBL_MODIFIKIMI"
         tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
         If tmp.EOF Then
            tmp.Close
            q_res1 = "INSERT INTO TBL_MODIFIKIMI (MD_DATA) VALUES ('" & data & "')"
            conn.Execute q_res1
         Else
            tmp.Close
            q_res1 = "UPDATE TBL_MODIFIKIMI SET MD_DATA = '" & data & "'"
            conn.Execute q_res1

         End If
      Case SHFAQ_SJELLJE
         numerAmze = active_form.txtAmzaNo.Text
         nrVitiShkollor = active_form.cboVitiShkollor.Text
         If active_form.optMesme.Value Then
            cikli = True
            strCikli = "1"
         Else
            cikli = False
            strCikli = "0"
         End If
         q_res = "SELECT SJ_SJELLJE FROM TBL_SJELLJE WHERE SJ_NR_AMZA = '" & numerAmze & "' AND SJ_CIKLI = " & strCikli & " AND SJ_VITI_SHKOLLOR = '" & nrVitiShkollor & "'"
         tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
         If tmp.EOF Then
            MsgBox "Per kete vit shkollor nxenesit nuk i eshte hedhur sjellja " & Chr(10) & "ose nuk ka qene nxenes ne kete shkolle ose ne ciklin e percaktuar me lart.", vbInformation, "Konsultimi i amzes."
         Else

            active_form.rtbShenime.Text = tmp.Fields("SJ_SJELLJE")
            
            ' Ndryshohet ngjyra e tekstit qe sherben per deftesen
            active_form.rtbShenime.SelStart = 0
            If active_form.optUlet Then
               active_form.rtbShenime.SelLength = 206
            Else
               active_form.rtbShenime.SelLength = 373
            End If
            active_form.rtbShenime.SelColor = vbBlue
            active_form.rtbShenime.SelLength = 0
         End If
         tmp.Close

      Case SJELLJE_HIDH

         Klasa = active_form.cboKlasa.Text
         nrVitiShkollor = active_form.cboVitiShkollor.Text
         indeksi = active_form.cboIndeksi.Text
         cikli = ktheCiklin(Klasa, CStr(nrVitiShkollor))
         If cikli Then
            strCikli = "1"
         Else
            strCikli = "0"
         End If
         Dim sjellja          As String
         sjellja = active_form.rtbSjellja.Text
         j = Len(sjellja)
         Dim sj               As String
         Dim simboli          As String
         a = 1
         Do While a <= j
            simboli = Mid(sjellja, a, 1)
            If Asc(simboli) <> 254 Then
               sj = sj & Mid(sjellja, a, 1)
            End If
            a = a + 1
         Loop
         i = active_form.lista.ListIndex
         If i <> -1 Then
            numerAmze = active_form.listAmza.List(i)
            q_res1 = "SELECT SJ_SJELLJE FROM TBL_SJELLJE WHERE SJ_NR_AMZA = '" & numerAmze & "' AND SJ_CIKLI = " & strCikli & " AND SJ_VITI_SHKOLLOR = '" & nrVitiShkollor & "'"
            tmp.Open q_res1, conn, adOpenDynamic, adLockOptimistic
            If tmp.EOF Then
               q_res = "INSERT INTO TBL_SJELLJE (SJ_NR_AMZA, SJ_CIKLI, SJ_VITI_SHKOLLOR, SJ_SJELLJE) VALUES ('" & numerAmze & "', " & strCikli & ", '" & nrVitiShkollor & "', '" & sj & "')"
               conn.Execute q_res
               MsgBox "Shenimet mbi sjelljen per nxenesin e zgjedhur nga ju u hodhen.", vbInformation, "Shenime mbi sjelljen."
            Else
               MsgBox "Nxenesit qe ju keni zgjedhur i eshte hedhur sjellja per vitin shkollor te dhene.", vbInformation, "Shenime sjelljen."
            End If
         Else
            MsgBox "Ju duhet te selektoni nje nxenes para se te hidhni sjelljen e tij.", vbInformation, "Shenime nbi sjelljen."
         End If
    Case SHENIME_PERKOHSHME_HIDH

         Klasa = active_form.cboKlasa.Text
         nrVitiShkollor = active_form.txtVitiShkollor.Text
         indeksi = active_form.cboIndeksi.Text
         cikli = ktheCiklin(Klasa, CStr(nrVitiShkollor))
         If cikli Then
            strCikli = "1"
         Else
            strCikli = "0"
         End If
         Dim sjellja1          As String
         Dim sjellja2          As String
         Dim sjellja3          As String
         sjellja1 = active_form.rtbSjellja1.Text
         sjellja2 = active_form.rtbSjellja2.Text
         sjellja3 = active_form.rtbSjellja3.Text
         'Shenimi i pare
         j = Len(sjellja1)
         Dim simbol          As String
         Dim sj1               As String
           a = 1
         Do While a <= j
            simbol = Mid(sjellja1, a, 1)
            If Asc(simbol) <> 254 Then
               sj1 = sj1 & Mid(sjellja1, a, 1)
            End If
            a = a + 1
         Loop
         'Shenimi i dyte
         j = Len(sjellja2)
         Dim sj2              As String
         a = 1
         Do While a <= j
            simbol = Mid(sjellja2, a, 1)
            If Asc(simbol) <> 254 Then
               sj2 = sj2 & Mid(sjellja2, a, 1)
            End If
            a = a + 1
         Loop
         'Shenimi i trete
         j = Len(sjellja3)
         Dim sj3              As String
         a = 1
         Do While a <= j
            simbol = Mid(sjellja3, a, 1)
            If Asc(simbol) <> 254 Then
               sj3 = sj3 & Mid(sjellja3, a, 1)
            End If
            a = a + 1
         Loop
         Dim strDone As String
         i = active_form.lista.ListIndex
         If i <> -1 Then
           numerAmze = active_form.listAmza.List(i)
           If Trim(sj1) <> "" Then
                'Kontrollohet nqs nxenesit i eshte hedhur shenimi i pare/mes semestri i pare
                q_res1 = "SELECT SH_SHENIM FROM TBL_SHENIME_PERKOHSHME WHERE SH_NR_AMZA = '" & numerAmze & "' AND SH_CIKLI = " & strCikli & " AND SH_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND SH_LLOJI = 1"
                tmp.Open q_res1, conn, adOpenDynamic, adLockOptimistic
                If tmp.EOF Then
                   q_res = "INSERT INTO TBL_SHENIME_PERKOHSHME (SH_NR_AMZA, SH_CIKLI, SH_VITI_SHKOLLOR, SH_SHENIM, SH_LLOJI) VALUES ('" & numerAmze & "', " & strCikli & ", '" & nrVitiShkollor & "', '" & sj1 & "', 1)"
                   conn.Execute q_res
                   strDone = strDone + "1"
                Else
                   MsgBox "Nxënësit që ju keni zgjedhur i është hedhur shënimi për mesin e semestrit të parë për vitin shkollor korrent.", vbInformation, "Shënime të përkohshne."
                End If
                tmp.Close
            End If
            
            If Trim(sj2) <> "" Then
                'Kontrollohet nqs nxenesit i eshte hedhur shenimi i dyte/fund semestri i pare
                q_res1 = "SELECT SH_SHENIM FROM TBL_SHENIME_PERKOHSHME WHERE SH_NR_AMZA = '" & numerAmze & "' AND SH_CIKLI = " & strCikli & " AND SH_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND SH_LLOJI = 2"
                tmp.Open q_res1, conn, adOpenDynamic, adLockOptimistic
                If tmp.EOF Then
                   q_res = "INSERT INTO TBL_SHENIME_PERKOHSHME (SH_NR_AMZA, SH_CIKLI, SH_VITI_SHKOLLOR, SH_SHENIM, SH_LLOJI) VALUES ('" & numerAmze & "', " & strCikli & ", '" & nrVitiShkollor & "', '" & sj2 & "', 2)"
                   conn.Execute q_res
                   strDone = strDone + "2"
                Else
                   MsgBox "Nxënësit që ju keni zgjedhur i është hedhur shënimi për fundin e semestrit të parë për vitin shkollor korrent.", vbInformation, "Shënime të përkohshne."
                End If
                tmp.Close
                
            End If
            
            If Trim(sj3) <> "" Then
                'Kontrollohet nqs nxenesit i eshte hedhur shenimi i trete/mes semestri i dyte
                q_res1 = "SELECT SH_SHENIM FROM TBL_SHENIME_PERKOHSHME WHERE SH_NR_AMZA = '" & numerAmze & "' AND SH_CIKLI = " & strCikli & " AND SH_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND SH_LLOJI = 3"
                tmp.Open q_res1, conn, adOpenDynamic, adLockOptimistic
                If tmp.EOF Then
                   q_res = "INSERT INTO TBL_SHENIME_PERKOHSHME (SH_NR_AMZA, SH_CIKLI, SH_VITI_SHKOLLOR, SH_SHENIM, SH_LLOJI) VALUES ('" & numerAmze & "', " & strCikli & ", '" & nrVitiShkollor & "', '" & sj3 & "', 3)"
                   conn.Execute q_res
                   strDone = strDone + "3"
                Else
                   MsgBox "Nxënësit që ju keni zgjedhur i është hedhur shënimi për mesin e semestrit të dytë për vitin shkollor korrent.", vbInformation, "Shënime të përkohshne."
                End If
                tmp.Close
            End If
            
            Dim nrDone As Integer
            nrDone = Len(strDone)
            If (nrDone >= 1) Then
                Dim k As Integer
                k = 1
                Dim mesazh As String
                mesazh = "Për nxënësin e zgjedhur u hodhën shënimet e përkohshme të mëposhtëm:"
                Do While (k <= nrDone)
                    Dim ch As String
                    ch = Mid(strDone, k, 1)
                    Select Case ch
                        Case "1"
                            mesazh = mesazh + Chr(13) + Chr(15) + "Shënimi për mesin e semestrit të parë"
                        Case "2"
                            mesazh = mesazh + Chr(13) + Chr(15) + "Shënimi për fundin e semestrit të parë"
                        Case "3"
                            mesazh = mesazh + Chr(13) + Chr(15) + "Shënimi për mesin e semestrit të dytë"
                    End Select
                    k = k + 1
                Loop
               MsgBox mesazh, vbInformation, "Shënime të përkohshme."
            End If
         Else
            MsgBox "Ju duhet të selektoni një nxënës para se të hidhni shënime të përkohshme për të.", vbInformation, "Shënime të përkohshme."
         End If
      Case GJEJ_SHENIME_PERKOHSHME_PER_NXENESIN
         Klasa = active_form.cboKlasa.Text
         nrVitiShkollor = active_form.cboVitiShkollor.Text
         indeksi = active_form.cboIndeksi.Text
         cikli = ktheCiklin(Klasa, CStr(nrVitiShkollor))
         If cikli Then
            strCikli = "1"
         Else
            strCikli = "0"
         End If
         i = active_form.lista.ListIndex
         If i <> -1 Then
           numerAmze = active_form.listAmza.List(i)
            'Gjenden shenimet e hedhura per nxenesin
            q_res1 = "SELECT SH_SHENIM, SH_LLOJI FROM TBL_SHENIME_PERKOHSHME WHERE SH_NR_AMZA = '" & numerAmze & "' AND SH_CIKLI = " & strCikli & " AND SH_VITI_SHKOLLOR = '" & nrVitiShkollor & "'"
            tmp.Open q_res1, conn, adOpenDynamic, adLockOptimistic
            Dim shenim1, shenim2, shenim3 As String
            Do While Not tmp.EOF
               If tmp.Fields("SH_LLOJI") = 1 Then
                    shenim1 = tmp.Fields("SH_SHENIM")
                Else
                If tmp.Fields("SH_LLOJI") = 2 Then
                    shenim2 = tmp.Fields("SH_SHENIM")
                Else
                If tmp.Fields("SH_LLOJI") = 3 Then
                    shenim3 = tmp.Fields("SH_SHENIM")
                End If
                End If
                End If
                tmp.MoveNext
            Loop
            tmp.Close
         End If
         active_form.rtbSjellja1.Text = shenim1
         active_form.rtbSjellja2.Text = shenim2
         active_form.rtbSjellja3.Text = shenim3
      Case MODIFIKO_SHENIME_PERKOHSHME_PER_NXENESIN
         Klasa = active_form.cboKlasa.Text
         nrVitiShkollor = active_form.cboVitiShkollor.Text
         indeksi = active_form.cboIndeksi.Text
         cikli = ktheCiklin(Klasa, CStr(nrVitiShkollor))
         If cikli Then
            strCikli = "1"
         Else
            strCikli = "0"
         End If
         'Dim sjellja1          As String
         'Dim sjellja2          As String
         'Dim sjellja3          As String
         sjellja1 = active_form.rtbSjellja1.Text
         sjellja2 = active_form.rtbSjellja2.Text
         sjellja3 = active_form.rtbSjellja3.Text
         'Shenimi i pare
         j = Len(sjellja1)
         'Dim simbol          As String
         'Dim sj1               As String
           a = 1
         Do While a <= j
            simbol = Mid(sjellja1, a, 1)
            If Asc(simbol) <> 254 Then
               sj1 = sj1 & Mid(sjellja1, a, 1)
            End If
            a = a + 1
         Loop
         'Shenimi i dyte
         j = Len(sjellja2)
         'Dim sj2              As String
         a = 1
         Do While a <= j
            simbol = Mid(sjellja2, a, 1)
            If Asc(simbol) <> 254 Then
               sj2 = sj2 & Mid(sjellja2, a, 1)
            End If
            a = a + 1
         Loop
         'Shenimi i trete
         j = Len(sjellja3)
         'Dim sj3              As String
         a = 1
         Do While a <= j
            simbol = Mid(sjellja3, a, 1)
            If Asc(simbol) <> 254 Then
               sj3 = sj3 & Mid(sjellja3, a, 1)
            End If
            a = a + 1
         Loop
         'Dim strDone As String
         i = active_form.lista.ListIndex
         If i <> -1 Then
           numerAmze = active_form.listAmza.List(i)
            'Kontrollohet nqs nxenesit i eshte hedhur shenimi i pare/mes semestri i pare
            q_res1 = "SELECT SH_SHENIM FROM TBL_SHENIME_PERKOHSHME WHERE SH_NR_AMZA = '" & numerAmze & "' AND SH_CIKLI = " & strCikli & " AND SH_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND SH_LLOJI = 1"
            tmp.Open q_res1, conn, adOpenDynamic, adLockOptimistic
            If tmp.EOF And Trim(sj1) = "" Then
              
                Else
                If Not tmp.EOF And Trim(sj1) = "" Then
                   MsgBox "Shënimi për mesin e semestrit të parë nuk duhet të jetë bosh në mënyre që të modifikohet.", vbInformation, "Modifikimi i shënimeve të përkohshne."
                   Else
                   If tmp.EOF And Trim(sj1) <> "" Then
                        MsgBox "Për nxënësin e zgjedhur nuk është hedhur ende shënimi për mesin e semestrit të parë!", vbInformation, "Modifikimi i shënimeve të përkohshne."
                        Else
                        If Not tmp.EOF And Trim(sj1) <> "" Then
                            q_res = "UPDATE TBL_SHENIME_PERKOHSHME SET SH_SHENIM = '" + sj1 + "' WHERE SH_NR_AMZA = '" + numerAmze + "' AND SH_CIKLI = " + strCikli + " AND SH_VITI_SHKOLLOR = '" + nrVitiShkollor + "' AND SH_LLOJI = 1"
                            conn.Execute q_res
                            strDone = strDone + "1"
                        End If
                    End If
                End If
            End If
            tmp.Close
            
            'Kontrollohet nqs nxenesit i eshte hedhur shenimi i dyte/fund semestri i pare
            q_res1 = "SELECT SH_SHENIM FROM TBL_SHENIME_PERKOHSHME WHERE SH_NR_AMZA = '" & numerAmze & "' AND SH_CIKLI = " & strCikli & " AND SH_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND SH_LLOJI = 2"
            tmp.Open q_res1, conn, adOpenDynamic, adLockOptimistic
            If tmp.EOF And Trim(sj2) = "" Then
              
                Else
                If Not tmp.EOF And Trim(sj2) = "" Then
                   MsgBox "Shënimi për fundin e semestrit të parë nuk duhet të jetë bosh në mënyre që të modifikohet.", vbInformation, "Modifikimi i shënimeve të përkohshne."
                   Else
                   If tmp.EOF And Trim(sj2) <> "" Then
                        MsgBox "Për nxënësin e zgjedhur nuk është hedhur ende shënimi për fundin e semestrit të parë!", vbInformation, "Modifikimi i shënimeve të përkohshne."
                        Else
                        If Not tmp.EOF And Trim(sj2) <> "" Then
                            q_res = "UPDATE TBL_SHENIME_PERKOHSHME SET SH_SHENIM = '" + sj2 + "' WHERE SH_NR_AMZA = '" + numerAmze + "' AND SH_CIKLI = " + strCikli + " AND SH_VITI_SHKOLLOR = '" + nrVitiShkollor + "' AND SH_LLOJI = 2"
                            conn.Execute q_res
                            strDone = strDone + "2"
                        End If
                    End If
                End If
            End If
            tmp.Close
           
           'Kontrollohet nqs nxenesit i eshte hedhur shenimi i trete/mes semestri i dyte
            q_res1 = "SELECT SH_SHENIM FROM TBL_SHENIME_PERKOHSHME WHERE SH_NR_AMZA = '" & numerAmze & "' AND SH_CIKLI = " & strCikli & " AND SH_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND SH_LLOJI = 3"
            tmp.Open q_res1, conn, adOpenDynamic, adLockOptimistic
            If tmp.EOF And Trim(sj3) = "" Then
              
                Else
                If Not tmp.EOF And Trim(sj3) = "" Then
                   MsgBox "Shënimi për mesin e semestrit të dytë nuk duhet të jetë bosh në mënyre që të modifikohet.", vbInformation, "Modifikimi i shënimeve të përkohshne."
                   Else
                   If tmp.EOF And Trim(sj3) <> "" Then
                        MsgBox "Për nxënësin e zgjedhur nuk është hedhur ende shënimi për mesin e semestrit të dytë!", vbInformation, "Modifikimi i shënimeve të përkohshne."
                        Else
                        If Not tmp.EOF And Trim(sj3) <> "" Then
                            q_res = "UPDATE TBL_SHENIME_PERKOHSHME SET SH_SHENIM = '" + sj3 + "' WHERE SH_NR_AMZA = '" + numerAmze + "' AND SH_CIKLI = " + strCikli + " AND SH_VITI_SHKOLLOR = '" + nrVitiShkollor + "' AND SH_LLOJI = 3"
                            conn.Execute q_res
                            strDone = strDone + "3"
                        End If
                    End If
                End If
            End If
            tmp.Close
            
            'Dim nrDone As Integer
            nrDone = Len(strDone)
            If (nrDone >= 1) Then
                'Dim k As Integer
                k = 1
                'Dim mesazh As String
                mesazh = "Për nxënësin e zgjedhur u modifikuan shënimet e përkohshme të mëposhtëm:"
                Do While (k <= nrDone)
                    'Dim ch As String
                    ch = Mid(strDone, k, 1)
                    Select Case ch
                        Case "1"
                            mesazh = mesazh + Chr(13) + Chr(15) + "Shënimi për mesin e semestrit të parë"
                        Case "2"
                            mesazh = mesazh + Chr(13) + Chr(15) + "Shënimi për fundin e semestrit të parë"
                        Case "3"
                            mesazh = mesazh + Chr(13) + Chr(15) + "Shënimi për mesin e semestrit të dytë"
                    End Select
                    k = k + 1
                Loop
               MsgBox mesazh, vbInformation, "Modifikimi i shënimeve të përkohshme."
            End If
         Else
            MsgBox "Ju duhet të selektoni një nxënës para se të modifikoni shënimet të përkohshme për të.", vbInformation, "Modifikimi i shënimeve të përkohshme."
         End If
         
      Case FSHI_TE_DHENAT

         Dim objKopjo         As New FileSystemObject
         Dim kopja            As String
         kopja = "kopja" & " " & Format(date, "dd.mm.yy")
         objKopjo.CopyFile App.Path & "\ActiveX\" & "DProgNec.dat", App.Path & "\Back up\" & kopja, True
         q_res = "DELETE * FROM TBL_AMZA"
         conn.Execute q_res
         q_res = "DELETE * FROM TBL_NOTA"
         conn.Execute q_res
         q_res = "DELETE * FROM TBL_KLASA"
         conn.Execute q_res
         q_res = "DELETE * FROM TBL_MODIFIKIMI"
         conn.Execute q_res
         q_res = "DELETE * FROM TBL_LENDA"
         conn.Execute q_res
         q_res = "DELETE * FROM TBL_NOTAAMZ"
         conn.Execute q_res
         q_res = "DELETE * FROM TBL_PERDORUES"
         conn.Execute q_res
         q_res = "DELETE * FROM TBL_SJELLJE"
         conn.Execute q_res
         q_res = "DELETE * FROM TBL_USERS"
         conn.Execute q_res
         MsgBox "Te dhenat e ruajtura deri tani u fshine plotesisht ne menyre te parekuperueshme.", vbInformation, "Fshirja e te dhenave!"

         Set objKopjo = Nothing

      Case FSHI_PROVIMET

         If active_form.optMesme.Value Then
            cikli = 1
            strCikli = "1"
         Else
            cikli = 0
            strCikli = "0"
         End If
         Dim gjatesia         As Integer
         nrVitiShkollor = active_form.cboVitiShkollor.Text

         gjatesia = active_form.listaProvimet.ListCount - 1
         i = 0
         Do While i <= gjatesia
            If active_form.listaProvimet.Selected(i) Then
               lenda = active_form.listaProvimet.List(i)
               q_res = "DELETE  FROM TBL_NOTA WHERE NT_CIKLI = " & strCikli & " AND NT_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND NT_EMRI_LENDA = '" & lenda & "' AND NT_PROVIM = 1 "
               conn.Execute q_res
               q_res = "DELETE FROM TBL_NOTAAMZ WHERE NTA_CIKLI = " & strCikli & " AND NTA_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND NTA_EMRI_PROVIMI = '" & lenda & "'"
               conn.Execute q_res
               active_form.listaProvimet.RemoveItem (i)
               i = i - 1
               gjatesia = gjatesia - 1
            End If
            i = i + 1
         Loop

      Case SHENIM_PERKOHSHEM_KERKO

         Klasa = active_form.cboKlasa.Text
         nrVitiShkollor = active_form.txtVitiShkollor.Text
         indeksi = active_form.cboIndeksi.Text

         q_res = "SELECT AMZA_EMRI , AMZA_MBIEMRI, AMZA_NR_AMZA, KL_KLASA, KL_INDEKSI, KL_VITI_SHKOLLOR FROM TBL_AMZA, TBL_KLASA  WHERE (AMZA_NR_AMZA = KL_NR_AMZA AND AMZA_CIKLI = KL_CIKLI) AND (KL_KLASA = '" & Klasa
         q_res1 = "') AND (KL_INDEKSI = '" & indeksi & "') AND (KL_VITI_SHKOLLOR = '" & nrVitiShkollor & "') AND (AMZA_LARGUAR = 0) ORDER BY AMZA_EMRI, AMZA_MBIEMRI "
         tmp.Open q_res + q_res1, conn, adOpenDynamic, adLockOptimistic
         active_form.listAmza.Clear
         If Not tmp.EOF Then
            Do While Not tmp.EOF
               emri = tmp.Fields("AMZA_EMRI")
               mbiemri = tmp.Fields("AMZA_MBIEMRI")
               numerAmze = tmp.Fields("AMZA_NR_AMZA")
               active_form.lista.AddItem emri & Space(4) & mbiemri
               active_form.listAmza.AddItem numerAmze
               tmp.MoveNext
            Loop
            active_form.cmdOK.Enabled = True
         Else
            MsgBox "Klasa e zgjedhur nga ju nuk ekziston .", vbInformation, "Shenimet."
            active_form.cmdOK.Enabled = False
            active_form.cboKlasa.ListIndex = -1
            active_form.cboIndeksi.ListIndex = -1
            active_form.rtbSjellja1.Text = ""
            active_form.rtbSjellja2.Text = ""
            active_form.rtbSjellja2.Text = ""
         End If
         tmp.Close
    Case SJELLJE_KERKO
         Klasa = active_form.cboKlasa.Text
         nrVitiShkollor = active_form.cboVitiShkollor.Text
         indeksi = active_form.cboIndeksi.Text

         q_res = "SELECT AMZA_EMRI , AMZA_MBIEMRI, AMZA_NR_AMZA, KL_KLASA, KL_INDEKSI, KL_VITI_SHKOLLOR FROM TBL_AMZA, TBL_KLASA  WHERE (AMZA_NR_AMZA = KL_NR_AMZA AND AMZA_CIKLI = KL_CIKLI) AND (KL_KLASA = '" & Klasa
         q_res1 = "') AND (KL_INDEKSI = '" & indeksi & "') AND (KL_VITI_SHKOLLOR = '" & nrVitiShkollor & "') AND (AMZA_LARGUAR = 0) ORDER BY AMZA_EMRI, AMZA_MBIEMRI "
         tmp.Open q_res + q_res1, conn, adOpenDynamic, adLockOptimistic
         active_form.listAmza.Clear
         If Not tmp.EOF Then
            Do While Not tmp.EOF
               emri = tmp.Fields("AMZA_EMRI")
               mbiemri = tmp.Fields("AMZA_MBIEMRI")
               numerAmze = tmp.Fields("AMZA_NR_AMZA")
               active_form.lista.AddItem emri & Space(4) & mbiemri
               active_form.listAmza.AddItem numerAmze
               tmp.MoveNext
            Loop
            active_form.cmdOK.Enabled = True
         Else
            MsgBox "Klasa e zgjedhur nga ju nuk ekziston .", vbInformation, "Shenimet."
            active_form.cmdOK.Enabled = False
            active_form.cboKlasa.ListIndex = -1
            active_form.cboIndeksi.ListIndex = -1
            active_form.rtbSjellja.Text = ""
            'active_form.rtbSjellja2.Text = ""
            'active_form.rtbSjellja3.Text = ""
         End If
         tmp.Close
     Case NUMERO_NXENES
        q_res = "SELECT COUNT(AMZA_NR_AMZA) FROM TBL_AMZA"
        tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
        numerNxenesish = tmp.Fields(0)
        tmp.Close
        
     Case VERIFIKO_NOTAT

         'Dim nota             As String
         '*** hedh notat e nje nxenesi ne gride
         nrVitiShkollor = active_form.cboVitiShkollor.Text
         Dim ka_nota          As Boolean
         ka_nota = True
         '*** ketu percaktohet cikli i nxenesit

         If active_form.optMesme.Value Then
            cikli = 1
            strCikli = "1"
         Else
            cikli = 0
            strCikli = "0"
         End If

         '***
         '*** ketu kryhet hedhja e te dhenave plotesuese te nxenesit
         If active_form.txtAmzaNo.Text <> "" Then
            numerAmze = active_form.txtAmzaNo.Text
            
            q_res = " SELECT AMZA_NR_AMZA, AMZA_EMRI, AMZA_MBIEMRI, AMZA_ATESIA, AMZA_DATELINDJA, AMZA_VREJTJE, AMZA_VENDLINDJA,KL_KLASA,KL_INDEKSI, KL_VITI_SHKOLLOR FROM TBL_AMZA, TBL_KLASA WHERE ( AMZA_NR_AMZA = KL_NR_AMZA AND AMZA_CIKLI = KL_CIKLI) AND (AMZA_NR_AMZA = '" & numerAmze & "' AND AMZA_CIKLI = " & strCikli & ") AND KL_VITI_SHKOLLOR = '" & nrVitiShkollor & "' ORDER BY AMZA_EMRI, AMZA_MBIEMRI"
            tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic

            q_res1 = "SELECT SJ_SJELLJE FROM TBL_SJELLJE WHERE (SJ_NR_AMZA = '" & _
               numerAmze & "') AND (SJ_CIKLI = " & cikli & ") AND (SJ_VITI_SHKOLLOR = '" & nrVitiShkollor & "')"
            tmp11.Open q_res1, conn, adOpenDynamic, adLockOptimistic

            'get_rec (q_res)
            If Not tmp.EOF Then
               active_form.txtEmri.Text = tmp.Fields("AMZA_EMRI")
               active_form.txtMbiemri.Text = tmp.Fields("AMZA_MBIEMRI")
               active_form.txtAtesia = tmp.Fields("AMZA_ATESIA")
               active_form.txtKlasa.Text = tmp.Fields("KL_KLASA")
               active_form.txtIndeksi.Text = tmp.Fields("KL_INDEKSI")
               active_form.Vendlindja = tmp.Fields("AMZA_VENDLINDJA")
               active_form.Datelindja = tmp.Fields("AMZA_DATELINDJA")
               tmp.Close
               If Not tmp11.EOF Then
                  active_form.Vrejtje = tmp11.Fields("SJ_SJELLJE")
               End If
               tmp11.Close
Else:
               tmp.Close
               tmp11.Close
               active_form.txtEmri.Text = ""
               active_form.txtMbiemri.Text = ""
               active_form.txtAtesia.Text = ""
               active_form.txtKlasa.Text = ""
               active_form.txtIndeksi.Text = ""
               active_form.lblLargimi.Text = ""
               ' Atesia, Vendlindja, Datelindja, Vrejtje,  sherbejne vetem per
               ' printimin e defteses

               active_form.Label3.Visible = False
               active_form.txtData.Visible = False
               active_form.Cgrid1.Visible = False
               active_form.Lendet.Visible = False
               active_form.Label2.Visible = False
               active_form.lblNotatDheMungesat.Visible = False
               active_form.lblS1.Visible = False
               active_form.lblS2.Visible = False
               active_form.lblV.Visible = False
               active_form.lblR.Visible = False
               active_form.fsbHorizontali.Visible = False
               active_form.fsbVertical.Visible = False
               scrbar = False
               scrbarver = False
               active_form.cmdPrint.Enabled = False
               MsgBox "   Nxenesi me numrin e amzes se dhene nuk eshte i regjistruar ne vitin shkollor qe eshte dhene" & vbCrLf & "   dhe ne ciklin qe eshte zgjedhur", vbInformation, "  Konsultime Nota"
               active_form.txtAmzaNo.Text = ""
               Exit Sub
            End If
         Else
            emri = active_form.txtEmri.Text
            mbiemri = active_form.txtMbiemri.Text
            atesia = active_form.txtAtesia.Text
            q_res = " SELECT AMZA_NR_AMZA,AMZA_CIKLI, AMZA_EMRI, AMZA_MBIEMRI, AMZA_ATESIA, AMZA_DATELINDJA,AMZA_VREJTJE, AMZA_VENDLINDJA,KL_KLASA,KL_INDEKSI, KL_VITI_SHKOLLOR FROM TBL_AMZA, TBL_KLASA WHERE (AMZA_NR_AMZA = KL_NR_AMZA AND AMZA_CIKLI = KL_CIKLI ) AND"
            Dim st As String
            st = " ( AMZA_EMRI LIKE '" & emri & "%' AND AMZA_MBIEMRI LIKE '" & mbiemri & "%' AND AMZA_ATESIA LIKE '" & atesia & "%' AND AMZA_CIKLI = " & strCikli & ")  AND KL_VITI_SHKOLLOR = '" & nrVitiShkollor & "' ORDER BY AMZA_EMRI, AMZA_ATESIA, AMZA_MBIEMRI"
            q_res = q_res + st
            tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
            If Not tmp.EOF Then
                If tmp.RecordCount = 1 Then
                    active_form.txtAmzaNo.Text = tmp.Fields("AMZA_NR_AMZA")
                    active_form.txtKlasa.Text = tmp.Fields("KL_KLASA")
                    active_form.txtIndeksi.Text = tmp.Fields("KL_INDEKSI")
                    active_form.txtEmri.Text = tmp.Fields("AMZA_EMRI")
                    active_form.txtMbiemri.Text = tmp.Fields("AMZA_MBIEMRI")
                    active_form.txtAtesia = tmp.Fields("AMZA_ATESIA")
                    tmp.Close
                Else
                    'nqs ka me teper se nje nxenes me keto te dhena atehere shfaq formen e kerkimit
                    'inicializo gridat ne forme
                    nrRreshta = tmp.RecordCount
                    Call inicializoGridaZgjidhNxenes(nrRreshta)
                    'mbush griden me te dhena
                    Index = 0
                    Do While Not tmp.EOF
                        frmZgjidhNxenes.gridaNxenesit.row = Index
                        frmZgjidhNxenes.gridaNxenesit.col = 0
                        frmZgjidhNxenes.gridaNxenesit.Text = tmp.Fields("AMZA_NR_AMZA").Value
                        frmZgjidhNxenes.gridaNxenesit.col = 1
                        frmZgjidhNxenes.gridaNxenesit.Text = tmp.Fields("AMZA_EMRI").Value + " " + tmp.Fields("AMZA_ATESIA").Value + " " + tmp.Fields("AMZA_MBIEMRI").Value
                        frmZgjidhNxenes.gridaNxenesit.col = 2
                        frmZgjidhNxenes.gridaNxenesit.Text = tmp.Fields("AMZA_CIKLI").Value
                        tmp.MoveNext
                        Index = Index + 1
                    Loop
                    tmp.Close
                    ShkollaManager.frmZgjidhNxenes.show vbModal
                    numerAmze = ShkollaManager.frmZgjidhNxenes.Amza
                    strCikli = ShkollaManager.frmZgjidhNxenes.cikli
                    If (numerAmze = "" Or strCikli = "") Then
                    'pastro
                        active_form.txtEmri.Text = ""
                        active_form.txtMbiemri.Text = ""
                        active_form.txtAtesia.Text = ""
                        active_form.txtKlasa.Text = ""
                        active_form.txtIndeksi.Text = ""
                        active_form.lblLargimi.Text = ""
                        ' Atesia, Vendlindja, Datelindja, Vrejtje,  sherbejne vetem per
                        ' printimin e defteses
                        active_form.Label3.Visible = False
                        active_form.txtData.Visible = False
                        active_form.Cgrid1.Visible = False
                        active_form.Lendet.Visible = False
                        active_form.Label2.Visible = False
                        active_form.lblNotatDheMungesat.Visible = False
                        active_form.lblS1.Visible = False
                        active_form.lblS2.Visible = False
                        active_form.lblV.Visible = False
                        active_form.lblR.Visible = False
                        active_form.fsbHorizontali.Visible = False
                        active_form.fsbVertical.Visible = False
                        scrbar = False
                        scrbarver = False
                        active_form.cmdPrint.Enabled = False
                        Exit Sub
                    End If
                    q_res = " SELECT AMZA_NR_AMZA, AMZA_EMRI, AMZA_MBIEMRI, AMZA_ATESIA, AMZA_DATELINDJA, AMZA_VREJTJE, AMZA_VENDLINDJA,KL_KLASA,KL_INDEKSI, KL_VITI_SHKOLLOR FROM TBL_AMZA, TBL_KLASA WHERE ( AMZA_NR_AMZA = KL_NR_AMZA AND AMZA_CIKLI = KL_CIKLI) AND (AMZA_NR_AMZA = '" & numerAmze & "' AND AMZA_CIKLI = " & strCikli & ") AND KL_VITI_SHKOLLOR = '" & nrVitiShkollor & "' ORDER BY AMZA_EMRI, AMZA_MBIEMRI"
                    tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
                    If Not tmp.EOF Then
                        active_form.txtEmri.Text = tmp.Fields("AMZA_EMRI")
                        active_form.txtMbiemri.Text = tmp.Fields("AMZA_MBIEMRI")
                        active_form.txtAtesia = tmp.Fields("AMZA_ATESIA")
                        active_form.txtKlasa.Text = tmp.Fields("KL_KLASA")
                        active_form.txtIndeksi.Text = tmp.Fields("KL_INDEKSI")
                        active_form.txtAmzaNo.Text = tmp.Fields("AMZA_NR_AMZA")
                        tmp.Close
                    End If
                End If
Else:
               tmp.Close
               active_form.txtEmri.Text = ""
               active_form.txtMbiemri.Text = ""
               active_form.txtAtesia.Text = ""
               active_form.lblLargimi.Text = ""
               active_form.Label3.Visible = False
               active_form.txtData.Visible = False
               active_form.Cgrid1.Visible = False
               active_form.Lendet.Visible = False
               active_form.Label2.Visible = False
               active_form.fsbHorizontali.Visible = False
               active_form.fsbVertical.Visible = False
               active_form.lblNotatDheMungesat.Visible = False
               active_form.lblS1.Visible = False
               active_form.lblS2.Visible = False
               active_form.lblV.Visible = False
               active_form.lblR.Visible = False
               MsgBox "   Nxenesi me kete emer dhe mbiemer  nuk eshte i regjistruar ne vitin shkollor qe eshte dhene" & vbCrLf & "   dhe ne ciklin qe eshte zgjedhur", vbInformation, "  Konsultime Nota"
               Exit Sub
            End If
         End If
         numerAmze = active_form.txtAmzaNo.Text

         '***** Ketu percaktohen numri i lendeve
         'Dim nr_lendesh       As Integer
         Klasa = active_form.txtKlasa.Text
         q_res = "SELECT COUNT(LN_EMRI) AS NR_LENDESH FROM TBL_LENDA WHERE LN_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND LN_KLASA" & Klasa & " = 1 "
         tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
         nr_lendesh = tmp.Fields("NR_LENDESH")
         tmp.Close
         If nr_lendesh = 0 Then
            MsgBox "Klases se frekuentuar  nga nxenesi ne kete vit shkollor " & Chr(10) & "nuk i eshte regjistruar asnje lende.", vbExclamation, "Konsultimi i nxenesit."
            active_form.Label3.Visible = False
            active_form.txtData.Visible = False

            Exit Sub
         End If

         '**** ketu gjendet numri maksimal i notave ne nje lende

         'Dim nr_max_notash    As Integer
         If active_form.optSemestri1.Value Then
            q_res = "SELECT NT_VLERESIMI FROM TBL_NOTA WHERE NT_NR_AMZA = '" & numerAmze & "' AND NT_CIKLI = " & strCikli & " AND NT_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND (NT_MOMENTALES1 = 1 OR NT_SEMESTRI1 = 1) "
            tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
            If tmp.EOF Then
               nr_max_notash = 0
               active_form.txtData.Visible = False
               active_form.Label2.Visible = False
               active_form.Label3.Visible = False
               ' MsgBox "Nxenesi nuk ka marre asnje note", vbOKOnly + vbInformation, "Konsultimi i nxenesit"
               ka_nota = False
               tmp.Close
               'Exit Sub
            Else
               tmp.Close
               q_res = "SELECT MAX(A.NR_NOTASH) AS MAKSI FROM (SELECT  COUNT(NT_VLERESIMI) AS NR_NOTASH FROM TBL_NOTA WHERE NT_NR_AMZA = '" & numerAmze & "' AND NT_CIKLI = " & strCikli & " AND NT_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND (NT_MOMENTALES1 = 1 OR NT_SEMESTRI1 = 1) GROUP BY NT_EMRI_LENDA) AS A"
               tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
               If tmp.EOF Then
                  nr_max_notash = 0

               Else
                  nr_max_notash = tmp.Fields("MAKSI")
                  ' nr_max_notash = nr_max_notash - 1
               End If
               tmp.Close
            End If

         ElseIf active_form.optSemestri2.Value Then
            q_res = "SELECT NT_VLERESIMI FROM TBL_NOTA WHERE NT_NR_AMZA = '" & numerAmze & "' AND NT_CIKLI = " & strCikli & " AND NT_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND (NT_MOMENTALES2 = 1 OR NT_SEMESTRI2 = 1) "
            tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
            If tmp.EOF Then
               nr_max_notash = 0
               ka_nota = False
               tmp.Close
               '  active_form.Label2.Visible = False
               ' active_form.Label3.Visible = False
               ' active_form.txtData.Visible = False
               ' active_form.cmdPrint.Enabled = False
               'MsgBox "Nxenesi nuk ka marre akoma asnje note", vbOKOnly + vbInformation, "Konsultimi i nxenesit"
               ' Exit Sub
            Else
               tmp.Close
               q_res = "SELECT MAX(a.NR_NOTASH) AS MAKSI FROM (SELECT  COUNT(NT_VLERESIMI) AS NR_NOTASH FROM TBL_NOTA WHERE NT_NR_AMZA = '" & numerAmze & "' AND NT_CIKLI = " & strCikli & " AND NT_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND (NT_MOMENTALES2 = 1 OR NT_SEMESTRI2 = 1) GROUP BY NT_EMRI_LENDA) as a"
               tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
               If tmp.EOF Then
                  nr_max_notash = 0

               Else
                  nr_max_notash = tmp.Fields("MAKSI")
                  'nr_max_notash = nr_max_notash - 1
               End If
               tmp.Close

            End If
         Else
            q_res = "SELECT NT_VLERESIMI FROM TBL_NOTA WHERE NT_NR_AMZA = '" & numerAmze & "' AND NT_CIKLI = " & strCikli & " AND NT_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND (NT_SEMESTRI1 = 1 OR NT_SEMESTRI2 = 1 OR NT_VJETORE = 1 OR  NT_RIPROVIM = 1) "
            tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
            If tmp.EOF Then
               active_form.txtData.Visible = False
               active_form.Label2.Visible = False
               active_form.Label3.Visible = False
               ' MsgBox "Nxenesi nuk ka marre asnje note", vbOKOnly + vbInformation, "Konsultimi i nxenesit"
               ka_nota = False
               tmp.Close
               ' Exit Sub
            Else
                tmp.Close
            End If
            nr_max_notash = 3
         End If

         Call shfaq_grile1(nr_max_notash + 1, nr_lendesh)

         '*** ketu hidhen emrat e lendeve ne gride
         If strCikli = "1" Then
            q_res = "SELECT LN_EMRI FROM TBL_LENDA WHERE LN_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND LN_KLASA" & Klasa & " = 1  ORDER BY LN_NUMRI_MESME "
         Else
            q_res = "SELECT LN_EMRI FROM TBL_LENDA WHERE LN_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND LN_KLASA" & Klasa & " = 1  ORDER BY LN_NUMRI_ULET "
         End If
         tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
         i = 0
         Do While Not tmp.EOF
            i = i + 1
            active_form.Lendet.Text(i, 1) = tmp.Fields("LN_EMRI")
            tmp.MoveNext
         Loop
         If Not ka_nota Then
            tmp.Close
            GoTo verifiko_fund
         End If

         If active_form.optSemestri1.Value Then
            q_res1 = "SELECT * FROM TBL_NOTA WHERE NT_NR_AMZA = '" & numerAmze & "' AND NT_CIKLI = " & strCikli & " AND NT_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND (NT_MOMENTALES1 = 1 OR NT_SEMESTRI1 = 1)  ORDER BY  NT_MOMENTALES1 DESC, NT_SEMESTRI1 DESC, NT_DATA "
         ElseIf active_form.optSemestri2.Value Then
            q_res1 = "SELECT * FROM TBL_NOTA WHERE NT_NR_AMZA = '" & numerAmze & "' AND NT_CIKLI = " & strCikli & " AND NT_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND (NT_MOMENTALES2 = 1 OR NT_SEMESTRI2 = 1) ORDER BY   NT_MOMENTALES2 DESC, NT_SEMESTRI2 DESC, NT_DATA"
         Else
            q_res1 = "SELECT * FROM TBL_NOTA WHERE NT_NR_AMZA = '" & numerAmze & "' AND NT_CIKLI = " & strCikli & " AND NT_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND (NT_SEMESTRI1 = 1 OR NT_SEMESTRI2 = 1 OR NT_VJETORE = 1 OR NT_RIPROVIM = 1) ORDER BY  NT_SEMESTRI1 DESC , NT_SEMESTRI2 DESC , NT_VJETORE DESC , NT_RIPROVIM DESC"
         End If

         rekset.Open q_res1, conn, adOpenDynamic, adLockOptimistic
         tmp.MoveFirst
         i = 0
         Do While Not tmp.EOF
            i = i + 1
            lenda = tmp.Fields("LN_EMRI")
            rekset.MoveFirst
            j = 0
            Do While Not rekset.EOF
               If lenda = rekset("NT_EMRI_LENDA") Then
                  j = j + 1
                  If Not active_form.optPerfundimtare.Value Then
                     If rekset.Fields("NT_SEMESTRI1") Or rekset.Fields("NT_SEMESTRI2") Then
                        active_form.Cgrid1.Text(i, nr_max_notash + 1) = rekset.Fields("NT_VLERESIMI")
                     Else
                        active_form.Cgrid1.Text(i, j) = rekset.Fields("NT_VLERESIMI")
                     End If
                  Else
                     If rekset.Fields("NT_SEMESTRI1") Then
                        active_form.Cgrid1.Text(i, 1) = rekset.Fields("NT_VLERESIMI")
                     ElseIf rekset.Fields("NT_SEMESTRI2") Then

                        active_form.Cgrid1.Text(i, 2) = rekset.Fields("NT_VLERESIMI")
                     ElseIf rekset.Fields("NT_VJETORE") Then
                        active_form.Cgrid1.Text(i, 3) = rekset.Fields("NT_VLERESIMI")
                     Else
                        active_form.Cgrid1.Text(i, 4) = rekset.Fields("NT_VLERESIMI")
                     End If
                  End If
                  'active_form.Cgrid1.Text(I, j) = rekset.Fields("NT_VLERESIMI")
                  If (rekset.Fields("NT_MOMENTALES1") Or rekset.Fields("NT_MOMENTALES2")) And rekset.Fields("NT_DETYREKONTROLLI") = False Then
                     active_form.Cgrid1.CellForeColor(i, j) = &H80000008
                     active_form.Cgrid1.CellFontBold(i, j) = True
                     modifikoNotat(i, j) = 1
                     matricaNotat(i, j) = rekset.Fields("NT_TEKST_DATA")
                  End If
                  If active_form.optPerfundimtare.Value Then
                     If rekset.Fields("NT_SEMESTRI1") Then
                        active_form.Cgrid1.CellForeColor(i, 1) = &HC00000
                        active_form.Cgrid1.CellFontBold(i, 1) = True
                        modifikoNotat(i, 1) = 2
                        matricaNotat(i, 1) = rekset.Fields("NT_TEKST_DATA")
                     End If
                     If rekset.Fields("NT_SEMESTRI2") Then
                        active_form.Cgrid1.CellForeColor(i, 2) = &HC00000
                        active_form.Cgrid1.CellFontBold(i, 2) = True
                        modifikoNotat(i, 2) = 3
                        matricaNotat(i, 2) = rekset.Fields("NT_TEKST_DATA")
                     End If
                  Else
                     If rekset.Fields("NT_SEMESTRI1") Then
                        active_form.Cgrid1.CellForeColor(i, nr_max_notash + 1) = &HC00000
                        active_form.Cgrid1.CellFontBold(i, nr_max_notash + 1) = True
                        modifikoNotat(i, nr_max_notash + 1) = 2
                        matricaNotat(i, nr_max_notash + 1) = rekset.Fields("NT_TEKST_DATA")
                     End If
                     If rekset.Fields("NT_SEMESTRI2") Then
                        active_form.Cgrid1.CellForeColor(i, nr_max_notash + 1) = &HC00000
                        active_form.Cgrid1.CellFontBold(i, nr_max_notash + 1) = True
                        modifikoNotat(i, nr_max_notash + 1) = 3
                        matricaNotat(i, nr_max_notash + 1) = rekset.Fields("NT_TEKST_DATA")
                     End If
                  End If
                  If rekset.Fields("NT_VJETORE") Then
                     active_form.Cgrid1.CellForeColor(i, 3) = &HFF&
                     active_form.Cgrid1.CellFontBold(i, 3) = True
                     modifikoNotat(i, 3) = 4
                     matricaNotat(i, 3) = rekset.Fields("NT_TEKST_DATA")
                  End If
                  If rekset.Fields("NT_DETYREKONTROLLI") = True Then
                     active_form.Cgrid1.CellForeColor(i, j) = &H8000&
                     active_form.Cgrid1.CellFontBold(i, j) = True
                     modifikoNotat(i, j) = 1
                     matricaNotat(i, j) = rekset.Fields("NT_TEKST_DATA")
                  End If
                  If rekset.Fields("NT_MUNGESE_ME") = True Then
                     active_form.Cgrid1.CellForeColor(i, j) = &H80000008
                     active_form.Cgrid1.CellFontBold(i, j) = True
                     modifikoNotat(i, j) = 1
                     matricaNotat(i, j) = rekset.Fields("NT_TEKST_DATA")
                  End If
                  If rekset.Fields("NT_MUNGESE_PA") = True Then
                     active_form.Cgrid1.CellForeColor(i, j) = &HFF&
                     active_form.Cgrid1.CellFontBold(i, j) = True
                     modifikoNotat(i, j) = 1
                     matricaNotat(i, j) = rekset.Fields("NT_TEKST_DATA")
                  End If
                  If rekset.Fields("NT_RIPROVIM") = True Then
                     active_form.Cgrid1.CellForeColor(i, 4) = &HC000C0
                     active_form.Cgrid1.CellFontBold(i, 4) = True
                     modifikoNotat(i, j) = 10
                     matricaNotat(i, j) = rekset.Fields("NT_TEKST_DATA")
                  End If

               End If
               rekset.MoveNext
            Loop
            tmp.MoveNext
         Loop
         rekset.Close
         tmp.Close

verifiko_fund:

         q_res = "SELECT AMZA_LARGUAR FROM TBL_AMZA WHERE AMZA_NR_AMZA = '" & numerAmze & "' AND AMZA_CIKLI = " & strCikli
         tmp.Open q_res
         If Not tmp.EOF Then
            If tmp.Fields("AMZA_LARGUAR") = True Then
               active_form.lblLargimi.Text = "Nxenesi eshte larguar nga shkolla"
            End If
         End If
         tmp.Close

         For i = 1 To nr_max_notash + 1

            For j = 1 To nr_lendesh
               active_form.Cgrid1.CellEditMode(j, i) = False
               If Trim(active_form.Cgrid1.Text(j, i)) <> "" Then
                  If Len(Trim(active_form.Cgrid1.Text(j, i))) <> 2 Then
                     active_form.Cgrid1.Text(j, i) = Space(1) & active_form.Cgrid1.Text(j, i)

                  End If
               End If
            Next j
         Next i

         For j = 1 To nr_lendesh
            active_form.Lendet.CellEditMode(j, 1) = False
         Next j

         active_form.Label3.Visible = True
         active_form.txtData.Visible = True
         active_form.Lendet.Visible = True
         active_form.Cgrid1.Visible = True
         active_form.txtData.Visible = True
         active_form.Label3.Visible = True
         If active_form.optPerfundimtare.Value Then
            active_form.lblS1.Visible = True
            active_form.lblS2.Visible = True
            active_form.lblV.Visible = True
            active_form.lblR.Visible = True
            active_form.lblNotatDheMungesat.Visible = False
         Else
            active_form.lblS1.Visible = False
            active_form.lblS2.Visible = False
            active_form.lblV.Visible = False
            active_form.lblR.Visible = False
            active_form.lblNotatDheMungesat.Visible = True
         End If
         If nr_max_notash + 1 > 43 Then
            active_form.fsbHorizontali.Visible = True
         End If
         'If nr_lendesh > 18 Then
         '  active_form.fsbVertikali.Visible = fale
         ' End If
         'active_form.fsbVertikali.Visible = False
        
     Case MODIFIKO_AMZA1

      If active_form.optMesme.Value Then
         cikli = 1
         strCikli = "1"
      Else
         cikli = 0
         strCikli = "0"
      End If
      If active_form.txtAmzaNo.Text <> "" Then
         numerAmze = active_form.txtAmzaNo.Text
         q_res = "SELECT AMZA_EMRI , AMZA_MBIEMRI , AMZA_KLASA , AMZA_INDEKSI , AMZA_VITI_SHKOLLOR, AMZA_ATESIA , AMZA_MEMESIA , AMZA_SEKSI , AMZA_VENDLINDJA , AMZA_DATELINDJA FROM TBL_AMZA WHERE ( AMZA_NR_AMZA = '" & numerAmze & "' AND AMZA_CIKLI = " & strCikli & ")"
         tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic

         'get_rec (q_res)
         If Not tmp.EOF Then
            active_form.txtEmri.Text = tmp.Fields("AMZA_EMRI")
            active_form.txtMbiemri.Text = tmp.Fields("AMZA_MBIEMRI")
            active_form.txtAtesia.Text = tmp.Fields("AMZA_ATESIA")
            active_form.cboKlasa.Text = tmp.Fields("AMZA_KLASA")
            active_form.cboIndeksi.Text = tmp.Fields("AMZA_INDEKSI")
            active_form.txtMemesia.Text = tmp.Fields("AMZA_MEMESIA")
            active_form.cboSeksi.Text = tmp.Fields("AMZA_SEKSI")
            active_form.txtVendlindja.Text = tmp.Fields("AMZA_VENDLINDJA")
            active_form.txtDatelindja.Value = CDate(tmp.Fields("AMZA_DATELINDJA"))
            active_form.cboVitiShkollor.Text = tmp.Fields("AMZA_VITI_SHKOLLOR")
            tmp.Close
            active_form.txtMbulo.Visible = False
         Else: MsgBox "   Nxenesi me numrin e amzes se dhene " & Chr(10) & " nuk eshte i regjistruar ne ciklin qe eshte dhene ", vbInformation, "  Modifikime gjeneralitete"
            tmp.Close
            active_form.txtAmzaNo.Text = ""
            active_form.txtEmri.Text = ""
            active_form.txtMbiemri.Text = ""
            active_form.cboKlasa.Text = ""
            active_form.cboIndeksi.Text = ""
            active_form.txtAtesia.Text = ""
            active_form.txtMemesia.Text = ""
            active_form.cboSeksi.Text = ""
            active_form.txtVendlindja.Text = ""
            active_form.cboVitiShkollor.Text = ""
            active_form.txtMbulo.Visible = True
            active_form.cmdOK.Enabled = False
            Exit Sub
         End If
      Else
         emri = active_form.txtEmri.Text
         mbiemri = active_form.txtMbiemri.Text
         atesia = active_form.txtAtesia.Text
         q_res = "SELECT AMZA_EMRI,AMZA_MBIEMRI,AMZA_CIKLI, AMZA_NR_AMZA , AMZA_KLASA , AMZA_INDEKSI ,AMZA_VITI_SHKOLLOR, AMZA_ATESIA , AMZA_MEMESIA , AMZA_SEKSI , AMZA_VENDLINDJA , AMZA_DATELINDJA FROM TBL_AMZA WHERE ( AMZA_EMRI LIKE '" & emri & "%' AND AMZA_MBIEMRI LIKE '" & mbiemri & "%' AND AMZA_ATESIA LIKE '" & atesia & "%' ) AND  (AMZA_CIKLI = " & strCikli & ") ORDER BY AMZA_EMRI, AMZA_MBIEMRI"
         tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
         'get_rec (q_res)
         If Not tmp.EOF Then
            If tmp.RecordCount = 1 Then
                active_form.txtAmzaNo.Text = tmp.Fields("AMZA_NR_AMZA")
                active_form.cboKlasa.Text = tmp.Fields("AMZA_KLASA")
                active_form.cboIndeksi.Text = tmp.Fields("AMZA_INDEKSI")
                active_form.txtAtesia.Text = tmp.Fields("AMZA_ATESIA")
                active_form.txtEmri.Text = tmp.Fields("AMZA_EMRI")
                active_form.txtMbiemri.Text = tmp.Fields("AMZA_MBIEMRI")
                active_form.txtMemesia.Text = tmp.Fields("AMZA_MEMESIA")
                active_form.cboSeksi.Text = tmp.Fields("AMZA_SEKSI")
                active_form.txtVendlindja.Text = tmp.Fields("AMZA_VENDLINDJA")
                active_form.txtDatelindja.Value = CDate(tmp.Fields("AMZA_DATELINDJA"))
                active_form.cboVitiShkollor.Text = tmp.Fields("AMZA_VITI_SHKOLLOR")
                active_form.txtMbulo.Visible = False
                tmp.Close
            Else
                'nqs ka me teper se nje nxenes me keto te dhena atehere shfaq formen e kerkimit
                'inicializo gridat ne forme
            
                nrRreshta = tmp.RecordCount
                Call inicializoGridaZgjidhNxenes(nrRreshta)
                'mbush griden me te dhena
                Index = 0
                Do While Not tmp.EOF
                    frmZgjidhNxenes.gridaNxenesit.row = Index
                    frmZgjidhNxenes.gridaNxenesit.col = 0
                    frmZgjidhNxenes.gridaNxenesit.Text = tmp.Fields("AMZA_NR_AMZA").Value
                    frmZgjidhNxenes.gridaNxenesit.col = 1
                    frmZgjidhNxenes.gridaNxenesit.Text = tmp.Fields("AMZA_EMRI").Value + " " + tmp.Fields("AMZA_ATESIA").Value + " " + tmp.Fields("AMZA_MBIEMRI").Value
                    frmZgjidhNxenes.gridaNxenesit.col = 2
                    frmZgjidhNxenes.gridaNxenesit.Text = tmp.Fields("AMZA_CIKLI").Value
                    tmp.MoveNext
                    Index = Index + 1
                Loop
                tmp.Close
                ShkollaManager.frmZgjidhNxenes.show vbModal
                numerAmze = ShkollaManager.frmZgjidhNxenes.Amza
                strCikli = ShkollaManager.frmZgjidhNxenes.cikli
                If (numerAmze = "" Or strCikli = "") Then
                    'pastro
                    active_form.txtAmzaNo.Text = ""
                    active_form.txtEmri.Text = ""
                    active_form.txtMbiemri.Text = ""
                    active_form.cboKlasa.Text = ""
                    active_form.cboIndeksi.Text = ""
                    active_form.txtAtesia.Text = ""
                    active_form.txtMemesia.Text = ""
                    active_form.cboSeksi.Text = ""
                    active_form.txtVendlindja.Text = ""
                    active_form.cboVitiShkollor.Text = ""
                    active_form.cmdOK.Enabled = False
                    active_form.txtMbulo.Visible = True
                End If
                q_res = "SELECT AMZA_NR_AMZA,AMZA_EMRI , AMZA_MBIEMRI , AMZA_KLASA , AMZA_INDEKSI , AMZA_VITI_SHKOLLOR, AMZA_ATESIA , AMZA_MEMESIA , AMZA_SEKSI , AMZA_VENDLINDJA , AMZA_DATELINDJA FROM TBL_AMZA WHERE ( AMZA_NR_AMZA = '" & numerAmze & "' AND AMZA_CIKLI = " & strCikli & ")"
                tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
                If Not tmp.EOF Then
                    active_form.txtAmzaNo.Text = tmp.Fields("AMZA_NR_AMZA")
                    active_form.cboKlasa.Text = tmp.Fields("AMZA_KLASA")
                    active_form.cboIndeksi.Text = tmp.Fields("AMZA_INDEKSI")
                    active_form.txtAtesia.Text = tmp.Fields("AMZA_ATESIA")
                    active_form.txtMemesia.Text = tmp.Fields("AMZA_MEMESIA")
                    active_form.cboSeksi.Text = tmp.Fields("AMZA_SEKSI")
                    active_form.txtVendlindja.Text = tmp.Fields("AMZA_VENDLINDJA")
                    active_form.txtDatelindja.Value = CDate(tmp.Fields("AMZA_DATELINDJA"))
                    active_form.cboVitiShkollor.Text = tmp.Fields("AMZA_VITI_SHKOLLOR")
                    active_form.txtEmri.Text = tmp.Fields("AMZA_EMRI")
                    active_form.txtMbiemri.Text = tmp.Fields("AMZA_MBIEMRI")
                    active_form.txtMbulo.Visible = False
                    tmp.Close
                End If
            End If
            
         Else: MsgBox "   Nxenesi me kete emer dhe mbiemer " & Chr(10) & "nuk eshte i regjistruar ne ciklin qe eshte dhene ", vbInformation, "  Modifikime gjeneralitete"
            tmp.Close
            active_form.txtAmzaNo.Text = ""
            active_form.txtEmri.Text = ""
            active_form.txtMbiemri.Text = ""
            active_form.cboKlasa.Text = ""
            active_form.cboIndeksi.Text = ""
            active_form.txtAtesia.Text = ""
            active_form.txtMemesia.Text = ""
            active_form.cboSeksi.Text = ""
            active_form.txtVendlindja.Text = ""
            active_form.cboVitiShkollor.Text = ""
            active_form.cmdOK.Enabled = False
            active_form.txtMbulo.Visible = True
            Exit Sub
         End If
      End If

      active_form.cmdOK.Enabled = True

      nrAmzaModifiko = numerAmze
      emerModifiko = active_form.txtEmri.Text
      mbiemerModifiko = active_form.txtMbiemri.Text
      atesiaModifiko = active_form.txtAtesia.Text
      klasaModifiko = active_form.cboKlasa.Text
      indeksiModifiko = active_form.cboIndeksi.Text
      nrVitiShkollorModifiko = active_form.cboVitiShkollor.Text
      cikel = cikli
    Case KERKO_NXENES_PER_LARGIM

      numerAmze = active_form.txtAmzaNo.Text
      If active_form.optMesme.Value = True Then
         strCikli = "1"
      Else
         strCikli = "0"
      End If

      If numerAmze = "" Then
         emri = active_form.txtEmri.Text
         mbiemri = active_form.txtMbiemri.Text
         atesia = active_form.txtAtesia.Text
         q_res = "SELECT AMZA_NR_AMZA,AMZA_EMRI,AMZA_MBIEMRI,AMZA_ATESIA,AMZA_CIKLI FROM TBL_AMZA WHERE AMZA_EMRI like '" & emri & "%' AND AMZA_MBIEMRI like '" & mbiemri & "%' AND AMZA_ATESIA like '" & atesia & "%' AND AMZA_CIKLI = " & strCikli & " AND AMZA_LARGUAR = 0"
         tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
         If tmp.EOF Then
            If statusi = "SupervizorEmesme" Or statusi = "SupervizorTetevjecare" Then
               MsgBox "Nxenesi me keto te dhena nuk eshte nxenes i kesaj shkolle " & Chr(10) & "  ose eshte larguar.", vbInformation, "Elimino nxenes."
            Else

               MsgBox "Nxenesi me keto te dhena nuk eshte nxenes i kesaj shkolle " & Chr(10) & " ne ciklin e dhene nga ju ose eshte larguar.", vbInformation, "Elimino nxenes."
            End If
            active_form.txtEmri.Text = ""
            active_form.txtMbiemri.Text = ""
            active_form.txtAtesia.Text = ""
            active_form.optMesme.Value = False
            active_form.optTetevjecare.Value = False
            active_form.txtAmzaNo.Text = ""
            GoTo fund
         Else
            If tmp.RecordCount = 1 Then
                numerAmze = tmp.Fields("AMZA_NR_AMZA")
            Else
                 'nqs ka me teper se nje nxenes me keto te dhena atehere shfaq formen e kerkimit
                'inicializo gridat ne forme
                nrRreshta = tmp.RecordCount
                Call inicializoGridaZgjidhNxenes(nrRreshta)
                'mbush griden me te dhena
                Index = 0
                Do While Not tmp.EOF
                    frmZgjidhNxenes.gridaNxenesit.row = Index
                    frmZgjidhNxenes.gridaNxenesit.col = 0
                    frmZgjidhNxenes.gridaNxenesit.Text = tmp.Fields("AMZA_NR_AMZA").Value
                    frmZgjidhNxenes.gridaNxenesit.col = 1
                    frmZgjidhNxenes.gridaNxenesit.Text = tmp.Fields("AMZA_EMRI").Value + " " + tmp.Fields("AMZA_ATESIA").Value + " " + tmp.Fields("AMZA_MBIEMRI").Value
                    frmZgjidhNxenes.gridaNxenesit.col = 2
                    frmZgjidhNxenes.gridaNxenesit.Text = tmp.Fields("AMZA_CIKLI").Value
                    tmp.MoveNext
                    Index = Index + 1
                Loop
                ShkollaManager.frmZgjidhNxenes.show vbModal
                numerAmze = ShkollaManager.frmZgjidhNxenes.Amza
                strCikli = ShkollaManager.frmZgjidhNxenes.cikli
                If (numerAmze = "" Or strCikli = "") Then
                    'pastro
                    active_form.txtEmri.Text = ""
                    active_form.txtMbiemri.Text = ""
                    active_form.txtAtesia.Text = ""
                    active_form.txtAmzaNo.Text = ""
                    If statusi = "Administrator" Then
                       active_form.optMesme.Value = False
                       active_form.optTetevjecare.Value = False
                    End If
                    Exit Sub
                End If
                
            End If
         End If
         tmp.Close
      End If
      q_res = "SELECT  * FROM TBL_AMZA WHERE AMZA_NR_AMZA = '" & numerAmze & "' AND AMZA_CIKLI = " & strCikli & " AND AMZA_LARGUAR = 0"
      tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
      If Not tmp.EOF Then
         active_form.txtAmzaNo.Text = tmp.Fields("AMZA_NR_AMZA")
         active_form.txtEmri.Text = tmp.Fields("AMZA_EMRI")
         active_form.txtMbiemri.Text = tmp.Fields("AMZA_MBIEMRI")
         active_form.txtKlasa.Text = tmp.Fields("AMZA_KLASA")
         active_form.txtIndeksi.Text = tmp.Fields("AMZA_INDEKSI")
         active_form.txtVitiShkollor.Text = tmp.Fields("AMZA_VITI_SHKOLLOR")
         active_form.txtAtesia.Text = tmp.Fields("AMZA_ATESIA")
         active_form.txtMemesia.Text = tmp.Fields("AMZA_MEMESIA")
         active_form.txtSeksi.Text = tmp.Fields("AMZA_SEKSI")
         active_form.txtVendlindja.Text = tmp.Fields("AMZA_VENDLINDJA")
         active_form.txtDatelindja.Text = tmp.Fields("AMZA_DATELINDJA")
         active_form.cmdOK.Enabled = True
      Else
         If statusi = "SupervizorEmesme" Or statusi = "SupervizorTetevjecare" Then
            MsgBox "Nxenesi me keto te dhena nuk eshte nxenes i kesaj shkolle " & Chr(10) & "  ose eshte larguar.", vbInformation, "Elimino nxenes."
         Else

            MsgBox "Nxenesi me keto te dhena nuk eshte nxenes i kesaj shkolle " & Chr(10) & " ne ciklin e dhene nga ju ose eshte larguar.", vbInformation, "Elimino nxenes."
         End If
         active_form.txtEmri.Text = ""
         active_form.txtMbiemri.Text = ""
         active_form.txtAtesia.Text = ""
         active_form.txtAmzaNo.Text = ""
         If statusi = "Administrator" Then
            active_form.optMesme.Value = False
            active_form.optTetevjecare.Value = False
         End If
      End If
fund:

    Case VIZUALIZO_MUNGESA_ACTION
        'kontrollojme nqs klasa e zgjedhur nga perdoruesi ka lende te konfiguruara
        Dim vitishkollor As String
        Klasa = active_form.cboKlasa.Text
        indeksi = active_form.cboIndeksi.Text
        vitishkollor = active_form.cboVitiShkollor.Text
        q_res = "SELECT LN_EMRI  FROM TBL_LENDA " & _
                "WHERE LN_KLASA" + Klasa + " = 1 AND LN_VITI_SHKOLLOR = '" + vitishkollor + "'"
        tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
        If tmp.EOF Then
            MsgBox "Klasa që keni zgjedhur nuk ka asnjë lëndë të konfiguruar për vitin shkollor " + vitishkollor + "!", vbInformation, "Modifiko Mungesa"
            Exit Sub
        End If
        tmp.Close
        'kontrollojme nqs klasa e zgjedhur ka nxenes
        q_res = "SELECT AMZA_EMRI , AMZA_MBIEMRI, AMZA_NR_AMZA, KL_KLASA, KL_INDEKSI, KL_VITI_SHKOLLOR FROM TBL_AMZA, TBL_KLASA  WHERE (AMZA_NR_AMZA = KL_NR_AMZA AND AMZA_CIKLI = KL_CIKLI) AND (KL_KLASA = '" & Klasa & "') AND (KL_INDEKSI = '" & indeksi & "') AND (KL_VITI_SHKOLLOR = '" & vitishkollor & "') AND (AMZA_LARGUAR = 0) ORDER BY AMZA_EMRI, AMZA_MBIEMRI "
        tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
        If tmp.EOF Then
            MsgBox "Klasa e zgjedhur nga ju nuk ka asnjë nxënës për vitin shkollor " + vitishkollor + "!", vbInformation, "Modifiko Mungesa"
            Exit Sub
        End If
        
        'Shfaqet grida me emrat e nxenesve
        
        Call inicializoGridaMungesat(active_form.amzaEmerMbiemer, active_form.emrat, CInt(tmp.RecordCount), active_form.dtpNga.Value, active_form.dtpDeri.Value, active_form.datat, active_form.mungesat, active_form.picUnchecked)
        i = 0
        Do While Not tmp.EOF
            active_form.emrat.row = i
            active_form.emrat.col = 0
            active_form.emrat.Text = i + 1
            active_form.emrat.CellAlignment = 1
        
            active_form.emrat.row = i
            active_form.emrat.col = 1
            active_form.emrat.Text = tmp.Fields("AMZA_NR_AMZA")
            active_form.emrat.CellAlignment = 1
            
            active_form.emrat.row = i
            active_form.emrat.col = 2
            active_form.emrat.Text = tmp.Fields("AMZA_EMRI") & Space(2) & tmp.Fields("AMZA_MBIEMRI")
            active_form.emrat.CellAlignment = 1
            tmp.MoveNext
            i = i + 1
        Loop
        tmp.Close
        'mbushet grida me mungesat
        q_res = "SELECT *, CONVERT(DATETIME,NT_DATA) AS DATA FROM " & _
            " (SELECT Q1.*, TBL_KLASA.KL_KLASA, TBL_KLASA.KL_INDEKSI FROM " & _
            " (SELECT COUNT(NT_VLERESIMI) AS NR,UN, NT_NR_AMZA, NT_CIKLI, " & _
            " NT_TEKST_DATA,NT_DATA,NT_VITI_SHKOLLOR FROM " & _
            " (SELECT TBL_NOTA.*, NT_NR_AMZA + '*'+ Convert(varchar(1),NT_CIKLI) + '*' + NT_DATA AS UN " & _
            " From TBL_NOTA " & _
            " WHERE CONVERT( DATETIME, NT_DATA) BETWEEN '" + konvertoDaten(CDate(active_form.dtpNga.Value)) + "' AND '" + konvertoDaten(CDate(active_form.dtpDeri.Value)) + "' " & _
            " AND NT_VLERESIMI = 'm' AND NT_VITI_SHKOLLOR = '" + vitishkollor + "' " & _
            " ) AS Q " & _
            " GROUP BY UN, NT_NR_AMZA,NT_CIKLI, NT_TEKST_DATA,NT_DATA,NT_VITI_SHKOLLOR) " & _
            " AS Q1,TBL_KLASA " & _
            " Where Q1.NT_NR_AMZA = TBL_KLASA.KL_NR_AMZA " & _
            " AND Q1.NT_CIKLI = TBL_KLASA.KL_CIKLI " & _
            " AND Q1.NT_VITI_SHKOLLOR = TBL_KLASA.KL_VITI_SHKOLLOR) AS Q2 " & _
            " WHERE KL_KLASA = '" + Klasa + "' AND KL_INDEKSI = '" + indeksi + "' " & _
            " ORDER BY NT_NR_AMZA, DATA "
        tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
        active_form.emrat.row = 0
        active_form.emrat.col = 1
        Dim currAmza As String
        currAmza = active_form.emrat.Text
        Index = 0
        Do While Not tmp.EOF
            If (currAmza <> tmp.Fields("NT_NR_AMZA").Value) Then
                currAmza = tmp.Fields("NT_NR_AMZA").Value
                'gjej I
                For i = 0 To active_form.emrat.Rows - 1
                    active_form.emrat.col = 1
                    active_form.emrat.row = i
                    If (active_form.emrat.Text = tmp.Fields("NT_NR_AMZA").Value) Then
                        Index = i
                        Exit For
                    End If
                Next i
            End If
            active_form.mungesat.row = Index
            Dim d1, d2 As Date
            d1 = CDate(Mid(CStr(active_form.dtpNga.Value), 1, 10))
            d2 = CDate(Mid(CStr(tmp.Fields("DATA").Value), 1, 10))
            active_form.mungesat.col = (d2 - d1) * 2
            Dim nrMungesa As Integer
            nrMungesa = CInt(tmp.Fields("NR").Value)
            Dim txt As String
            txt = ""
            For k = 1 To nrMungesa
                txt = txt + "m "
            Next k
            txt = txt + "(" + CStr(nrMungesa) + ")"
            active_form.mungesat.Text = txt
            tmp.MoveNext
        Loop
        tmp.Close
        
        'Behen gridat te dukshme
        active_form.amzaEmerMbiemer.Visible = True
        active_form.emrat.Visible = True
        active_form.datat.Visible = True
        active_form.mungesat.Visible = True
        'formatohen scrollBaret
        Call FormatoScrollBarMungesa(active_form.amzaEmerMbiemer, active_form.emrat, active_form.datat, active_form.mungesat, active_form.fsbHorizontali, active_form.fsbVertikali)
        
    Case MUNGESA_ME_ARSYE_ACTION
        ModifikoMungesa "MeArsye"
    Case MUNGESA_PA_ARSYE_ACTION
        ModifikoMungesa "PaArsye"
    Case FSHI_MUNGESA_ACTION
        ModifikoMungesa "FshiMungesa"
    Case VIZUALIZO_MESATARET_MOMENTALE_ACTION
        'kontrollojme nqs klasa e zgjedhur nga perdoruesi ka lende te konfiguruara
        Klasa = active_form.cboKlasa.Text
        indeksi = active_form.cboIndeksi.Text
        vitishkollor = active_form.cboVitiShkollor.Text
        If (ktheCiklin(Klasa, CStr(vitishkollor)) = False) Then
            cikli = 0
            q_res = "SELECT LN_EMRI  FROM TBL_LENDA " & _
                "WHERE LN_KLASA" + Klasa + " = 1 AND LN_VITI_SHKOLLOR = '" + vitishkollor + "' ORDER BY LN_NUMRI_ULET"
        Else
            cikli = 1
            q_res = "SELECT LN_EMRI  FROM TBL_LENDA " & _
                "WHERE LN_KLASA" + Klasa + " = 1 AND LN_VITI_SHKOLLOR = '" + vitishkollor + "' ORDER BY LN_NUMRI_MESME"
        End If
        tmp1.Open q_res, conn, adOpenDynamic, adLockOptimistic
        If tmp1.EOF Then
            MsgBox "Klasa që keni zgjedhur nuk ka asnjë lëndë të konfiguruar për vitin shkollor " + vitishkollor + "!", vbInformation, "Modifiko Mungesa"
            Exit Sub
        End If
        Dim nrLende As Integer
        nrLende = tmp1.RecordCount
        'kontrollojme nqs klasa e zgjedhur ka nxenes
        'q_res = "SELECT AMZA_NR_AMZA,AMZA_EMRI, AMZA_MBIEMRI " & _
                '" FROM TBL_AMZA WHERE AMZA_KLASA = '" + Klasa + "' " & _
                '" AND AMZA_INDEKSI = '" + indeksi + "' AND AMZA_VITI_SHKOLLOR = '" + vitishkollor + "' ORDER BY AMZA_EMRI, AMZA_MBIEMRI"
        q_res = "SELECT AMZA_EMRI , AMZA_MBIEMRI, AMZA_NR_AMZA, KL_KLASA, KL_INDEKSI, KL_VITI_SHKOLLOR FROM TBL_AMZA, TBL_KLASA  WHERE (AMZA_NR_AMZA = KL_NR_AMZA AND AMZA_CIKLI = KL_CIKLI) AND (KL_KLASA = '" & Klasa & "') AND (KL_INDEKSI = '" & indeksi & "') AND (KL_VITI_SHKOLLOR = '" & vitishkollor & "') AND (AMZA_LARGUAR = 0) ORDER BY AMZA_EMRI, AMZA_MBIEMRI "
        tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
        If tmp.EOF Then
            MsgBox "Klasa e zgjedhur nga ju nuk ka asnjë nxënës për vitin shkollor " + vitishkollor + "!", vbInformation, "Modifiko Mungesa"
            Exit Sub
        End If
        
        'Shfaqet grida me emrat e nxenesve
        
        Call inicializoGridaMesataret(active_form.amzaEmerMbiemer, active_form.emrat, CInt(tmp.RecordCount), nrLende, active_form.Lendet, active_form.notat)
        i = 0
        Do While Not tmp.EOF
            active_form.emrat.row = i
            active_form.emrat.col = 0
            active_form.emrat.Text = i + 1
            active_form.emrat.CellAlignment = 1
        
            active_form.emrat.row = i
            active_form.emrat.col = 1
            active_form.emrat.Text = tmp.Fields("AMZA_NR_AMZA")
            active_form.emrat.CellAlignment = 1
            
            active_form.emrat.row = i
            active_form.emrat.col = 2
            active_form.emrat.Text = tmp.Fields("AMZA_EMRI") & Space(2) & tmp.Fields("AMZA_MBIEMRI")
            active_form.emrat.CellAlignment = 1
            tmp.MoveNext
            i = i + 1
        Loop
        tmp.Close
        active_form.emrat.row = i
        active_form.emrat.col = 2
        active_form.emrat.Text = "Mesatarja"
        'Mbushet grida me lendet perkatese
        i = 0
        active_form.Lendet.row = 0
        Do While Not tmp1.EOF
            active_form.Lendet.col = i
            active_form.Lendet.Text = CStr(tmp1.Fields("LN_EMRI"))
            i = i + 1
            tmp1.MoveNext
        Loop
        tmp1.Close
        active_form.Lendet.col = i
        active_form.Lendet.Text = "Mesatarja"
        'mbushet grida me mesataret
        If (active_form.optSemestri1) Then
            q_res = "SELECT NT_NR_AMZA,NT_EMRI_LENDA,CONVERT(VARCHAR,ROUND(SUM(CONVERT(DECIMAL, NT_VLERESIMI))/COUNT(NT_VLERESIMI), 2)) " & _
                    " AS MESATARJA FROM TBL_NOTA WHERE NT_VITI_SHKOLLOR = '" + vitishkollor + "' " & _
                    " AND NT_MOMENTALES1 = 1 AND NT_MUNGESE_ME = 0 And NT_MUNGESE_PA = 0 " & _
                    " AND NT_KLASA = '" + Klasa + "' AND NT_INDEKSI = '" + indeksi + "'" & _
                    " GROUP BY  NT_NR_AMZA,NT_EMRI_LENDA ORDER BY NT_NR_AMZA"
        Else
            q_res = "SELECT NT_NR_AMZA,NT_EMRI_LENDA,CONVERT(VARCHAR,ROUND(SUM(CONVERT(DECIMAL, NT_VLERESIMI))/COUNT(NT_VLERESIMI), 2)) " & _
                    " AS MESATARJA FROM TBL_NOTA WHERE NT_VITI_SHKOLLOR = '" + vitishkollor + "' " & _
                    " AND NT_MOMENTALES2 = 1 AND NT_MUNGESE_ME = 0 And NT_MUNGESE_PA = 0 " & _
                    " AND NT_KLASA = '" + Klasa + "' AND NT_INDEKSI = '" + indeksi + "'" & _
                    " GROUP BY  NT_NR_AMZA,NT_EMRI_LENDA ORDER BY NT_NR_AMZA"
        End If
        For i = 0 To active_form.notat.Rows - 1
            active_form.notat.row = i
            For j = 0 To active_form.notat.Cols - 1
                active_form.notat.col = j
                active_form.notat.Text = "--"
                active_form.notat.CellFontBold = False
            Next
        Next
        tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
        'mesataret per secilin nxenes
        If (active_form.optSemestri1) Then
            q_res = "SELECT NT_NR_AMZA," & _
                    " CONVERT(VARCHAR,ROUND(SUM(CONVERT(DECIMAL, NT_VLERESIMI))/COUNT(NT_VLERESIMI), 2)) " & _
                     " AS MESATARJA FROM TBL_NOTA " & _
                    " WHERE NT_VITI_SHKOLLOR = '" + vitishkollor + "' " & _
                    " AND NT_MOMENTALES1 = 1 AND " & _
                    " NT_MUNGESE_ME = 0 And NT_MUNGESE_PA = 0 " & _
                    " AND NT_KLASA = '" + Klasa + "' AND NT_INDEKSI = '" + indeksi + "' " & _
                    " GROUP BY  NT_NR_AMZA " & _
                    " ORDER BY NT_NR_AMZA "
        Else
            q_res = "SELECT NT_NR_AMZA, " & _
                    " CONVERT(VARCHAR,ROUND(SUM(CONVERT(DECIMAL, NT_VLERESIMI))/COUNT(NT_VLERESIMI), 2)) " & _
                     " AS MESATARJA FROM TBL_NOTA " & _
                    " WHERE NT_VITI_SHKOLLOR = '" + vitishkollor + "' " & _
                    " AND NT_MOMENTALES2 = 1 AND " & _
                    " NT_MUNGESE_ME = 0 And NT_MUNGESE_PA = 0 " & _
                    " AND NT_KLASA = '" + Klasa + "' AND NT_INDEKSI = '" + indeksi + "' " & _
                    " GROUP BY  NT_NR_AMZA " & _
                    " ORDER BY NT_NR_AMZA "
        End If
        tmp1.Open q_res, conn, adOpenDynamic, adLockOptimistic
        
        'mesatarja sipas lendeve
        If (active_form.optSemestri1) Then
            q_res = "SELECT NT_EMRI_LENDA, " & _
                " CONVERT(VARCHAR,ROUND(SUM(CONVERT(DECIMAL, NT_VLERESIMI))/COUNT(NT_VLERESIMI), 2)) " & _
                " AS MESATARJA FROM TBL_NOTA " & _
                " WHERE NT_VITI_SHKOLLOR = '" + vitishkollor + "' " & _
                " AND NT_MOMENTALES1 = 1 AND " & _
                " NT_MUNGESE_ME = 0 And NT_MUNGESE_PA = 0 " & _
                " AND NT_KLASA = '" + Klasa + "' AND NT_INDEKSI = '" + indeksi + "' " & _
                " GROUP BY  NT_EMRI_LENDA " & _
                " ORDER BY NT_EMRI_LENDA "
        Else
            q_res = "SELECT NT_EMRI_LENDA, " & _
                " CONVERT(VARCHAR,ROUND(SUM(CONVERT(DECIMAL, NT_VLERESIMI))/COUNT(NT_VLERESIMI), 2)) " & _
                " AS MESATARJA FROM TBL_NOTA " & _
                " WHERE NT_VITI_SHKOLLOR = '" + vitishkollor + "' " & _
                " AND NT_MOMENTALES2 = 1 AND " & _
                " NT_MUNGESE_ME = 0 And NT_MUNGESE_PA = 0 " & _
                " AND NT_KLASA = '" + Klasa + "' AND NT_INDEKSI = '" + indeksi + "' " & _
                " GROUP BY  NT_EMRI_LENDA " & _
                " ORDER BY NT_EMRI_LENDA "
        End If
        tmp11.Open q_res, conn, adOpenDynamic, adLockOptimistic
        
        currAmza = ""
        Dim indeks As Integer
        active_form.Lendet.row = 0
        ugjet = False
        Dim mes As String
        Do While Not tmp.EOF
            'gjendet rreshti ne gride per nxenesin
            If (tmp.Fields("NT_NR_AMZA") <> currAmza) Then
                active_form.emrat.col = 1
                For i = 0 To active_form.emrat.Rows - 1
                    active_form.emrat.row = i
                    If (active_form.emrat.Text = tmp.Fields("NT_NR_AMZA")) Then
                        active_form.notat.row = i
                        currAmza = CStr(tmp.Fields("NT_NR_AMZA"))
                        'mesatarja e pergjithshme per nxenesin
                        ugjet = False
                        tmp1.MoveFirst
                        Do While Not tmp1.EOF And Not ugjet
                            If tmp1.Fields("NT_NR_AMZA") = currAmza Then
                                active_form.notat.col = active_form.notat.Cols - 1
                                active_form.notat.CellFontBold = True
                                mes = tmp1.Fields("MESATARJA")
                                If (Len(mes) = 8) Then
                                    mes = Mid(mes, 1, 4)
                                Else
                                    mes = Mid(mes, 1, 5)
                                End If
                                active_form.notat.Text = mes
                                ugjet = True
                            End If
                            tmp1.MoveNext
                        Loop
                        Exit For
                    End If
                Next
            End If
            indeks = active_form.notat.row
            'gjendet kolona ne gride per lenden
            For j = 0 To active_form.Lendet.Cols
                active_form.Lendet.col = j
                If (tmp.Fields("NT_EMRI_LENDA") = active_form.Lendet.Text) Then
                    active_form.notat.col = j
                    Exit For
                End If
            Next
            mes = tmp.Fields("MESATARJA")
            If (Len(mes) = 8) Then
                mes = Mid(mes, 1, 4)
            Else
                mes = Mid(mes, 1, 5)
            End If
            active_form.notat.Text = mes
            'vendoset mesatarja per lenden
            active_form.notat.row = active_form.notat.Rows - 1
            If (active_form.notat.Text = "--") Then
                ugjet = False
                tmp11.MoveFirst
                Do While Not tmp11.EOF And Not ugjet
                    If (tmp11.Fields("NT_EMRI_LENDA") = tmp.Fields("NT_EMRI_LENDA")) Then
                        active_form.notat.CellFontBold = True
                        mes = tmp11.Fields("MESATARJA")
                        If (Len(mes) = 8) Then
                            mes = Mid(mes, 1, 4)
                        Else
                            mes = Mid(mes, 1, 5)
                        End If
                        active_form.notat.Text = mes
                        ugjet = True
                    End If
                    tmp11.MoveNext
                Loop
            End If
            active_form.notat.row = indeks
            tmp.MoveNext
        Loop
        tmp.Close
        tmp1.Close
        tmp11.Close
        'mesatarja e pergjithshme
        If (active_form.optSemestri1) Then
            q_res = " SET CONCAT_NULL_YIELDS_NULL OFF SELECT " & _
                    " CONVERT(VARCHAR,ROUND(SUM(CONVERT(DECIMAL, NT_VLERESIMI))/COUNT(NT_VLERESIMI), 2)) + ''" & _
                    " AS MESATARJA FROM TBL_NOTA " & _
                    " WHERE NT_VITI_SHKOLLOR = '" + vitishkollor + "' " & _
                    " AND NT_MOMENTALES1 = 1 AND " & _
                    " NT_MUNGESE_ME = 0 And NT_MUNGESE_PA = 0 " & _
                    " AND NT_KLASA = '" + Klasa + "' AND NT_INDEKSI = '" + indeksi + "' SET CONCAT_NULL_YIELDS_NULL ON"
        Else
            q_res = " SET CONCAT_NULL_YIELDS_NULL OFF SELECT " & _
                    " CONVERT(VARCHAR,ROUND(SUM(CONVERT(DECIMAL, NT_VLERESIMI))/COUNT(NT_VLERESIMI), 2)) + ''" & _
                    " AS MESATARJA FROM TBL_NOTA " & _
                    " WHERE NT_VITI_SHKOLLOR = '" + vitishkollor + "' " & _
                    " AND NT_MOMENTALES2 = 1 AND " & _
                    " NT_MUNGESE_ME = 0 And NT_MUNGESE_PA = 0 " & _
                    " AND NT_KLASA = '" + Klasa + "' AND NT_INDEKSI = '" + indeksi + "' SET CONCAT_NULL_YIELDS_NULL ON"
        End If
        tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
        Do While Not tmp.EOF
            active_form.notat.row = active_form.notat.Rows - 1
            active_form.notat.col = active_form.notat.Cols - 1
            active_form.notat.CellFontBold = True
            mes = tmp.Fields("MESATARJA")
            If (mes = "") Then
                mes = "--"
            End If
            If (Len(mes) = 8) Then
                mes = Mid(mes, 1, 4)
            Else
                mes = Mid(mes, 1, 5)
            End If
            active_form.notat.Text = mes
            tmp.MoveNext
        Loop
        tmp.Close
        'Behen gridat te dukshme
        active_form.amzaEmerMbiemer.Visible = True
        active_form.emrat.Visible = True
        active_form.Lendet.Visible = True
        active_form.notat.Visible = True
        'formatohen scrollBaret
        Call FormatoScrollBarMesataret(active_form.amzaEmerMbiemer, active_form.emrat, active_form.Lendet, active_form.notat, active_form.fsbHorizontali, active_form.fsbVertikali)
    Case VIZUALIZO_NXENESIT_DALLUAR_ACTION
        If (active_form.optUlet.Value) Then
            cikli = 0
        Else
            cikli = 1
        End If
        
        vitishkollor = active_form.cboVitiShkollor.Text
        'kontrollohet nqs jane regjistruar klasa per kete vit shkollor
        q_res = "SELECT LN_EMRI FROM TBL_LENDA WHERE LN_VITI_SHKOLLOR = '" + vitishkollor + "'"
        tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
        If (tmp.RecordCount = 0) Then
            tmp.Close
            MsgBox "Për vitin shkollor të zgjedhur nuk është konfiguruar asnjë klasë!", vbExclamation, "Nxënësit e dalluar"
            Exit Sub
        End If
        tmp.Close
        Dim kushti As String
        kushti = ""
        q_res = "SELECT AMZA_NR_AMZA, AMZA_KLASA, AMZA_INDEKSI," & _
                " UPPER(AMZA_EMRI + ' ' + AMZA_ATESIA + ' ' + AMZA_MBIEMRI) AS NXENESI FROM TBL_AMZA " & _
                " Where "
        If (active_form.cbNota.Value = 1) Then
            kushti = "AMZA_NR_AMZA NOT IN " & _
                    " (SELECT NT_NR_AMZA " & _
                    " FROM TBL_NOTA WHERE NT_VITI_SHKOLLOR = '" + vitishkollor + "' " & _
                    " AND NT_CIKLI = " + CStr(cikli) + " AND NT_VJETORE = 1 " & _
                    " AND CONVERT(DECIMAL, NT_VLERESIMI) < " + active_form.cboNota.Text + " " & _
                    " GROUP BY NT_NR_AMZA) "
        End If
        
        If (active_form.cbMesatarja.Value = 1) Then
            If kushti = "" Then
                kushti = " AMZA_NR_AMZA IN " & _
                        " (SELECT NT_NR_AMZA FROM " & _
                        " (SELECT NT_NR_AMZA,SUM(CONVERT(DECIMAL, NT_VLERESIMI))/COUNT(NT_VLERESIMI) AS MESATARJA " & _
                        " FROM TBL_NOTA WHERE NT_VITI_SHKOLLOR = '" + vitishkollor + "' " & _
                        " AND NT_CIKLI = " + CStr(cikli) + " AND NT_VJETORE = 1 " & _
                        " GROUP BY NT_NR_AMZA) AS Q " & _
                        " WHERE Q.MESATARJA >= " + active_form.txtMesatarja.Text + ") "
            Else
                kushti = kushti + " AND " + " AMZA_NR_AMZA IN " & _
                        " (SELECT NT_NR_AMZA FROM " & _
                        " (SELECT NT_NR_AMZA,SUM(CONVERT(DECIMAL, NT_VLERESIMI))/COUNT(NT_VLERESIMI) AS MESATARJA " & _
                        " FROM TBL_NOTA WHERE NT_VITI_SHKOLLOR = '" + vitishkollor + "' " & _
                        " AND NT_CIKLI = " + CStr(cikli) + " AND NT_VJETORE = 1 " & _
                        " GROUP BY NT_NR_AMZA) AS Q " & _
                        " WHERE Q.MESATARJA >= " + active_form.txtMesatarja.Text + ") "
            End If
        End If
        q_res = q_res + kushti
        q_res = q_res + "AND AMZA_CIKLI = " + CStr(cikli) + " AND AMZA_VITI_SHKOLLOR = '" + vitishkollor + "' " & _
                " AND AMZA_LARGUAR = 0 " & _
                " AND AMZA_NR_AMZA IN (SELECT NT_NR_AMZA FROM TBL_NOTA " & _
                " WHERE NT_VITI_SHKOLLOR = '" + vitishkollor + "' " & _
                " AND NT_CIKLI = " + CStr(cikli) + " AND NT_VJETORE = 1 " & _
                " GROUP BY NT_NR_AMZA) " & _
                " ORDER BY CONVERT(DECIMAL,AMZA_KLASA), AMZA_INDEKSI "
        tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
        Call inicializoGridaNxenesitDalluar(active_form.Header, active_form.Nxenesit, tmp.RecordCount)
        'mbush griden me te dhenat e nxenesve
        i = 0
        Do While Not tmp.EOF
            active_form.Nxenesit.row = i
            active_form.Nxenesit.RowHeight(i) = 300
            
            active_form.Nxenesit.col = 0
            active_form.Nxenesit.Text = CStr(i + 1)
            active_form.Nxenesit.CellAlignment = 1
            
            active_form.Nxenesit.col = 1
            active_form.Nxenesit.Text = CStr(tmp.Fields("AMZA_NR_AMZA"))
            active_form.Nxenesit.CellAlignment = 1
            
            active_form.Nxenesit.col = 2
            active_form.Nxenesit.Text = CStr(tmp.Fields("AMZA_KLASA"))
            active_form.Nxenesit.CellAlignment = 1
            
            active_form.Nxenesit.col = 3
            active_form.Nxenesit.Text = CStr(tmp.Fields("AMZA_INDEKSI"))
            active_form.Nxenesit.CellAlignment = 1
            
            active_form.Nxenesit.col = 4
            active_form.Nxenesit.Text = CStr(tmp.Fields("NXENESI"))
            active_form.Nxenesit.CellAlignment = 1
            
            tmp.MoveNext
            i = i + 1
        Loop
        'shfaq gridat
        active_form.Header.Visible = True
        If tmp.RecordCount > 0 Then
            active_form.Nxenesit.Visible = True
        End If
        tmp.Close
    Case VIZUALIZO_MESATARET_CIKLI_ACTION

        Dim nr1, nr2, nr3 As Integer
        vitishkollor = active_form.cboVitiShkollor.Text
        Dim vitFillimi As Integer
        vitFillimi = CInt(Mid(vitishkollor, 1, 4))
        Dim kushtUlet, kushtLart, kushtMesem As String
        Dim strMidisUlet, strMidisLart, strMidisMesem As String
        'sistemi i vjeter
        If (vitFillimi <= 2007) Then
            active_form.lblUlet.Caption = "Cikli i ulët(1 - 4)"
            active_form.lblTetevjecare.Caption = "Cikli i lartë(5 - 8)"
            active_form.lblMesme.Caption = "Shkolla e mesme(9 - 12)"
            kushtUlet = " (LN_KLASA1 = 1 OR LN_KLASA2 = 1 OR LN_KLASA3 = 1 OR LN_KLASA4 = 1) "
            kushtLart = " (LN_KLASA5 = 1 OR LN_KLASA6 = 1 OR LN_KLASA7 = 1 OR LN_KLASA8 = 1) "
            kushtMesem = " (LN_KLASA9 = 1 OR LN_KLASA10 = 1 OR LN_KLASA11 = 1 OR LN_KLASA12 = 1) "
            strMidisUlet = " 1 AND 4 "
            strMidisLart = " 5 AND 8 "
            strMidisMesem = " 9 AND 12 "
        'sistemi i ri
        Else
            active_form.lblUlet.Caption = "Cikli i ulët(1 - 5)"
            active_form.lblTetevjecare.Caption = "Cikli i lartë(6 - 9)"
            active_form.lblMesme.Caption = "Shkolla e mesme(10 - 12)"
            kushtUlet = " (LN_KLASA1 = 1 OR LN_KLASA2 = 1 OR LN_KLASA3 = 1 OR LN_KLASA4 = 1 OR LN_KLASA5 = 1) "
            kushtLart = " (LN_KLASA6 = 1 OR LN_KLASA7 = 1 OR LN_KLASA8 = 1 OR LN_KLASA9 = 1) "
            kushtMesem = " (LN_KLASA10 = 1 OR LN_KLASA11 = 1 OR LN_KLASA12 = 1) "
            strMidisUlet = " 1 AND 5 "
            strMidisLart = " 6 AND 9 "
            strMidisMesem = " 10 AND 12 "
       End If
        Dim strSem As String
        If (active_form.optSemestri1.Value) Then
            strSem = "NT_MOMENTALES1"
        Else
            strSem = "NT_MOMENTALES2"
        End If
        
        'klasa 1 deri 5
        q_res = " SET CONCAT_NULL_YIELDS_NULL OFF SELECT LN_EMRI, MESATARJA + '' AS MESATARJA FROM" & _
            " (SELECT NT_EMRI_LENDA, " & _
            " CONVERT(VARCHAR,ROUND(SUM(CONVERT(DECIMAL, NT_VLERESIMI))/COUNT(NT_VLERESIMI), 2)) " & _
            " AS MESATARJA FROM TBL_NOTA " & _
            " WHERE NT_VITI_SHKOLLOR = '" + vitishkollor + "' " & _
            " AND " + strSem + " = 1 AND " & _
            " NT_MUNGESE_ME = 0 And NT_MUNGESE_PA = 0 " & _
            " AND CONVERT(INT, NT_KLASA) BETWEEN 1 AND 5 " & _
            " GROUP BY  NT_EMRI_LENDA) AS A " & _
            " Right Join " & _
            " (SELECT LN_EMRI, LN_NUMRI_ULET FROM TBL_LENDA WHERE LN_VITI_SHKOLLOR = '" + vitishkollor + "' " & _
            " AND " + kushtUlet + " " & _
            " ) AS B ON A.NT_EMRI_LENDA = B.LN_EMRI " & _
            " ORDER BY LN_NUMRI_ULET SET CONCAT_NULL_YIELDS_NULL ON"
        tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
        If (tmp.RecordCount = 0) Then
            PastroGrideMesataretCikli active_form.Mesataret1
            tmp.Close
        Else
            MbushGrideMesataretCikli active_form.Mesataret1, tmp
            tmp.Close
            'mesatarja e pergjithshme per ciklin
            q_res = " SET CONCAT_NULL_YIELDS_NULL OFF" & _
                    " SELECT " & _
                    " CONVERT(VARCHAR,ROUND(SUM(CONVERT(DECIMAL, NT_VLERESIMI))/COUNT(NT_VLERESIMI), 2)) + '' " & _
                    " AS MESATARJA FROM TBL_NOTA " & _
                    " WHERE NT_VITI_SHKOLLOR = '" + vitishkollor + "'" & _
                    " AND " + strSem + " = 1 AND " & _
                    " NT_MUNGESE_ME = 0 And NT_MUNGESE_PA = 0 " & _
                    " AND CONVERT(INT, NT_KLASA) BETWEEN  " + strMidisUlet & _
                    " SET CONCAT_NULL_YIELDS_NULL ON"
            tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
            active_form.Mesataret1.row = active_form.Mesataret1.Rows - 1
            If (tmp.Fields("MESATARJA") = "") Then
                mes = "--"
            Else
                mes = tmp.Fields("MESATARJA")
                If (Len(mes) = 8) Then
                    mes = Mid(mes, 1, 4)
                Else
                    mes = Mid(mes, 1, 5)
                End If
            End If
            active_form.Mesataret1.Text = mes
            tmp.Close
        End If
        
        'klasa 6 - 9
        q_res = " SET CONCAT_NULL_YIELDS_NULL OFF SELECT LN_EMRI, MESATARJA + '' AS MESATARJA FROM" & _
            " (SELECT NT_EMRI_LENDA, " & _
            " CONVERT(VARCHAR,ROUND(SUM(CONVERT(DECIMAL, NT_VLERESIMI))/COUNT(NT_VLERESIMI), 2)) " & _
            " AS MESATARJA FROM TBL_NOTA " & _
            " WHERE NT_VITI_SHKOLLOR = '" + vitishkollor + "' " & _
            " AND " + strSem + " = 1 AND " & _
            " NT_MUNGESE_ME = 0 And NT_MUNGESE_PA = 0 " & _
            " AND CONVERT(INT, NT_KLASA) BETWEEN 6 AND 9 " & _
            " GROUP BY  NT_EMRI_LENDA) AS A " & _
            " Right Join " & _
            " (SELECT LN_EMRI, LN_NUMRI_ULET FROM TBL_LENDA WHERE LN_VITI_SHKOLLOR = '" + vitishkollor + "' " & _
            " AND " + kushtLart + " " & _
            " ) AS B ON A.NT_EMRI_LENDA = B.LN_EMRI " & _
            " ORDER BY LN_NUMRI_ULET SET CONCAT_NULL_YIELDS_NULL ON"
        tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
        If (tmp.RecordCount = 0) Then
            PastroGrideMesataretCikli active_form.Mesataret2
            tmp.Close
        Else
            MbushGrideMesataretCikli active_form.Mesataret2, tmp
            tmp.Close
            'mesatarja e pergjithshme per ciklin
            q_res = " SET CONCAT_NULL_YIELDS_NULL OFF" & _
                    " SELECT " & _
                    " CONVERT(VARCHAR,ROUND(SUM(CONVERT(DECIMAL, NT_VLERESIMI))/COUNT(NT_VLERESIMI), 2)) + '' " & _
                    " AS MESATARJA FROM TBL_NOTA " & _
                    " WHERE NT_VITI_SHKOLLOR = '" + vitishkollor + "'" & _
                    " AND " + strSem + " = 1 AND " & _
                    " NT_MUNGESE_ME = 0 And NT_MUNGESE_PA = 0 " & _
                    " AND CONVERT(INT, NT_KLASA) BETWEEN  " + strMidisLart & _
                    " SET CONCAT_NULL_YIELDS_NULL ON"
            tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
            active_form.Mesataret2.row = active_form.Mesataret2.Rows - 1
            If (tmp.Fields("MESATARJA") = "") Then
                mes = "--"
            Else
                mes = tmp.Fields("MESATARJA")
                If (Len(mes) = 8) Then
                    mes = Mid(mes, 1, 4)
                Else
                    mes = Mid(mes, 1, 5)
                End If
            End If
            active_form.Mesataret2.Text = mes
            tmp.Close
        End If
        
        
        'klasa 10 -12
        q_res = " SET CONCAT_NULL_YIELDS_NULL OFF SELECT LN_EMRI, MESATARJA + '' AS MESATARJA FROM" & _
            " (SELECT NT_EMRI_LENDA, " & _
            " CONVERT(VARCHAR,ROUND(SUM(CONVERT(DECIMAL, NT_VLERESIMI))/COUNT(NT_VLERESIMI), 2)) " & _
            " AS MESATARJA FROM TBL_NOTA " & _
            " WHERE NT_VITI_SHKOLLOR = '" + vitishkollor + "' " & _
            " AND " + strSem + " = 1 AND " & _
            " NT_MUNGESE_ME = 0 And NT_MUNGESE_PA = 0 " & _
            " AND CONVERT(INT, NT_KLASA) BETWEEN 10 AND 12 " & _
            " GROUP BY  NT_EMRI_LENDA) AS A " & _
            " Right Join " & _
            " (SELECT LN_EMRI, LN_NUMRI_ULET FROM TBL_LENDA WHERE LN_VITI_SHKOLLOR = '" + vitishkollor + "' " & _
            " AND " + kushtMesem + " " & _
            " ) AS B ON A.NT_EMRI_LENDA = B.LN_EMRI " & _
            " ORDER BY LN_NUMRI_ULET SET CONCAT_NULL_YIELDS_NULL ON"
        tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
        If (tmp.RecordCount = 0) Then
            PastroGrideMesataretCikli active_form.Mesataret3
            tmp.Close
        Else
            MbushGrideMesataretCikli active_form.Mesataret3, tmp
            tmp.Close
            'mesatarja e pergjithshme per ciklin
            q_res = " SET CONCAT_NULL_YIELDS_NULL OFF" & _
                    " SELECT " & _
                    " CONVERT(VARCHAR,ROUND(SUM(CONVERT(DECIMAL, NT_VLERESIMI))/COUNT(NT_VLERESIMI), 2)) + '' " & _
                    " AS MESATARJA FROM TBL_NOTA " & _
                    " WHERE NT_VITI_SHKOLLOR = '" + vitishkollor + "'" & _
                    " AND " + strSem + " = 1 AND " & _
                    " NT_MUNGESE_ME = 0 And NT_MUNGESE_PA = 0 " & _
                    " AND CONVERT(INT, NT_KLASA) BETWEEN  " + strMidisMesem & _
                    " SET CONCAT_NULL_YIELDS_NULL ON"
            tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
            active_form.Mesataret3.row = active_form.Mesataret3.Rows - 1
            If (tmp.Fields("MESATARJA") = "") Then
                mes = "--"
            Else
                mes = tmp.Fields("MESATARJA")
                If (Len(mes) = 8) Then
                    mes = Mid(mes, 1, 4)
                Else
                    mes = Mid(mes, 1, 5)
                End If
            End If
            active_form.Mesataret3.Text = mes
            tmp.Close
        End If
    Case SHFAQ_PROVIME_ACTION
        If (active_form.optMesme.Value) Then
            strCikli = "1"
        Else
            strCikli = "0"
        End If
        nrVitiShkollor = active_form.cboVitiShkollor.Text
        If strCikli <> "" And nrVitiShkollor <> "" Then
         q_res = "SELECT NTA_EMRI_PROVIMI FROM TBL_NOTAAMZ WHERE  NTA_VITI_SHKOLLOR = '" & nrVitiShkollor & "' AND NTA_CIKLI = " & strCikli
         tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
         Do While Not tmp.EOF

            active_form.listaProvimet.AddItem tmp.Fields("NTA_EMRI_PROVIMI")
            tmp.MoveNext
         Loop
         tmp.Close
      End If
      Case MODIFIKO_LENDE_ACTION
        Dim lendaVjeter, lendaRe As String
        lendaVjeter = CStr(active_form.lista2.List(active_form.lista2.ListIndex))
        vitishkollor = active_form.cboVitiShkollor.Text
        lendaRe = active_form.txtLendaRe.Text
        Klasa = active_form.cboKlasa.Text
        'kontroll nqs ekziston ndonje lende me te njejtin emer si emri i ri i lendes
        q_res = "SELECT * FROM TBL_LENDA WHERE LN_VITI_SHKOLLOR = '" + vitishkollor + "'" & _
                " AND LN_EMRI = '" + CStr(lendaRe) + "'"
        tmp.Open q_res, conn, adOpenDynamic, adLockOptimistic
        'nqs lenda ekziston modifiko kete lende
        If (tmp.RecordCount > 0) Then
            i = MsgBox("Lënda me emrin " + lendaRe + " ekziston e konfiguruar për vitin shkollor " & _
                    "" + vitishkollor + "." & Chr(10) & "Nqs do të kryeni këtë modifikim të gjitha notat e klasës për lëndën " & _
                    "" + lendaVjeter + Chr(10) + "do të transferohen në lëndën " + lendaRe + "." & _
                    "" + Chr(10) + "Jeni të sigurtë që doni të vazhdoni?", vbQuestion + vbOKCancel, "Modifiko lëndë!")
            If i = vbOK Then
                'modifikohet emri i lendes ne tabelen e notave
                q_res = "UPDATE TBL_NOTA SET NT_EMRI_LENDA = '" + lendaRe + "' " & _
                " WHERE NT_VITI_SHKOLLOR = '" + vitishkollor + "' AND NT_EMRI_LENDA = '" & _
                lendaVjeter + "' AND NT_KLASA = '" + Klasa + "'"
                conn.Execute q_res

                'shtohet klasa ne lenden e re
                q_res = "UPDATE TBL_LENDA SET LN_KLASA" + Klasa + " = 1 WHERE " & _
                        "LN_VITI_SHKOLLOR = '" + vitishkollor + "' AND LN_EMRI = '" + lendaRe + "'"
                conn.Execute q_res
                
                'hiqet klasa nga lenda e vjeter
                q_res = "UPDATE TBL_LENDA SET LN_KLASA" + Klasa + " = 0 WHERE " & _
                        "LN_VITI_SHKOLLOR = '" + vitishkollor + "' AND LN_EMRI = '" + lendaVjeter + "'"
                conn.Execute q_res
                
                'kontrollohet nqs lenda e vjeter eshte e konfiguruar per ndonje klase tjeter
                q_res = "SELECT  * FROM TBL_LENDA WHERE LN_EMRI = '" + lendaVjeter + "' AND " & _
                        " LN_VITI_SHKOLLOR = '" + vitishkollor + "'"
                tmp1.Open q_res, conn, adOpenDynamic, adLockOptimistic
                ugjet = False
                For j = 1 To 12
                    If tmp1.Fields("LN_KLASA" + CStr(j)) = True Then
                        ugjet = True
                        Exit For
                    End If
                Next
                tmp1.Close
                'nqs lenda e vjeter nuk gjendet per asnje nga klasat
                'fshihet nga baza
                If (Not ugjet) Then
                    q_res = "DELETE FROM TBL_LENDA WHERE LN_VITI_SHKOLLOR = '" + vitishkollor + "'" & _
                    " AND LN_EMRI = '" + lendaVjeter + "'"
                    conn.Execute q_res
                End If
            
            Else
                tmp.Close
                Exit Sub
            End If
            
        'nqs lenda nuk ekziston modifiko lenden e vjeter
        Else
            'modifikohet emri i lendes ne tabelen e notave
            q_res = "UPDATE TBL_NOTA SET NT_EMRI_LENDA = '" + lendaRe + "' " & _
            " WHERE NT_VITI_SHKOLLOR = '" + vitishkollor + "' AND NT_EMRI_LENDA = '" & _
            lendaVjeter + "' AND NT_KLASA = '" + Klasa + "'"
            conn.Execute q_res

            'hiqet lenda e vjeter nga klasa
            q_res = "UPDATE TBL_LENDA SET LN_KLASA" + Klasa + " = 0 WHERE LN_VITI_SHKOLLOR = '" + vitishkollor + "'" & _
                    " AND LN_EMRI = '" + lendaVjeter + "'"
            conn.Execute q_res

            'kontrollohet nqs lenda e vjeter eshte e konfiguruar per ndonje klase tjeter
            q_res = "SELECT  * FROM TBL_LENDA WHERE LN_EMRI = '" + lendaVjeter + "' AND " & _
                        " LN_VITI_SHKOLLOR = '" + vitishkollor + "'"
            tmp1.Open q_res, conn, adOpenDynamic, adLockOptimistic
            ugjet = False
            For j = 1 To 12
                If tmp1.Fields("LN_KLASA" + CStr(j)) = True Then
                    ugjet = True
                    Exit For
                End If
            Next
            Dim nrUlet, nrMesem As String
            nrUlet = tmp1.Fields("LN_NUMRI_ULET")
            nrMesem = tmp1.Fields("LN_NUMRI_MESME")
            tmp1.Close
            
            'nqs lenda e vjeter nuk gjendet per asnje klase tjeter
            'modifikohet emri i saj
            If (Not ugjet) Then
                q_res = "UPDATE TBL_LENDA SET LN_EMRI = '" + lendaRe + "',LN_KLASA" + Klasa + "= 1 WHERE LN_EMRI = '" + lendaVjeter + "'" & _
                " AND LN_VITI_SHKOLLOR = '" + vitishkollor + "'"
                conn.Execute q_res
                
            'nqs lenda e vjeter gjendet per klasa te tjera
            Else
                
                'krijohet lende e re me emrin e dhene si lenda e re
                q_res = "INSERT INTO TBL_LENDA(LN_VITI_SHKOLLOR, LN_EMRI, LN_KLASA" + CStr(Klasa) + ", LN_NUMRI_ULET, LN_NUMRI_MESME) " & _
                " VALUES('" + CStr(vitishkollor) + "','" + CStr(lendaRe) + "', 1, " + CStr(nrUlet) + ", " + CStr(nrMesem) + ")"
                conn.Execute q_res
            End If
        End If
        tmp.Close
        Case HIDH_EKSEL_ACTION
            
            
   End Select

End Sub

Private Sub MbushGrideMesataretCikli(grid As MSHFlexGrid, tmp As ADODB.Recordset)
    grid.Rows = tmp.RecordCount + 1
    Dim mes As String
    Dim i As Integer
    i = 0
    Do While Not tmp.EOF
        grid.row = i
        grid.col = 0
        grid.CellFontBold = False
        grid.Text = tmp.Fields("LN_EMRI")
        
        
        grid.col = 1
        grid.CellFontBold = False
        If (tmp.Fields("MESATARJA") = "") Then
            mes = "--"
        Else
            mes = tmp.Fields("MESATARJA")
            If (Len(mes) = 8) Then
                mes = Mid(mes, 1, 4)
            Else
                mes = Mid(mes, 1, 5)
            End If
        End If
        grid.Text = mes
        grid.CellAlignment = 3
        If (i Mod 2) Then
            grid.col = 0
            grid.CellBackColor = &HE0E0E0
            grid.col = 1
            grid.CellBackColor = &HE0E0E0
        End If
        grid.RowHeight(i) = 300
        tmp.MoveNext
        i = i + 1
    Loop
    grid.row = tmp.RecordCount
    grid.RowHeight(tmp.RecordCount) = 300
    If (i Mod 2 = 1) Then
        grid.col = 0
        grid.CellBackColor = &HE0E0E0
        grid.col = 1
        grid.CellBackColor = &HE0E0E0
    End If
    grid.col = 0
    grid.CellFontBold = True
    grid.Text = "Mesatarja"
    grid.col = 1
    grid.CellFontBold = True
    grid.CellAlignment = 3
    If (tmp.RecordCount + 1 > 20) Then
        grid.Height = 106.098
        grid.ScrollBars = flexScrollBarVertical
    Else
        grid.Height = (tmp.RecordCount + 1) * CDbl(300 / 56.7)
        grid.ScrollBars = flexScrollBarNone
    End If
End Sub
Private Sub ModifikoMungesa(veprimi As String)
    active_form.emrat.Visible = False
    active_form.datat.Visible = False
    active_form.mungesat.Visible = False
    active_form.amzaEmerMbiemer.Visible = False
    active_form.fsbHorizontali.Visible = False
    active_form.fsbVertikali.Visible = False
    
    Dim nrKlasa, cikli, i, j As Integer
    nrKlasa = CInt(active_form.cboKlasa.Text)
    Dim vitishkollor, q_res, numerAmze As String
    vitishkollor = active_form.cboVitiShkollor.Text
    If ktheCiklin(CStr(nrKlasa), CStr(vitishkollor)) = False Then
        cikli = 0
    Else
        cikli = 1
    End If
    
    Dim conn             As New ADODB.Connection
    With conn
        .Provider = "Microsoft.Access.OLEDB.10.0"
        .Properties("Data Provider").Value = "SQLOLEDB"
        .Properties("Data Source").Value = EmerServer + "\SHM"
        .Properties("User ID").Value = "sa"
        .Properties("Password").Value = "vision"
        .Properties("Initial Catalog").Value = "DProgNec"
        If (conn.state = 1) Then
            .Close
        End If
        .Open
    End With
    Dim affected As Integer
    affected = 0
    For i = 0 To active_form.emrat.Rows - 1
        active_form.emrat.row = i
        active_form.emrat.col = 1
        numerAmze = active_form.emrat.Text
        For j = 0 To active_form.datat.Cols - 1
            active_form.datat.row = 0
            active_form.datat.col = j
            active_form.mungesat.row = i
            Dim kaMungese As Boolean
            kaMungese = False
            active_form.mungesat.col = 2 * j
            If (Trim(active_form.mungesat.Text) <> "") Then
                kaMungese = True
            End If
            active_form.mungesat.col = 2 * j + 1
            'nqs perdoruesi e ka zgjedhur diten per nxenesin
            If (active_form.mungesat.CellPicture = active_form.picChecked And kaMungese = True) Then
                Select Case veprimi
                    Case "MeArsye"
                        q_res = "UPDATE TBL_NOTA SET NT_MUNGESE_ME = 1, NT_MUNGESE_PA = 0" & _
                                " WHERE  (NT_MUNGESE_ME = 1 OR NT_MUNGESE_PA = 1) AND NT_VITI_SHKOLLOR = '" + vitishkollor + "' AND NT_NR_AMZA = '" + numerAmze + "' " & _
                                " AND NT_CIKLI = " + CStr(cikli) + " AND NT_TEKST_DATA = '" + active_form.datat.Text + "'"
                  
                    Case "PaArsye"
                        q_res = "UPDATE TBL_NOTA SET NT_MUNGESE_ME = 0, NT_MUNGESE_PA = 1" & _
                                " WHERE (NT_MUNGESE_ME = 1 OR NT_MUNGESE_PA = 1) AND NT_VITI_SHKOLLOR = '" + vitishkollor + "' AND NT_NR_AMZA = '" + numerAmze + "' " & _
                                " AND NT_CIKLI = " + CStr(cikli) + " AND NT_TEKST_DATA = '" + active_form.datat.Text + "'"
                  
                    Case "FshiMungesa"
                        q_res = "DELETE FROM TBL_NOTA WHERE (NT_MUNGESE_ME = 1 OR NT_MUNGESE_PA = 1) AND " & _
                                " NT_VITI_SHKOLLOR = '" + vitishkollor + "' AND NT_NR_AMZA = '" + numerAmze + "' " & _
                                " AND NT_CIKLI = " + CStr(cikli) + " AND NT_TEKST_DATA = '" + active_form.datat.Text + "'"
                  
                End Select
                conn.Execute q_res
                affected = affected + 1
            End If
        Next
    Next
    conn.Close
    If (affected = 0) Then
        MsgBox "Nuk keni zgjedhur asnjë mungesë për modifikim ose fshirje!", vbExclamation, "Modifiko Mungesa"
    Else
        Select Case veprimi
            Case "MeArsye"
                MsgBox "Mungesat ditore të zgjedhura u modifikuan si me arsye!", vbInformation, "Modifiko Mungesa"
            
            Case "PaArsye"
                MsgBox "Mungesat ditore të zgjedhura u modifikuan si pa arsye!", vbInformation, "Modifiko Mungesa"
            
            Case "FshiMungesa"
                MsgBox "Mungesat ditore të zgjedhura u fshinë!", vbInformation, "Modifiko Mungesa"
        End Select
    End If
    
    
End Sub

Public Sub ex(strSql As String)
   Dim s                As String
   Dim con              As ADODB.Connection

   Set con = New ADODB.Connection
   
   With con
      .Provider = "Microsoft.Access.OLEDB.10.0"
      .Properties("Data Provider").Value = "SQLOLEDB"
      .Properties("Data Source").Value = EmerServer + "\SHM"
      .Properties("User ID").Value = "sa"
      .Properties("Password").Value = "vision"
      .Properties("Initial Catalog").Value = "DProgNec"
      If (con.state = 1) Then
        .Close
      End If
      .Open
   End With
   'con.ConnectionString = getConnectionString()
  ' con.Open
   con.Execute strSql

End Sub
Public Sub get_rec(strSql As String)
   Dim con              As ADODB.Connection
   Set con = New ADODB.Connection
   Set record1 = New ADODB.Recordset

    With con
      .Provider = "Microsoft.Access.OLEDB.10.0"
      .Properties("Data Provider").Value = "SQLOLEDB"
      .Properties("Data Source").Value = EmerServer + "\SHM"
      .Properties("User ID").Value = "sa"
      .Properties("Password").Value = "vision"
      .Properties("Initial Catalog").Value = "DProgNec"
      If (con.state = 1) Then
        .Close
      End If
      .Open
   End With
   
   'con.ConnectionString = getConnectionString()
   'con.Open

   'record1.Source = strSql
   record1.ActiveConnection = con
   record1.Open strSql, con, adOpenDynamic, adLockOptimistic

End Sub

Public Function getConnectionString() As String
   dbPath = App.Path & "\ActiveX\"
   'dbPath = "C:\Documents and Settings\user\Desktop\"
   'getConnectionString = "Provider = Microsoft.Jet.OLEDB.4.0; Persist Security Info = False; Data Source = " & dbPath & dbName & "; User Id = Admin; Password =;"
   
   ' getConnectionString = "workstation id=(LOCAL);packet size=4096;user id=sa;data source=""(LOCAL)"";persist security info=True;initial catalog=ResManager;password=vision"
    getConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Program Files\Microsoft SQL Server\MSSQL\Data\DProgNec.mdf"



End Function


Public Sub get_rec1(strSql As String)
 Dim con As ADODB.Connection
 Set con = New ADODB.Connection
 Set record2 = New ADODB.Recordset
 With con
      .Provider = "Microsoft.Access.OLEDB.10.0"
      .Properties("Data Provider").Value = "SQLOLEDB"
      .Properties("Data Source").Value = EmerServer + "\SHM"
      .Properties("User ID").Value = "sa"
      .Properties("Password").Value = "vision"
      .Properties("Initial Catalog").Value = "DProgNec"
      If (con.state = 1) Then
        .Close
      End If
      .Open
   End With
 
 'con.ConnectionString = getConnectionString()
 'con.Open
 
 record2.Source = strSql
 record2.ActiveConnection = con
 record2.Open
 
End Sub

Private Sub inicializoGrida(grida As MSFlexGrid, numer_rreshti As Integer, numer_shtylla As Integer, rreshtaFikse As Integer, shtyllaFikse As Integer)
    
    grida.Rows = numer_rreshti
    grida.Cols = numer_shtylla
    grida.FixedRows = rreshtaFikse
    grida.FixedCols = shtyllaFikse
    grida.ColWidth(0) = 500
    Dim i As Integer
    For i = 1 To numer_shtylla - 1
        grida.ColWidth(i) = 2400
    Next i
    
    grida.Height = 250 * numer_rreshti + 20
    
    
    
End Sub

Private Sub inicializoGridaAmzaEmesme(gridaKlasat As MSHFlexGrid, gridaAmza As MSHFlexGrid, gridaLendet As MSHFlexGrid, gridaLabel As MSHFlexGrid, numer_shtylla As Integer)

   gridaKlasat.Height = 2450
   gridaKlasat.Width = 2000
   gridaKlasat.Rows = 8
   gridaKlasat.Cols = 2
   gridaKlasat.FixedCols = 0
   gridaKlasat.FixedRows = 0
   gridaKlasat.ColWidth(0) = 800
   gridaKlasat.ColWidth(1) = 1200
   
   gridaLabel.Rows = 1
   gridaLabel.Cols = 2
   gridaLabel.FixedCols = 0
   gridaLabel.FixedRows = 0
   gridaLabel.Height = 300
   gridaLabel.Width = 2000
   gridaLabel.ColWidth(0) = 800
   gridaLabel.ColWidth(1) = 1200
   gridaLabel.BackColor = &H8000000F
   gridaLendet.Rows = 1
   gridaLendet.Cols = numer_shtylla + 1
   gridaLendet.Height = 300
   gridaLendet.Width = 12000
   gridaLendet.FixedRows = 0
   gridaLendet.FixedCols = 0
   gridaLendet.RowHeight(0) = 300
   gridaLabel.RowHeight(0) = 300
   
   gridaAmza.Height = 2450
   gridaAmza.Width = 15000
   gridaAmza.Rows = 8
   gridaAmza.Cols = numer_shtylla + 1
   
   gridaAmza.FixedRows = 0
   gridaAmza.FixedCols = 0
   
   Dim i                As Integer
   For i = 0 To numer_shtylla - 1
      gridaAmza.ColWidth(i) = 1500
      gridaLendet.ColWidth(i) = 1500
   Next i
   For i = 0 To 7
     gridaKlasat.RowHeight(i) = 300
     
   Next i
   
   For i = 0 To 7
     
     gridaAmza.RowHeight(i) = 300
   Next i
   gridaKlasat.row = 0
   
   gridaKlasat.col = 0
   gridaKlasat.row = 0
   gridaKlasat.Text = "9"
   gridaKlasat.row = 1
   gridaKlasat.Text = "9"
   gridaKlasat.row = 2
   gridaKlasat.Text = "10"
   gridaKlasat.row = 3
   gridaKlasat.Text = "10"
   gridaKlasat.row = 4
   gridaKlasat.Text = "11"
   gridaKlasat.row = 5
   gridaKlasat.Text = "11"
   gridaKlasat.row = 6
   gridaKlasat.Text = "12"
   gridaKlasat.row = 7
   gridaKlasat.Text = "12"
   
   
   gridaLendet.col = numer_shtylla
   gridaLendet.ColWidth(numer_shtylla) = 3000
   gridaLendet.row = 0
   gridaLendet.Text = "Vrejtje"
   
   gridaAmza.col = numer_shtylla
   gridaAmza.ColWidth(numer_shtylla) = 3000
   
   
   

   Dim j                As Integer
   For j = 0 To gridaAmza.Rows - 1
      If j Mod 2 = 1 Then
       For i = 0 To gridaAmza.Cols - 1
         gridaAmza.row = j
         gridaAmza.col = i
         gridaAmza.CellBackColor = &HE0E0E0
         
        
       Next i
      End If
   Next j
   
   For i = 0 To gridaLendet.Cols - 1
      gridaLendet.row = 0
      gridaLendet.col = i
      gridaLendet.CellBackColor = &H8000000F
   Next i
   
   For j = 0 To gridaKlasat.Rows - 1
      If j Mod 2 = 1 Then
       For i = 0 To gridaKlasat.Cols - 1
         
         gridaKlasat.row = j
         gridaKlasat.col = i
         gridaKlasat.CellBackColor = &HE0E0E0
         
         
       Next i
      End If
   Next j
   
   For j = 0 To gridaKlasat.Rows - 1
     For i = 0 To gridaKlasat.Cols - 1
         
         gridaKlasat.row = j
         gridaKlasat.col = i
         gridaKlasat.CellFontBold = True
         gridaKlasat.CellAlignment = 1
         
         
       Next i
     Next j
     
     For j = 0 To gridaAmza.Rows - 1
        For i = 0 To gridaAmza.Cols - 1
         gridaAmza.row = j
         gridaAmza.col = i
         gridaAmza.CellFontBold = True
         gridaAmza.CellAlignment = 1
         
        
       Next i
     Next j
   
   gridaKlasat.ScrollBars = flexScrollBarNone
   gridaAmza.ScrollBars = flexScrollBarNone
   gridaLendet.ScrollBars = flexScrollBarNone
   gridaLabel.ScrollBars = flexScrollBarNone
   gridaAmza.ColWidth(numer_shtylla) = 3000
   gridaLendet.ColWidth(numer_shtylla) = 3000
End Sub

Private Sub inicializoGridaAmza9Vjecare(gridaKlasat As MSHFlexGrid, gridaAmza As MSHFlexGrid, vitishkollor As String, Klasa As Integer, cikli As Integer)
Dim vitFillimi As Integer
Dim i As Integer
vitFillimi = CInt(Mid(vitishkollor, 1, 4))
If (cikli = 0) Then
    'shkolla 8/9 vjecare
    'nqs nxenesi e ka mbaruar shkollen para 2000 eshte ne 8vjecare
    If (vitFillimi < 2000) Then
        Exit Sub
    'Nqs eshte ne vitet shkollore midis 2000-2001 dhe 2006-2007
    ElseIf (vitFillimi < 2007) Then
        'sistem 8vjecar
        If (CInt(Mid(vitishkollor, 9, 1)) < Klasa) Then
            Exit Sub
        Else
        'sistem 9vjecar modifiko gride me 9 klasa
            gridaKlasat.Rows = 18
            gridaAmza.Rows = 18
            
            gridaKlasat.RowHeight(16) = 300
            gridaKlasat.RowHeight(17) = 300
            gridaAmza.RowHeight(16) = 300
            gridaAmza.RowHeight(17) = 300
            
            gridaKlasat.col = 0
            gridaKlasat.row = 16
            gridaKlasat.CellFontBold = True
            gridaKlasat.CellAlignment = 1
            gridaKlasat.Text = "9"
            
            gridaKlasat.row = 17
            gridaKlasat.CellFontBold = True
            gridaKlasat.CellAlignment = 1
            gridaKlasat.Text = "9"
            gridaKlasat.CellBackColor = &HE0E0E0
            
            gridaKlasat.col = 1
            gridaKlasat.CellBackColor = &HE0E0E0
            gridaKlasat.CellFontBold = True
            gridaKlasat.CellAlignment = 1
            
            gridaKlasat.row = 16
            gridaKlasat.CellFontBold = True
            gridaKlasat.CellAlignment = 1
            
            gridaAmza.row = 17
            For i = 0 To gridaAmza.Cols - 1
                gridaAmza.col = i
                gridaAmza.CellBackColor = &HE0E0E0
            Next
            'gridaKlasat.Height = 2750
            'gridaAmza.Height = 2750
        End If
    'nqs eshte ne vitin shkollor 2007-2008 ose me tej
    'e kap sistemi 9vjecar modifiko gride me 9 klasa
    Else
        gridaKlasat.Rows = 18
        gridaAmza.Rows = 18
        
        gridaKlasat.RowHeight(16) = 300
        gridaKlasat.RowHeight(17) = 300
        gridaAmza.RowHeight(16) = 300
        gridaAmza.RowHeight(17) = 300
        
        gridaKlasat.col = 0
        gridaKlasat.row = 16
        gridaKlasat.CellFontBold = True
        gridaKlasat.CellAlignment = 1
        gridaKlasat.Text = "9"
        
        gridaKlasat.row = 17
        gridaKlasat.CellFontBold = True
        gridaKlasat.CellAlignment = 1
        gridaKlasat.Text = "9"
        gridaKlasat.CellBackColor = &HE0E0E0
        
        gridaKlasat.col = 1
        gridaKlasat.CellBackColor = &HE0E0E0
        gridaKlasat.CellFontBold = True
        gridaKlasat.CellAlignment = 1
        
        gridaKlasat.row = 16
        gridaKlasat.CellFontBold = True
        gridaKlasat.CellAlignment = 1
        
        gridaAmza.row = 17
        For i = 0 To gridaAmza.Cols - 1
            gridaAmza.col = i
            gridaAmza.CellBackColor = &HE0E0E0
        Next
        'gridaKlasat.Height = 2750
        'gridaAmza.Height = 2750
    End If
Else
    'shkolla e mesme
    'Per vitin shkollor 2008-2009 e tutje shkolla e mesme perbehet vetem nga 3 klase
    If (vitFillimi >= 2008) Then
        gridaKlasat.Rows = 6
        gridaAmza.Rows = 6
        
        gridaKlasat.col = 0
        gridaKlasat.row = 0
        gridaKlasat.Text = "10"
        gridaKlasat.row = 1
        gridaKlasat.Text = "10"
        
        gridaKlasat.row = 2
        gridaKlasat.Text = "11"
        gridaKlasat.row = 3
        gridaKlasat.Text = "11"
        
        gridaKlasat.row = 4
        gridaKlasat.Text = "12"
        gridaKlasat.row = 5
        gridaKlasat.Text = "12"
    'per vitet shkollor 2007-2008 shkolla e mesme perbehet nga 4 klase
    Else
        Exit Sub
    End If
End If
End Sub

Private Sub inicializoGridaTetevjecare(gridaKlasat As MSHFlexGrid, gridaAmza As MSHFlexGrid, gridaLendet As MSHFlexGrid, gridaLabel As MSHFlexGrid, numer_shtylla As Integer)
   
   gridaKlasat.Height = 2450
   gridaKlasat.Width = 2000
   gridaKlasat.Rows = 16
   gridaKlasat.Cols = 2
   gridaKlasat.FixedCols = 0
   gridaKlasat.FixedRows = 0
   gridaKlasat.ColWidth(0) = 800
   gridaKlasat.ColWidth(1) = 1200
   
   gridaLabel.Rows = 1
   gridaLabel.Cols = 2
   gridaLabel.FixedCols = 0
   gridaLabel.FixedRows = 0
   gridaLabel.Height = 300
   gridaLabel.Width = 2000
   gridaLabel.ColWidth(0) = 800
   gridaLabel.ColWidth(1) = 1200
   gridaLabel.BackColor = &H8000000F
   
   gridaLendet.Rows = 1
   gridaLendet.Cols = numer_shtylla + 1
   gridaLendet.Height = 300
   gridaLendet.Width = 14575
   gridaLendet.FixedRows = 0
   gridaLendet.FixedCols = 0
   gridaLendet.RowHeight(0) = 300
   
   gridaAmza.Height = 2450
   gridaAmza.Width = 12000
   gridaAmza.Rows = 16
   gridaAmza.Cols = numer_shtylla + 1
   
   gridaAmza.FixedRows = 0
   gridaAmza.FixedCols = 0
   
   Dim i                As Integer
   For i = 0 To numer_shtylla - 1
      gridaAmza.ColWidth(i) = 1500
      gridaLendet.ColWidth(i) = 1500
   Next i
   For i = 0 To 15
     gridaKlasat.RowHeight(i) = 300
     
   Next i
   
   For i = 0 To 15
     
     gridaAmza.RowHeight(i) = 300
   Next i
   
   gridaLabel.RowHeight(0) = 300
   
   gridaKlasat.row = 0
   
   gridaKlasat.col = 0
   gridaKlasat.row = 0
   gridaKlasat.Text = "1"
   gridaKlasat.row = 1
   gridaKlasat.Text = "1"
   gridaKlasat.row = 2
   gridaKlasat.Text = "2"
   gridaKlasat.row = 3
   gridaKlasat.Text = "2"
   gridaKlasat.row = 4
   gridaKlasat.Text = "3"
   gridaKlasat.row = 5
   gridaKlasat.Text = "3"
   gridaKlasat.row = 6
   gridaKlasat.Text = "4"
   gridaKlasat.row = 7
   gridaKlasat.Text = "4"
   gridaKlasat.row = 8
   gridaKlasat.Text = "5"
   gridaKlasat.row = 9
   gridaKlasat.Text = "5"
   gridaKlasat.row = 10
   gridaKlasat.Text = "6"
   gridaKlasat.row = 11
   gridaKlasat.Text = "6"
   gridaKlasat.row = 12
   gridaKlasat.Text = "7"
   gridaKlasat.row = 13
   gridaKlasat.Text = "7"
   gridaKlasat.row = 14
   gridaKlasat.Text = "8"
   gridaKlasat.row = 15
   gridaKlasat.Text = "8"
   
   gridaLendet.col = numer_shtylla
   gridaLendet.ColWidth(numer_shtylla) = 3000
   gridaLendet.row = 0
   gridaLendet.Text = "Vrejtje"
   
   gridaAmza.col = numer_shtylla
   gridaAmza.ColWidth(numer_shtylla) = 3000
   
   
   

   Dim j                As Integer
   For j = 0 To gridaAmza.Rows - 1
      If j Mod 2 = 1 Then
       For i = 0 To gridaAmza.Cols - 1
         gridaAmza.row = j
         gridaAmza.col = i
         gridaAmza.CellBackColor = &HE0E0E0
         
        
       Next i
      End If
   Next j
   
   For i = 0 To gridaLendet.Cols - 1
      gridaLendet.row = 0
      gridaLendet.col = i
      gridaLendet.CellBackColor = &H8000000F
   Next i
   
   For j = 0 To gridaKlasat.Rows - 1
      If j Mod 2 = 1 Then
       For i = 0 To gridaKlasat.Cols - 1
         
         gridaKlasat.row = j
         gridaKlasat.col = i
         gridaKlasat.CellBackColor = &HE0E0E0
         
         
       Next i
      End If
   Next j
   
   For j = 0 To gridaKlasat.Rows - 1
     For i = 0 To gridaKlasat.Cols - 1
         
         gridaKlasat.row = j
         gridaKlasat.col = i
         gridaKlasat.CellFontBold = True
         gridaKlasat.CellAlignment = 1
         
         
       Next i
     Next j
     
     For j = 0 To gridaAmza.Rows - 1
        For i = 0 To gridaAmza.Cols - 1
         gridaAmza.row = j
         gridaAmza.col = i
         gridaAmza.CellFontBold = True
         gridaAmza.CellAlignment = 1
         
        
       Next i
     Next j
   
   gridaKlasat.ScrollBars = flexScrollBarNone
   gridaAmza.ScrollBars = flexScrollBarNone
   gridaLendet.ScrollBars = flexScrollBarNone
   gridaLabel.ScrollBars = flexScrollBarNone
   gridaAmza.ColWidth(numer_shtylla) = 3000
   gridaLendet.ColWidth(numer_shtylla) = 3000
End Sub


Private Function gjendetNeDatabase(lenda As String, nr As String, conn As ADODB.Connection) As Boolean
    
    Dim ugjet As Boolean
    ugjet = False
    Dim rekset As New ADODB.Recordset
    Dim q_res As String
    q_res = " SELECT LN_EMRI, LN_VITI_SHKOLLOR FROM TBL_LENDA"
    rekset.Open q_res, conn, adOpenDynamic, adLockOptimistic
    Do While Not rekset.EOF
        If (rekset.Fields("LN_EMRI") = lenda) And (rekset.Fields("LN_VITI_SHKOLLOR") = nr) Then
            ugjet = True
            rekset.MoveLast
        Else
            rekset.MoveNext
        End If
    Loop
        
    gjendetNeDatabase = ugjet
    
    
    
End Function

Private Function ktheDistancen(emri As String) As Integer
    Dim nr As Integer
    nr = Len(emri)
    ktheDistancen = 40 - nr
End Function

Private Function ktheKlase(rreshti As String) As Double
   Dim Klasa            As String
   Dim s3 As String
   s3 = Trim(rreshti)
   Dim s2               As String
   s2 = ""
   Dim ugjet            As Boolean
   Dim i, j             As Integer
   ugjet = False
   i = 1
   Do While Not ugjet
      If Mid(s3, i, 1) <> " " Then
         s2 = s2 & Mid(s3, i, 1)
         i = i + 1
      Else
         ugjet = True
      End If
   Loop

   Klasa = ""
   ugjet = False
   j = 1
   Do While Not ugjet
      If Mid(s2, j, 1) <> "-" Then
         Klasa = Klasa & Mid(s2, j, 1)
         j = j + 1
      Else
         ugjet = True
      End If
   Loop

   ktheKlase = Klasa
End Function


Private Function ktheIndeksi(rreshti As String) As String
   Dim ugjet As Boolean
   Dim s3               As String
   s3 = Trim(rreshti)
   Dim Klasa            As String
   Dim indeksi          As String
   Dim s2               As String
   Dim i, j             As Integer
   s2 = ""
   ugjet = False
   i = 1
   Do While Not ugjet
      If Mid(s3, i, 1) <> " " Then
         s2 = s2 & Mid(s3, i, 1)
         i = i + 1
      Else
         ugjet = True
      End If
   Loop

   Klasa = ""
   ugjet = False
   j = 1
   Do While Not ugjet
      If Mid(s2, j, 1) <> "-" Then
         Klasa = Klasa & Mid(s2, j, 1)
         j = j + 1
      Else
         ugjet = True
      End If
   Loop
   j = j + 1
   indeksi = Mid(s2, j)
   ktheIndeksi = indeksi
End Function

Private Function kontrollo_notat(grida As Cgrid, rreshti As Integer, shtylla As Integer) As Boolean
    Dim nota As String
    Dim ugjet As Boolean
    ugjet = False
    Dim i As Integer
    Dim j As Integer
    i = 1
    
    Do While i <= rreshti And Not ugjet
        j = 1
        Do While j <= shtylla And Not ugjet
            nota = grida.Text(i, j)
            If Not (nota = "" Or nota = "4" Or nota = "5" Or nota = "6" Or nota = "7" Or nota = "8" Or nota = "9" Or nota = "10") Then
                ugjet = True
                nota_rreshti = i
                nota_shtylla = j
            End If
            j = j + 1
        Loop
        i = i + 1
    Loop
    
    kontrollo_notat = ugjet
    
            
            
End Function


Private Function gjendetNeListe(emri As String, liste As ListBox)
    Dim j As Integer
    Dim ugjet As Boolean
    ugjet = False
    j = 0
    Do While (j <= liste.ListCount - 1) And (ugjet = False)
        If LCase(emri) = LCase(liste.List(j)) Then
            ugjet = True
        Else
            j = j + 1
        End If
    Loop
    
    gjendetNeListe = ugjet
        
End Function

Private Sub inicializo_grida_provime(grida As MSFlexGrid, shtylla As Integer)
    
    Dim i As Integer
    grida.Cols = shtylla
    grida.Rows = 2
    grida.RowHeight(0) = 300
    grida.RowHeight(1) = 300
    For i = 0 To shtylla - 1
        grida.ColWidth(i) = 1500
    Next i
    grida.Height = 2 * 300 + 30
    grida.Width = shtylla * 1500 + 50
End Sub

Private Function konvertoDaten(data As String) As String
    Dim dataDt As Date
    dataDt = CDate(data)
    Dim dataMbrapsht As String
    Dim i As String
    dataMbrapsht = ""
    Dim dita As String
    Dim muaji As String
    Dim viti As String
    dita = DateTime.Day(dataDt)
    muaji = DateTime.Month(dataDt)
    viti = DateTime.Year(dataDt)
    'viti = Mid(data, 7, 4)
    If (Len(dita) = 1) Then
        dita = "0" + dita
    End If
    
    If (Len(muaji) = 1) Then
        muaji = "0" + muaji
    End If
    dataMbrapsht = viti & "/" & muaji & "/" & dita
    konvertoDaten = dataMbrapsht
End Function

Private Function DataTekst(data As String) As String
    Dim dataDt As Date
    dataDt = CDate(data)
    Dim dataMbrapsht As String
    Dim i As String
    dataMbrapsht = ""
    Dim dita As String
    Dim muaji As String
    Dim viti As String
    dita = DateTime.Day(dataDt)
    muaji = DateTime.Month(dataDt)
    viti = DateTime.Year(dataDt)
    'viti = Mid(data, 7, 4)
    If (Len(dita) = 1) Then
        dita = "0" + dita
    End If
    
    If (Len(muaji) = 1) Then
        muaji = "0" + muaji
    End If
    dataMbrapsht = dita & "/" & muaji & "/" & viti
    DataTekst = dataMbrapsht
End Function

Private Function DataToTekst(dataDt As Date) As String
    Dim dataMbrapsht As String
    Dim i As String
    dataMbrapsht = ""
    Dim dita As String
    Dim muaji As String
    Dim viti As String
    dita = DateTime.Day(dataDt)
    muaji = DateTime.Month(dataDt)
    viti = DateTime.Year(dataDt)
    'viti = Mid(data, 7, 4)
    If (Len(dita) = 1) Then
        dita = "0" + dita
    End If
    
    If (Len(muaji) = 1) Then
        muaji = "0" + muaji
    End If
    dataMbrapsht = dita & "/" & muaji & "/" & viti
    DataToTekst = dataMbrapsht
End Function

