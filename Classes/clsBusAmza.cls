VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBusAmza"
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
' * Time             : 20:51
' * Module Name      : clsAmza
' * Module Filename  : clsAmza.cls
' * *******************************************************************
' * Comments         :
' *
' *
' * *******************************************************************


'local variable(s) to hold property value(s)
Private mvarEmri As String
Private mvarNo As String
Private mvarKlasa As String
Private mvarKlaseIndex As String
Private mvarMbiemri As String
Private mvarAtesia As String
Private mvarMemesia As String
Private mvarSeksi As String
Private mvarVendlindja As String
Private mvarDatelindja As Date
Private mvarVrejtje As String
Private mvarNotes As String
Private mvarErrorHandler As New clsErrorHandler
Private mvarKey As String




Public Property Let Key(ByVal vData As String)
    mvarKey = vData
End Property


Public Property Get Key() As String
   Key = mvarKey
End Property


Public Property Let Notes(ByVal vData As String)
    mvarNotes = vData
End Property


Public Property Get Notes() As String
   Notes = mvarNotes
End Property


Public Property Set ErrorHandler(ByVal vData As clsErrorHandler)
    Set mvarErrorHandler = vData
End Property


Public Property Get ErrorHandler() As clsErrorHandler
   Set ErrorHandler = mvarErrorHandler
End Property


Public Property Let Vrejtje(ByVal vData As String)
    mvarVrejtje = vData
End Property


Public Property Get Vrejtje() As String
   Vrejtje = mvarVrejtje
End Property


Public Property Let Datelindja(ByVal vData As Date)
    mvarDatelindja = vData
End Property


Public Property Get Datelindja() As Date
   Datelindja = mvarDatelindja
End Property


Public Property Let Vendlindja(ByVal vData As String)
    mvarVendlindja = vData
End Property


Public Property Get Vendlindja() As String
   Vendlindja = mvarVendlindja
End Property


Public Property Let Seksi(ByVal vData As String)
    mvarSeksi = vData
End Property


Public Property Get Seksi() As String
   Seksi = mvarSeksi
End Property


Public Property Let Memesia(ByVal vData As String)
    mvarMemesia = vData
End Property


Public Property Get Memesia() As String
   Memesia = mvarMemesia
End Property


Public Property Let Atesia(ByVal vData As String)
    mvarAtesia = vData
End Property


Public Property Get Atesia() As String
   Atesia = mvarAtesia
End Property


Public Property Let Mbiemri(ByVal vData As String)
    mvarMbiemri = vData
End Property


Public Property Get Mbiemri() As String
   Mbiemri = mvarMbiemri
End Property


Public Property Let KlaseIndex(ByVal vData As String)
    mvarKlaseIndex = vData
End Property


Public Property Get KlaseIndex() As String
   KlaseIndex = mvarKlaseIndex
End Property


Public Property Let Klasa(ByVal vData As String)
    mvarKlasa = vData
End Property


Public Property Get Klasa() As String
   Klasa = mvarKlasa
End Property


Public Property Let No(ByVal vData As String)
    mvarNo = vData
End Property


Public Property Get No() As String
   No = mvarNo
End Property


Public Property Let Emri(ByVal vData As String)
    mvarEmri = vData
End Property


Public Property Get Emri() As String
   Emri = mvarEmri
End Property


Private Sub Class_Terminate()
   Set ErrorHandler = Nothing
End Sub


Public Function Clone() As clsBusAmza
   Dim ClonedObject     As New clsBusAmza

   With ClonedObject
      .Emri = Me.Emri
      .No = Me.No
      .Klasa = Me.Klasa
      .KlaseIndex = Me.KlaseIndex
      .Mbiemri = Me.Mbiemri
      .Atesia = Me.Atesia
      .Memesia = Me.Memesia
      .Seksi = Me.Seksi
      .Vendlindja = Me.Vendlindja
      .Datelindja = Me.Datelindja
      .Vrejtje = Me.Vrejtje
      .Key = Me.Key
      Set .ErrorHandler = Me.ErrorHandler.Clone
   End With

   Set Clone = ClonedObject

   Set ClonedObject = Nothing
End Function


Public Function Load(rs As ADODB.Recordset) As clsBusAmza

   If Not rs.EOF Then
      With Me
         .Key = rs("AMZ_ID")
         .No = rs("AMZ_NO")
         .Klasa = rs("AMZ_KLASA")
         .KlaseIndex = rs("AMZ_KLASEINDEX")
         .Emri = rs("AMZ_EMRI")
         .Mbiemri = rs("AMZ_MBIEMRI")
         .Atesia = rs("AMZ_ATESIA")
         .Seksi = rs("AMZ_SEKSI")
         .Memesia = rs("AMZ_MEMESIA")
         .Datelindja = rs("AMZ_DATELINDJA")
         .Vendlindja = rs("AMZ_VENDLINDJA")
      End With
   End If

   Set Load = Me

End Function


Public Function LoadV(vParams() As Variant) As clsBusAmza

   If Not IsArray(vParams) Then
      Exit Function
   End If

   With objAmza
      .Key = vParams(0, 0)
      .No = vParams(1, 0)
      .Klasa = vParams(2, 0)
      .Emri = vParams(3, 0)
      .Mbiemri = vParams(4, 0)
      .Atesia = vParams(5, 0)
      .Memesia = vParams(6, 0)
      .Seksi = vParams(7, 0)
      .Vendlindja = vParams(8, 0)
      .Datelindja = vParams(9, 0)
      .Vrejtje = vParams(10, 0)
      .KlaseIndex = vParams(11, 0)
   End With

   Set LoadV = Me

End Function

Public Function GetArray() As Variant

   With Me
      GetArray = Array(.Key, .No, .Klasa, .Emri, .Mbiemri, .Atesia, .Memesia, .Seksi, .Vendlindja, .Datelindja, .Vrejtje, .KlaseIndex)
   End With

End Function


Public Function GetGUIAmza_HedhjeGjeneralitete() As clsBusAmza

   With active_form
      Me.Atesia = Trim(.txtAtesia.Text)
      Me.Datelindja = .DTPicker1.Value
      Me.Emri = Trim(.txtEmri.Text)
      Me.Klasa = Trim(.txtKlasa.Text)
      Me.KlaseIndex = Trim(.txtKlasa.Text)
      Me.Mbiemri = Trim(.txtMbiemri.Text)
      Me.Memesia = Trim(.txtMemesia.Text)
      Me.No = Trim(.txtAmzaNo.Text)
      Me.Seksi = Trim(.txtSeksi.Text)
      Me.Vendlindja = Trim(.txtVendlindja.Text)
      Me.Vrejtje = Trim(.rtbShenime.Text)
   End With

   Set GetGUIAmza_HedhjeGjeneralitete = Me
   
End Function


Public Function GetGUIAmza_() As clsBusAmza
   ' to do
   With active_form

   End With

   Set GetGUIAmza_ = Me

End Function
