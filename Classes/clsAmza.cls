VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAmza"
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
Private mvarEmri As String 'local copy
Private mvarNo As String 'local copy
Private mvarKlasa As String 'local copy
Private mvarKlaseIndex As String 'local copy
Private mvarMbiemri As String 'local copy
Private mvarAtesia As String 'local copy
Private mvarMemesia As String 'local copy
Private mvarSeksi As String 'local copy
Private mvarVendlindja As String 'local copy
Private mvarDatelindja As Date 'local copy
Private mvarVrejtje As String 'local copy
Private mvarNotes As String 'local copy
Private mvarErrorHandler As New clsErrorHandler 'local copy


Public Property Let Notes(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Notes = 5
    mvarNotes = vData
End Property


Public Property Get Notes() As String
   'used when retrieving value of a property, on the right side of an assignment.
   'Syntax: Debug.Print X.Notes
   Notes = mvarNotes
End Property




Public Property Set ErrorHandler(ByVal vData As clsErrorHandler)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.ErrorHandler = Form1
    Set mvarErrorHandler = vData
End Property


Public Property Get ErrorHandler() As clsErrorHandler
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ErrorHandler
    Set ErrorHandler = mvarErrorHandler
End Property

Public Property Let Vrejtje(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Vrejtje = 5
    mvarVrejtje = vData
End Property


Public Property Get Vrejtje() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Vrejtje
    Vrejtje = mvarVrejtje
End Property



Public Property Let Datelindja(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Datelindja = 5
    mvarDatelindja = vData
End Property


Public Property Get Datelindja() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Datelindja
    Datelindja = mvarDatelindja
End Property



Public Property Let Vendlindja(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Vendlindja = 5
    mvarVendlindja = vData
End Property


Public Property Get Vendlindja() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Vendlindja
    Vendlindja = mvarVendlindja
End Property



Public Property Let Seksi(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Seksi = 5
    mvarSeksi = vData
End Property


Public Property Get Seksi() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Seksi
    Seksi = mvarSeksi
End Property



Public Property Let Memesia(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Memesia = 5
    mvarMemesia = vData
End Property


Public Property Get Memesia() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Memesia
    Memesia = mvarMemesia
End Property



Public Property Let Atesia(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Atesia = 5
    mvarAtesia = vData
End Property


Public Property Get Atesia() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Atesia
    Atesia = mvarAtesia
End Property



Public Property Let Mbiemri(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Mbiemri = 5
    mvarMbiemri = vData
End Property


Public Property Get Mbiemri() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Mbiemri
    Mbiemri = mvarMbiemri
End Property



Public Property Let KlaseIndex(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.KlaseIndex = 5
    mvarKlaseIndex = vData
End Property


Public Property Get KlaseIndex() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.KlaseIndex
    KlaseIndex = mvarKlaseIndex
End Property



Public Property Let Klasa(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Klasa = 5
    mvarKlasa = vData
End Property


Public Property Get Klasa() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Klasa
    Klasa = mvarKlasa
End Property



Public Property Let No(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.No = 5
    mvarNo = vData
End Property


Public Property Get No() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.No
    No = mvarNo
End Property



Public Property Let Emri(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Emri = 5
    mvarEmri = vData
End Property


Public Property Get Emri() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Emri
    Emri = mvarEmri
End Property


Public Function Clone() As clsAmza
   Dim clonedObject     As New clsAmza

   With clonedObject
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
      Set .ErrorHandler = Me.ErrorHandler.Clone
   End With

   Set Clone = clonedObject

   Set clonedObject = Nothing
End Function

Private Sub Class_Terminate()
  Set ErrorHandler = Nothing
End Sub
