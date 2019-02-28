VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNota"
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
' * Time             : 20:54
' * Module Name      : clsNota
' * Module Filename  : clsNota.cls
' * *******************************************************************
' * Comments         :
' *
' *
' * *******************************************************************



'local variable(s) to hold property value(s)
Private mvarNotat() As String 'local copy
Private mvarData As Date 'local copy

Public Enum NOTE_TIP_ENUM
    ntMomentaleI = 1
    ntMomentaleII = 2
    ntSemestraleI = 3
    ntSemestraleII = 4
    ntVjetore = 5
    ntMungese = 6
    ntDetyreKontrolli = 7
End Enum

Private mvarKlasa As String 'local copy
Private mvarKlaseIndex As String 'local copy
Private mvarLenda As String 'local copy
Private mvarErrorHandler As New clsErrorHandler 'local copy

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



Public Property Let Lenda(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Lenda = 5
    mvarLenda = vData
End Property


Public Property Get Lenda() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Lenda
    Lenda = mvarLenda
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



Public Property Let Data(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Data = 5
    mvarData = vData
End Property


Public Property Get Data() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Data
    Data = mvarData
End Property



Public Property Let Notat(ByRef vData() As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Vleresimi = 5
    mvarNotat = vData
End Property


Public Property Get Notat() As String()
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Vleresimi
    Notat = mvarNotat
End Property



Private Sub Class_Terminate()
  Set ErrorHandler = Nothing
End Sub


Public Function Clone() As clsNota
   Dim clonedObject     As New clsNota

   With clonedObject
      .Notat = Me.Notat
      .Data = Me.Data
      .Klasa = Me.Klasa
      .KlaseIndex = Me.KlaseIndex
      .Lenda = Me.Lenda
      Set .ErrorHandler = Me.ErrorHandler.Clone
   End With

   Set Clone = clonedObject

   Set clonedObject = Nothing
End Function
