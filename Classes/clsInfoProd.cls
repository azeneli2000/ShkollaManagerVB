VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInfoProd"
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
' * Time             : 21:28
' * Module Name      : clsInfoProd
' * Module Filename  : clsInfoProd.cls
' * *******************************************************************
' * Comments         :
' *
' *
' * *******************************************************************


'local variable(s) to hold property value(s)
Private mvarErrorHandler As New clsErrorHandler 'local copy
Private mvarAmzaData As New clsAmza 'local copy
Private mvarLendet As New colLenda 'local copy
Private mvarNotat As New colNota 'local copy

Public Sub getMestareNota()
   
End Sub

Public Property Set Notat(ByVal vData As colNota)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Notat = Form1
    Set mvarNotat = vData
End Property


Public Property Get Notat() As colNota
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Notat
    Set Notat = mvarNotat
End Property



Public Property Set Lendet(ByVal vData As colLenda)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Lendet = Form1
    Set mvarLendet = vData
End Property


Public Property Get Lendet() As colLenda
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Lendet
    Set Lendet = mvarLendet
End Property



Public Property Set AmzaData(ByVal vData As clsAmza)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.AmzaData = Form1
    Set mvarAmzaData = vData
End Property


Public Property Get AmzaData() As clsAmza
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AmzaData
    Set AmzaData = mvarAmzaData
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

Private Sub Class_Terminate()
   Set ErrorHandler = Nothing
   Set mvarAmzaData = Nothing
   Set mvarLendet = Nothing
   Set mvarNotat = Nothing
End Sub

Public Sub Clear()
   Set mvarErrorHandler = New clsErrorHandler
   Set mvarAmzaData = New clsAmza
   Set mvarLendet = New clsLenda
   Set mvarNotat = New clsNota
End Sub

Public Function Clone() As clsInfoProd
   Dim objClone         As New clsInfoProd

   With objClone
      Set .AmzaData = Me.AmzaData.Clone
      Set .ErrorHandler = Me.ErrorHandler.Clone
      Set .Lendet = Me.Lendet.Clone
      Set .Notat = Me.Notat.Clone
   End With

   Set Clone = objClone

   Set objClone = Nothing
End Function
