VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

' * Author           :
' * Web Site         :
' * E-Mail           :
' * Date             : 11/22/2004
' * Time             : 15:44
' * Module Name      : clsLenda
' * Module Filename  : clsLenda.cls
' * *******************************************************************
' * Comments         :
' *
' *
' * *******************************************************************

'local variable(s) to hold property value(s)
Private mvarEmri As String 'local copy
Private mvarVitiShkollor As String 'local copy
Private mvarKlasa As Variant 'local copy
Public Property Let Klasa(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Klasa = 5
    mvarKlasa = vData
End Property


Public Property Set Klasa(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Klasa = Form1
    Set mvarKlasa = vData
End Property


Public Property Get Klasa() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Klasa
    If IsObject(mvarKlasa) Then
        Set Klasa = mvarKlasa
    Else
        Klasa = mvarKlasa
    End If
End Property


Public Property Let VitiShkollor(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.VitiShkollor = 5
    mvarVitiShkollor = vData
End Property


Public Property Get VitiShkollor() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.VitiShkollor
    VitiShkollor = mvarVitiShkollor
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

Public Function Clone() As clsLenda
   Dim clonedObject     As New clsLenda

   With clonedObject
      .Emri = Me.Emri
      .Klasa = Me.Klasa
      .VitiShkollor = Me.VitiShkollor
   End With

   Set Clone = clonedObject

   Set clonedObject = Nothing
End Function




