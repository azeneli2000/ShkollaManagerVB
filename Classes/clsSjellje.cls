VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSjellje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarVitiShkollor As String 'local copy
Private mvarShenimeSjellje As String 'local copy

Public Function Clone() As clsSjellje
End Function

Public Property Let ShenimeSjellje(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ShenimeSjellje = 5
    mvarShenimeSjellje = vData
End Property


Public Property Get ShenimeSjellje() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ShenimeSjellje
    ShenimeSjellje = mvarShenimeSjellje
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



