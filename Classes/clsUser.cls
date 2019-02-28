VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarUserName As String 'local copy
Private mvarUserPassword As String 'local copy


Public Function Verifiko() As Boolean
  
End Function

Public Property Let UserPassword(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UserPassword = 5
    mvarUserPassword = vData
End Property


Public Property Get UserPassword() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UserPassword
    UserPassword = mvarUserPassword
End Property



Public Property Let UserName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UserName = 5
    mvarUserName = vData
End Property


Public Property Get UserName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UserName
    UserName = mvarUserName
End Property

Public Function Clone() As clsUser
   Dim objClone         As New clsUser

   With objClone
      .UserName = Me.UserName
      .UserPassword = Me.UserPassword
   End With
   
   Set Clone = objClone
   
   Set objClone = Nothing
   
End Function
