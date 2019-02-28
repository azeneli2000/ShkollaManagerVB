VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsActions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

' * Author           :
' * Web Site         :
' * E-Mail           :
' * Date             : 11/10/2004
' * Time             : 16:44
' * Module Name      : clsActions
' * Module Filename  : clsActions.cls
' * *******************************************************************
' * Comments         :
' *
' *
' * *******************************************************************


'local variable(s) to hold property value(s)
Private mvarActionName As ACTION_CASE_ENUM 'local copy
Private mvarErrorHandler As clsErrorHandler 'local copy
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


Public Property Let ActionName(ByVal vData As ACTION_CASE_ENUM)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.actionName = Form1
    mvarActionName = vData
End Property


Public Property Get ActionName() As ACTION_CASE_ENUM
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.actionName
    ActionName = mvarActionName
End Property

