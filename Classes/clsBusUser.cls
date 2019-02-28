VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBusUser"
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
' * Date             : 11/29/2004
' * Time             : 14:58
' * Module Name      : clsBusUser
' * Module Filename  : clsBusUser.cls
' * *******************************************************************
' * Comments         :
' *
' *
' * *******************************************************************





'local variable(s) to hold property value(s)
Private mvarUserName As String
Private mvarUserPassword As String
Private mvarKey As String


Public Property Let Key(ByVal vData As String)
  mvarKey = vData
End Property


Public Property Get Key() As String
   Key = mvarKey
End Property


Public Property Let UserPassword(ByVal vData As String)
    mvarUserPassword = vData
End Property


Public Property Get UserPassword() As String
   UserPassword = mvarUserPassword
End Property


Public Property Let UserName(ByVal vData As String)
    mvarUserName = vData
End Property


Public Property Get UserName() As String
   UserName = mvarUserName
End Property


Public Function Clone() As clsBusUser
   Dim objClone         As New clsBusUser

   With objClone
      .UserName = Me.UserName
      .UserPassword = Me.UserPassword
   End With

   Set Clone = objClone

   Set objClone = Nothing

End Function


Public Function Load(rs As ADODB.Recordset) As clsBusUser

   If Not rs.EOF Then
      With Me
         .Key = rs("USER_ID")
         .UserName = rs("USER_NAME")
         .UserPassword = rs("USER_PASSWORD")
      End With
   End If

   Set Load = Me

End Function


Public Function LoadV(vParams As Variant) As clsBusUser

   With Me
      .Key = vParams(0, 0)
      .UserName = vParams(1, 0)
      .UserPassword = vParams(2, 0)
   End With

   Set LoadV = Me
End Function


Public Function GetGUIUser_() As clsBusUser

   '' to do
   With active_form

   End With

   Set GetGUIUser_ = Me

End Function
