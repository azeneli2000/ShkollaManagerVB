VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDBMapper"
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
' * Time             : 20:53
' * Module Name      : clsDBMapper
' * Module Filename  : clsDBMapper.cls
' * *******************************************************************
' * Comments         :
' *
' *
' * *******************************************************************




'tabelat
Const tblAmza As String = "TBL_AMZA"
Const tblLenda As String = "TBL_LENDA"
Const tblNota As String = "TBL_NOTA"
Const tblNotaAmz As String = "TBL_NOTAAMZ"
Const tblSjellje As String = "TBL_SJELLJE"
Const tblPerdorues As String = "TBL_PERDORUES"



'local variable(s) to hold property value(s)
Private mvarErrorHandler As New clsErrorHandler
Private mvarQueryString As String
Private mvarActionName As DB_MANAGER_ENUM

Public Enum QUERY_TYPE_ENUM
   qtSel = 1
   qtIns = 2
   qtDel = 3
   qtUpd = 4
End Enum



Public Property Let actionName(ByVal vData As DB_MANAGER_ENUM)
    mvarActionName = vData
End Property


Public Property Get actionName() As DB_MANAGER_ENUM
   actionName = mvarActionName
End Property


Public Property Let QueryString(ByVal vData As String)
    mvarQueryString = vData
End Property


Public Property Get QueryString() As String
   QueryString = mvarQueryString
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


Public Function getQuery() As String

   Select Case Me.actionName
     
      
      Case INSERT_PERDORUES
         getQuery = " Insert Into TBL_USERS (USER_NAME, USER_PASSWORD) values "
         
      Case UPDATE_PERDORUES
          getQuery = "Update TBL_USERS (USER_NAME, USER_PASSWORD) values "
      
      Case GJEJ_PERDORUES
           getQuery = "Select * from TBL_USERS where USER_NAME ="
   End Select
End Function


Private Function Qo_tblLenda_Sel_All() As String
   Qo_tblLenda_Sel_All = "Select * From " & tblLenda
End Function


Private Function getQueryStringSelDel(tableName As String, queryType As QUERY_TYPE_ENUM, Optional Cond As String = "") As String
   Select Case queryType

      Case qtSel
         getQueryStringNoP = "Select * From " & tableName

         If Len(Cond) <> 0 Then
            getQueryStringNoP = getQueryStringNoP & " Where " & Cond
         End If
      Case qtDel

         getQueryStringNoP = "Delete From " & tableName

         If Len(Cond) <> 0 Then
            getQueryStringNoP = getQueryStringNoP & " Where " & Cond
         End If

   End Select

End Function


Private Function Condition(argParams As Variant) As String

   Condition = ""

   Dim vParams          As Variant
   Dim i                As Integer, v As Variant

   'Point to first element in array
   vParams = argParams

   If IsArray(vParams) Then

      vParams = " Where "
      'For each array element
      For i = LBound(vParams) To UBound(vParams)
         Condition = Condition & vParams(i)(0) & " = " & vParams(i)(1)
      Next i

      Condition = Left(Condition, Len(Condition) - 7)

   End If

End Function
