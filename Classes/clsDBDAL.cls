VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDBDAL"
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
' * Date             : 11/08/2004
' * Time             : 17:47
' * Module Name      : clsDAL
' * Module Filename  : clsDAL.cls
' * *******************************************************************
' * Comments         :
' *
' *
' * *******************************************************************


Option Explicit

'Class constants
Const C_CLASSNAME = "clsDAL"

Public rec As ADODB.Recordset

Public dbPath As String
Const dbName As String = "dbStudAdmin1.mdb"


'Wrap ADO CommandTypeEnum with our own to future proof
Public Enum QueryOptions
    QO_UseTable = ADODB.adCmdTable
    QO_UseText = ADODB.adCmdText
    QO_UseStoredProc = ADODB.adCmdStoredProc
    QO_UseUnknown = ADODB.adCmdUnknown
    QO_UseDefault = -1
End Enum

'String handling options
Public Enum StringHandling
    SH_GetCSV = 1
    SH_GetHTML = 2
End Enum

'local variable(s) to hold property value(s)
Private mvarConnection_String As String
Private mvarErrorHandler As New clsErrorHandler
Private mvarActionName As DB_MANAGER_ENUM


Public Property Let actionName(ByVal vData As DB_MANAGER_ENUM)
    mvarActionName = vData
End Property


Public Property Get actionName() As DB_MANAGER_ENUM
   actionName = mvarActionName
End Property


Public Property Set ErrorHandler(ByVal vData As clsErrorHandler)
    Set mvarErrorHandler = vData
End Property


Public Property Get ErrorHandler() As clsErrorHandler
   Set ErrorHandler = mvarErrorHandler
End Property


Public Property Let Connection_String(ByVal vData As String)
    mvarConnection_String = vData
End Property


Public Property Get Connection_String() As String
   Connection_String = mvarConnection_String
End Property


Private Sub Class_Terminate()
   Set ErrorHandler = Nothing
End Sub


Public Function getConnectionString() As String
   dbPath = App.Path & "\"
   getConnectionString = "Provider = Microsoft.Jet.OLEDB.4.0; Persist Security Info = False; Data Source = " & dbPath & dbName & "; User Id = Admin; Password =;"
End Function


Public Function GetRecordSet(ByVal sQuery As String, ByVal sConnect As String, eType As QueryOptions) As Recordset

   'Error handling
  ' On Error GoTo ErrorHandler

   'Create the ADO objects
   Dim rs               As ADODB.Recordset, Cmd As ADODB.Command
   Set rs = New ADODB.Recordset
   Set Cmd = New ADODB.Command

   Cmd.CommandText = sQuery
   Cmd.CommandType = eType

   'Use helper function to build parameters for command object

   'Determine whether passed in, or hard coded connection
   If sConnect = vbNullString Then
      Cmd.ActiveConnection = getConnectionString()
   Else
      Cmd.ActiveConnection = sConnect
   End If

   'Init the ADO objects & the query parameters (if any)
   'CollectParams Cmd, aParams
   'Cmd.CommandTimeout = 0
   'Cmd.ActiveConnection.ConnectionTimeout = 0

   'Execute the query for readonly
   rs.CursorLocation = adUseClient
   rs.CursorType = adOpenForwardOnly
   rs.LockType = adLockBatchOptimistic
   rs.Open Cmd

   'Disconnect the recordset
   Set Cmd.ActiveConnection = Nothing
   Set Cmd = Nothing
   Set rs.ActiveConnection = Nothing

   'Return the resultant recordset
   Set GetRecordSet = rs

   Exit Function

'ErrorHandler:

   'Clean up
 '  Set rs = Nothing
  ' Set Cmd = Nothing
   'Set GetRecordSet = Nothing

   'Raise error
   'Call RaiseError(C_CLASSNAME, "GetRecordSet")

End Function


Public Sub PutRecordset(rs As ADODB.Recordset, Optional ByVal sConnect As String = "")

   'Error handling
   On Error GoTo ErrorHandler

   'Reconnect to database
   If sConnect = vbNullString Then
      rs.ActiveConnection = getConnectionString()
   Else
      rs.ActiveConnection = sConnect
   End If

   'Update records
   rs.UpdateBatch

   'Disconnect the recordset again and exit
   Set rs.ActiveConnection = Nothing
   Exit Sub

ErrorHandler:

   'Raise error
   Call RaiseError(C_CLASSNAME, "PutRecordSet")

End Sub


Public Function Execute(ByVal sQuery As String, ByVal sConnect As String, eType As QueryOptions) As Boolean

   Dim Cmd              As ADODB.Command
   Dim i                As Integer
   'Enable error handler
   'On Error GoTo ErrorHandler

   'Assume failure
   Execute = False

   'Create the ADO objects
   Set Cmd = New ADODB.Command
   'Init the ADO objects
   Cmd.CommandText = sQuery
   Cmd.CommandType = eType
   'Use helper function to build parameters for command object

   'todo
   Cmd.CommandTimeout = 10
   'Determine whether passed in, or hard coded connection
   If sConnect = vbNullString Then
      Cmd.ActiveConnection = getConnectionString()
   Else
      Cmd.ActiveConnection = sConnect
   End If

   'CollectParams Cmd, aParams

   'Execute the query without returning a recordset
   Cmd.Execute , , ADODB.adExecuteNoRecords

  ' If Cmd.Parameters.Count > 0 Then NewID = Cmd(aParams(0)(0))


   'Cleanup and return nothing
   Set Cmd.ActiveConnection = Nothing
   Set Cmd = Nothing

   'Return

   Execute = True
   Exit Function

'ErrorHandler:

   'Cleanup and return nothing
  ' Set Cmd.ActiveConnection = Nothing
  ' Set Cmd = Nothing

   'Raise error
  ' Call RaiseError(C_CLASSNAME, "Execute")

End Function


Public Function GetArray(ByVal sQuery As String, ByVal sConnect As String, eType As QueryOptions, ParamArray aParams() As Variant) As Variant()

   'Error handling
   On Error GoTo ErrorHandler

   'Create the ADO objects
   Dim rs               As ADODB.Recordset, Cmd As ADODB.Command
   Set rs = New ADODB.Recordset
   Set Cmd = New ADODB.Command

   'Use helper function to build parameters for command object
   CollectParams Cmd, aParams

   'Determine whether passed in, or hard coded connection
   If sConnect = vbNullString Then
      Cmd.ActiveConnection = getConnectionString()
   Else
      Cmd.ActiveConnection = sConnect
   End If

   'Init the ADO objects & the query parameters (if any)
   Cmd.CommandText = sQuery
   Cmd.CommandType = eType

   'Execute the query for readonly
   rs.CursorLocation = adUseClient
   rs.CursorType = adOpenForwardOnly
   rs.LockType = adLockBatchOptimistic
   rs.Open Cmd

   'Convert the recordset to an array
   If rs.EOF = False And rs.BOF = False Then
      GetArray = rs.GetRows
   Else
      ReDim GetArray(-1 To -1, -1 To -1)
   End If

   'Clean up and exit
   Set rs = Nothing
   Set Cmd = Nothing
   Exit Function

ErrorHandler:

   'Clean up (array left empty)
   Set rs = Nothing
   Set Cmd = Nothing

   'Raise error
   Call RaiseError(C_CLASSNAME, "GetArray")

End Function


Public Function GetString(ByVal sQuery As String, ByVal sConnect As String, eType As QueryOptions, eStringType As StringHandling, ParamArray aParams() As Variant) As String

   'Error handling
   On Error GoTo ErrorHandler

   'Create the ADO objects
   Dim rs               As ADODB.Recordset, Cmd As ADODB.Command
   Set rs = New ADODB.Recordset
   Set Cmd = New ADODB.Command

   'Use helper function to build parameters for command object
   CollectParams Cmd, aParams

   'Determine whether passed in, or hard coded connection
   If sConnect = vbNullString Then
      Cmd.ActiveConnection = getConnectionString()
   Else
      Cmd.ActiveConnection = sConnect
   End If

   'Init the ADO objects & the query parameters (if any)
   Cmd.CommandText = sQuery
   Cmd.CommandType = eType

   'Execute the query for readonly
   rs.CursorLocation = adUseClient
   rs.CursorType = adOpenForwardOnly
   rs.LockType = adLockBatchOptimistic
   rs.Open Cmd

   'Convert the recordset to a string
   If rs.EOF = False And rs.BOF = False Then

      'Dermine formatting
      If eStringType = SH_GetHTML Then
         GetString = ConvertToHTML(rs)
      Else
         GetString = ConvertToCSV(rs)
      End If

   Else
      GetString = ""
   End If

   'Clean up and exit
   Set rs = Nothing
   Set Cmd = Nothing
   Exit Function

ErrorHandler:

   'Clean up (string left empty)
   Set rs = Nothing
   Set Cmd = Nothing

   'Raise error
   Call RaiseError(C_CLASSNAME, "GetString")

End Function


Public Function ConvertToCSV(rs As ADODB.Recordset) As String

   'Check for closed or empty recordset
   If rs.EOF = True Or rs.BOF = True Then
      ConvertToCSV = ""
      Exit Function
   End If

   'Convert recordset to comma seperated values
   ConvertToCSV = rs.GetString(adClipString, -1, ",", vbCrLf, "(NULL)")

End Function


Public Function ConvertToHTML(rs As ADODB.Recordset) As String

   'Check for closed or empty recordset
   If rs.EOF = True Or rs.BOF = True Then
      ConvertToHTML = ""
      Exit Function
   End If

   'Convert recordset to HTML table syntax
   ConvertToHTML = rs.GetString(adClipString, -1, "</TD><TD>", "</TD></TR>" & vbCrLf & "<TR><TD>", "(NULL)")
   ConvertToHTML = "<TR><TD>" & Left(ConvertToHTML, Len(ConvertToHTML) - 8)

End Function


Private Sub CollectParams(ByRef Cmd As ADODB.Command, ParamArray argParams() As Variant)
   On Error GoTo err_handler

   Dim vParams          As Variant
   Dim i                As Integer, v As Variant

   'Point to first element in array
   vParams = argParams(0)

   'For each array element
   For i = LBound(vParams) To UBound(vParams)
      Cmd(vParams(i)(0)) = vParams(i)(3)
   Next i

   Exit Sub

err_handler:
   MsgBox Err.Description & " tek fusha '" & vParams(i)(0) & "'"
   Call RaiseError(C_CLASSNAME, "CollectParams")
End Sub

Public Sub ex(strsql As String)
Dim s As String
Dim con As ADODB.Connection

' s = getConnectionString()
 Set con = New ADODB.Connection
 con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & "C:\Documents and Settings\Andi\Desktop\dbStudAdmin1.mdb;Persist Security Info=False"
 con.Open
 con.Execute strsql
 
 End Sub
Public Sub get_rec(strsql As String)
 Dim con As ADODB.Connection
 Set con = New ADODB.Connection
 Set record1 = New ADODB.Recordset
 
 con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & "C:\Documents and Settings\Andi\Desktop\dbStudAdmin1.mdb;Persist Security Info=False"
 con.Open
 
 record1.Source = strsql
 record1.ActiveConnection = con
 record1.Open
 
End Sub
