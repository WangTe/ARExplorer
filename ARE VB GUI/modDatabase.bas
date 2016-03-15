Attribute VB_Name = "modDatabase"
Option Explicit
'This is where all of the actual interfacing with the Database is handled.


'Private Const sDBFileName = "\OC.cch"  'Name of our DB, will be kept in
Private Const sDBFileName = "\OC.cch"  'same dir as the application
Private Const sTempDBFileName = "\toc.cch"
Private Const sOldDBFileName = "\ooc.cch"

Private wsAccess As Workspace
Private Const sWSName = "Access"

Private dbFile As Database

Private rsServerTable As Recordset
Private rsFormTable As Recordset
Private rsFieldTable As Recordset
Private rsQueryTable As Recordset
Private rsQueryItemsTable As Recordset



'Initializes the Database.
Public Function InitializeDB() As Boolean
Dim sMSG As String
Dim i As Integer

  If Len(Dir(App.Path & sDBFileName)) > 0 Then
    Set wsAccess = DBEngine.CreateWorkspace(sWSName, "Admin", "")
    Set dbFile = wsAccess.OpenDatabase(App.Path & sDBFileName)
    Set rsServerTable = dbFile.OpenRecordset(sServerTableName, dbOpenTable)
    Set rsFormTable = dbFile.OpenRecordset(sFormTableName, dbOpenTable)
    Set rsFieldTable = dbFile.OpenRecordset(sFieldTableName, dbOpenTable)
    Set rsQueryTable = dbFile.OpenRecordset(sQueryTableName, dbOpenTable)
    Set rsQueryItemsTable = dbFile.OpenRecordset(sQueryItemsTableName, dbOpenTable)
    InitializeDB = True
  Else
    sMSG = "The database " & App.Path & sDBFileName & " does not exist, can not continue."
    i = MsgBox(sMSG, vbOKOnly + vbCritical)
    InitializeDB = False
  End If
  
End Function


'Checks for a duplicate QueryName
Public Function CheckForDup(sQueryName As String) As Boolean
Dim sSQL As String
Dim rsResult As Recordset

  sSQL = "SELECT * "  '[" & fldQueryName & "] "
  sSQL = sSQL & "FROM [" & sQueryTableName & "] As Tmp "
  sSQL = sSQL & "WHERE [" & fldQueryName & "] LIKE '" & sQueryName & "';"
  
  wsAccess.BeginTrans
  Set rsResult = dbFile.OpenRecordset(sSQL, dbOpenDynaset)
  wsAccess.CommitTrans
  
  On Error Resume Next    'This is needed because the next two statements will
  rsResult.MoveLast
  rsResult.MoveFirst
  On Error GoTo 0
  
  If rsResult.RecordCount > 0 Then
    CheckForDup = True
  Else
    CheckForDup = False
  End If

End Function


Public Function GetSavedQueryID(lIndex As Long) As Long
Dim i As Long
Dim lReturnValue As Long

  If lIndex <= rsQueryTable.RecordCount Then
    On Error Resume Next
    rsQueryTable.MoveLast
    rsQueryTable.MoveFirst
    On Error GoTo 0
    
    If lIndex > 1 Then
      For i = 1 To lIndex - 1
        rsQueryTable.MoveNext
      Next i
      
    End If
    
    lReturnValue = rsQueryTable(fldID)
  Else
    lReturnValue = 0
  End If
  
  GetSavedQueryID = lReturnValue

End Function



Public Function GetSavedQueryName(lIndex As Long) As String
Dim i As Long
Dim sReturnValue As String

  If lIndex <= rsQueryTable.RecordCount Then
    On Error Resume Next
    rsQueryTable.MoveLast
    rsQueryTable.MoveFirst
    On Error GoTo 0
    
    If lIndex > 1 Then
      For i = 1 To lIndex - 1
        rsQueryTable.MoveNext
      Next i
      
    End If
    
    sReturnValue = rsQueryTable(fldQueryName)
  Else
    sReturnValue = "Index excedes record count"
  End If
  
  GetSavedQueryName = sReturnValue

End Function

Public Function GetSavedQueryType(lIndex As Long) As String
Dim i As Long
Dim sReturnValue As String

  If lIndex <= rsQueryTable.RecordCount Then
    On Error Resume Next
    rsQueryTable.MoveLast
    rsQueryTable.MoveFirst
    On Error GoTo 0
    
    If lIndex > 1 Then
      For i = 1 To lIndex - 1
        rsQueryTable.MoveNext
      Next i
      
    End If
    
    sReturnValue = rsQueryTable(fldType)
  Else
    sReturnValue = "Index excedes record count"
  End If
  
  GetSavedQueryType = sReturnValue

End Function

Public Function GetSavedQueryCount() As Long
  GetSavedQueryCount = rsQueryTable.RecordCount
End Function


'Public Function DeleteQuery(lQueryID As Long) As Boolean
'Dim sSQL As String
'Dim i As Integer
'Dim sMSG As String
'Dim rsResult As Recordset
'
'  sSQL = "DELETE tblQueryItems.*, tblQueryItems.QueryID "
'  sSQL = sSQL & "FROM tblQueryItems "
'  sSQL = sSQL & "WHERE (((tblQueryItems.QueryID)= " & Trim(Str(lQueryID)) & "));"
'
'  wsAccess.BeginTrans
'  dbFile.Execute (sSQL)
' ' Set rsResult = dbFile.OpenRecordset(sSQL, dbOpenDynaset)
'  wsAccess.CommitTrans
'
'  sSQL = "DELETE * "
'  sSQL = sSQL & "FROM " & sQueryTableName & " "
'  sSQL = sSQL & "WHERE " & fldID & " = " & Trim(Str(lQueryID)) & ";"
'
'  wsAccess.BeginTrans
'  dbFile.Execute (sSQL)
'  wsAccess.CommitTrans
'
'
'End Function


'Save one query to the DB.
'If query already exists, will overwrite only if bOverwrite = true
Public Function SaveQuery(ByRef QueryToAdd As colQueryList, Optional bOverwrite As Boolean) As Boolean
Dim i As Integer
Dim sMSG As String
Dim qryiItem As clsQueryItem
Dim lQueryID As Long

  'make sure that there isn't already a query with the same query name.
  If (CheckForDup(QueryToAdd.SaveName) = False) Or (bOverwrite = True) Then
    wsAccess.BeginTrans

    rsQueryTable.AddNew
    With QueryToAdd
      rsQueryTable(fldQueryName) = .SaveName
      rsQueryTable(fldType) = .SearchType
      'rsQueryTable(fldConditionString) = .SearchType
      rsQueryTable(fldExecuteOnValue) = .ExecuteOnValue
      rsQueryTable(fldExecuteOnCount) = .ExecuteOnCount
      If .CaseSensitive = True Then
        rsQueryTable(fldCaseSensitive) = vbYes
      Else
        rsQueryTable(fldCaseSensitive) = vbNo
      End If
      lQueryID = rsQueryTable(fldID)
    End With
    
    rsQueryTable.Update

    'wsAccess.CommitTrans

    For i = 1 To QueryToAdd.Count
      Set qryiItem = QueryToAdd.Item(i)
      
      With qryiItem
        rsQueryItemsTable.AddNew
        rsQueryItemsTable(fldQueryID) = lQueryID
        rsQueryItemsTable(fldSearchParam) = .SearchParam
        rsQueryItemsTable(fldConditionString) = .SearchConditionString
        rsQueryItemsTable(fldConditionValue) = .SearchConditionValue
        rsQueryItemsTable(fldSearchType) = .SearchType
        rsQueryItemsTable(fldSearchValueNumber) = .SearchValueNum
        rsQueryItemsTable(fldSearchValueString) = .SearchValueString
        rsQueryItemsTable(fldSearchANDOR) = .SearchCondition
      End With
      
      rsQueryItemsTable.Update
    Next i
    

    wsAccess.CommitTrans
    
    SaveQuery = True

  Else
    sMSG = "There is already a query by this name saved.  "
    sMSG = sMSG & "Please choose a new name."
    i = MsgBox(sMSG, vbOKOnly + vbInformation)
    SaveQuery = False
  End If

End Function


Public Function DeleteQuery(sQueryName As String) As Boolean
Dim rsResult As Recordset
Dim sSQL As String
Dim lParentID As Long

  sSQL = "SELECT * "
  sSQL = sSQL & "FROM [" & sQueryTableName & "] "
  sSQL = sSQL & "WHERE [" & fldQueryName & "] LIKE '" & sQueryName & "';"
  
  wsAccess.BeginTrans
  'dbFile.Execute sSQL, dbFailOnError
  Set rsResult = dbFile.OpenRecordset(sSQL, dbOpenDynaset)
  wsAccess.CommitTrans
  
  On Error Resume Next
  rsResult.MoveLast
  rsResult.MoveFirst
  On Error GoTo 0
  
  lParentID = rsResult(fldID)
  
  'Set rsResult = Nothing
  If rsResult.RecordCount > 0 Then
  
    sSQL = "DELETE * "
    sSQL = sSQL & "FROM [" & sQueryItemsTableName & "] "
    sSQL = sSQL & "WHERE [" & fldQueryID & "] = " & Trim(Str(lParentID)) & ";"
    
    wsAccess.BeginTrans
    dbFile.Execute sSQL, dbFailOnError
    wsAccess.CommitTrans
    
    sSQL = "DELETE * "
    sSQL = sSQL & "FROM [" & sQueryTableName & "] "
    sSQL = sSQL & "WHERE [" & fldID & "] = " & Trim(Str(lParentID)) & ";"
    
    wsAccess.BeginTrans
    dbFile.Execute sSQL, dbFailOnError
    wsAccess.CommitTrans
  
    DeleteQuery = True
  
  Else
    DeleteQuery = False
  End If

End Function


'Add one entry to the DB.
'Pulls it's information from the form frmData, (too many to pass, and is
'specific to this app)
Public Function OpenQuery(ByRef qryQuery As colQueryList) As Boolean
Dim i As Long
Dim qryiItem As clsQueryItem
Dim lQueryID As Long
Dim sSQL As String
Dim rsResult As Recordset
Dim lParamater As Long
Dim sConditionString As String
Dim lConditionValue As Long
Dim sType As String
Dim lValueNum As Long
Dim sValueString As String
Dim sAndOr As String
Dim lCaseValue As Long
Dim sKey As String

  'make sure that there is a query with the same query name.
  If CheckForDup(qryQuery.SaveName) = True Then

    sSQL = "SELECT  * "
    sSQL = sSQL & "FROM [" & sQueryTableName & "] As Tmp "
    sSQL = sSQL & "WHERE [" & fldQueryName & "] LIKE '" & qryQuery.SaveName & "';"
  
    wsAccess.BeginTrans
    Set rsResult = dbFile.OpenRecordset(sSQL, dbOpenDynaset)
    wsAccess.CommitTrans
    
    On Error Resume Next
    rsResult.MoveLast
    rsResult.MoveFirst
    On Error GoTo 0
    
    With qryQuery
      .SearchType = rsResult(fldType)
      .ExecuteOnValue = rsResult(fldExecuteOnCount)
      .ExecuteOnCount = rsResult(fldExecuteOnValue)
      lCaseValue = rsResult(fldCaseSensitive)
      .SavedID = rsResult(fldID)
      If lCaseValue = vbYes Then
        .CaseSensitive = True
      Else
        .CaseSensitive = False
      End If
      lQueryID = rsResult(fldID)
      .Saved = True
      .Dirty = False
    End With
    
    rsResult.Close
    
    sSQL = "SELECT * "
    sSQL = sSQL & "FROM [" & sQueryItemsTableName & "] "
    sSQL = sSQL & "WHERE [" & fldQueryID & "] = " & Trim(Str(lQueryID)) & ";"
    
    Set rsResult = Nothing
    
    wsAccess.BeginTrans
    Set rsResult = dbFile.OpenRecordset(sSQL, dbOpenDynaset)
    wsAccess.CommitTrans

    On Error Resume Next
    rsResult.MoveLast
    rsResult.MoveFirst
    On Error GoTo 0

    For i = 1 To rsResult.RecordCount
      sType = rsResult(fldSearchType)
      lParamater = rsResult(fldSearchParam)
      sValueString = rsResult(fldSearchValueString)
      lValueNum = rsResult(fldSearchValueNumber)
      sAndOr = rsResult(fldSearchANDOR)
      sConditionString = rsResult(fldConditionString)
      lConditionValue = rsResult(fldConditionValue)
      
      sKey = rsResult(fldSearchType) & rsResult(fldSearchValueString)
      
      'ARQuery.Add cboxProperties.Text, cboxConditions.tag, sValueString, sValueNum, "AND", cboxConditions.Text
      qryQuery.Add sType, lParamater, sValueString, lValueNum, sAndOr, sConditionString, sKey

      rsResult.MoveNext
    Next i
    
    OpenQuery = True

  Else
    OpenQuery = False
  End If
  

End Function

''Delete one entry from the DB.
'Public Function DeleteEntryFromDB(sBarCode As String) As Boolean
'Dim rsResult As Recordset
'Dim sSQL As String
'Dim sMSG As String
'Dim i As Integer
'
'  sSQL = "SELECT * "
'  sSQL = sSQL & "FROM [RetentionData] "
'  sSQL = sSQL & "WHERE [BarCodeNumber] LIKE '" & sBarCode & "';"
'
'  wsAccess.BeginTrans
'    Set rsResult = dbFile.OpenRecordset(sSQL, dbOpenDynaset)
'  wsAccess.CommitTrans
'
'  sMSG = "Are you sure you wish to delete the following record:  "
'  sMSG = sMSG & "Series: " & rsResult("SeriesCode")
'  sMSG = sMSG & "; Description: " & rsResult("Description")
'  sMSG = sMSG & "; Bar Code: " & rsResult("BarCodeNumber")
'
'
'  i = MsgBox(sMSG, vbYesNo + vbQuestion)
'
'  If i = vbYes Then
'    sSQL = "DELETE [BarCodeNumber] "
'    sSQL = sSQL & "FROM [RetentionData] "
'    sSQL = sSQL & "WHERE [BarCodeNumber] LIKE '" & sBarCode & "';"
'
'    wsAccess.BeginTrans
'      dbFile.Execute sSQL, dbFailOnError
'    wsAccess.CommitTrans
'
'  End If
'
'  Set rsResult = Nothing
'
'End Function
'
'
'Close the Database and free up our objects
Public Sub CloseDB()

  rsFormTable.Close
  Set rsFormTable = Nothing
  rsFieldTable.Close
  Set rsFieldTable = Nothing
  rsQueryTable.Close
  Set rsQueryTable = Nothing
  rsQueryItemsTable.Close
  Set rsQueryItemsTable = Nothing
  dbFile.Close
  Set dbFile = Nothing
  wsAccess.Close
  Set wsAccess = Nothing
    
  CompactDB

End Sub


Private Sub CompactDB()

  If Len(Dir(App.Path & sOldDBFileName)) > 0 Then
    Kill (App.Path & sOldDBFileName)
  End If
  'Need to figure out how to rename the file back to the original and delete the temp file
  CompactDatabase App.Path & sDBFileName, App.Path & sTempDBFileName
  
  FileCopy App.Path & sDBFileName, App.Path & sOldDBFileName
  FileCopy App.Path & sTempDBFileName, App.Path & sDBFileName
  
  If Len(Dir(App.Path & sTempDBFileName)) > 0 Then
    Kill (App.Path & sTempDBFileName)
  End If
  
End Sub



'***************************************************************************************************
'Cache stuff go here
'***************************************************************************************************
Public Function AddServerToCache(sServerName As String) As Long
Dim lID As Long

  wsAccess.BeginTrans
  
  rsServerTable.AddNew
  rsServerTable(fldName) = sServerName
  
  lID = rsServerTable(fldID)
  rsServerTable.Update
  
  wsAccess.CommitTrans
  
  AddServerToCache = lID

End Function


'Adds one form to the cache and returns it's ID
Public Function AddFormToCache(sFormName As String, lModTime As Long, lServerID As Long) As Long
Dim lID As Long

  wsAccess.BeginTrans
  
  rsFormTable.AddNew
  rsFormTable(fldName) = sFormName
  rsFormTable(fldModTime) = lModTime
  rsFormTable(fldParentServerID) = lServerID
  
  lID = rsFormTable(fldID)
  rsFormTable.Update
  
  wsAccess.CommitTrans
  
  AddFormToCache = lID
  
End Function


Public Function AddFieldToCache(sFieldName As String, lARID As Long, sFieldType As String, lParentFormID As Long) As Long
Dim lID As Long

  wsAccess.BeginTrans
  
  rsFieldTable.AddNew
  rsFieldTable(fldName) = sFieldName
  rsFieldTable(fldARID) = lARID
  rsFieldTable(fldType) = sFieldType
  rsFieldTable(fldParentFormID) = lParentFormID
  
  lID = rsFieldTable(fldID)
  
  rsFieldTable.Update
  
  wsAccess.CommitTrans
  
  AddFieldToCache = lID
  
End Function


Public Function GetServerCacheID(sServerName As String) As Long
Dim sSQL As String
Dim lID As Long
Dim rsResult As Recordset

  sSQL = "SELECT * "
  sSQL = sSQL & "FROM [" & sServerTableName & "] "
  sSQL = sSQL & "WHERE [" & fldName & "] LIKE '" & sServerName & "';"
  
  wsAccess.BeginTrans
  Set rsResult = dbFile.OpenRecordset(sSQL, dbOpenDynaset)
  wsAccess.CommitTrans
  
  On Error Resume Next
  rsResult.MoveLast
  rsResult.MoveFirst
  On Error GoTo 0
  
  If rsResult.RecordCount > 0 Then
    lID = rsResult(fldID)
  Else
    lID = 0
  End If
  
  GetServerCacheID = lID

End Function


Public Function GetFormName(lCacheID As Long) As String
Dim sSQL As String
Dim sName As String
Dim rsResult As Recordset
Dim i As Integer

  sSQL = "SELECT * "
  sSQL = sSQL & "FROM [" & sFormTableName & "] "
  sSQL = sSQL & "WHERE ([" & fldID & "] = " & Trim(Str(lCacheID)) & ");"
  
'  i = MsgBox("Getting Form Name for form id: " & Trim(Str(lCacheID)), vbOKOnly)
  wsAccess.BeginTrans
  Set rsResult = dbFile.OpenRecordset(sSQL, dbOpenDynaset)
  wsAccess.CommitTrans
  
  On Error Resume Next
  rsResult.MoveLast
  rsResult.MoveFirst
  On Error GoTo 0
  
  If rsResult.RecordCount > 0 Then
    sName = rsResult(fldName)
  Else
    sName = " "
  End If
  
  GetFormName = sName

End Function


Public Function GetFormCacheID(sFormName As String, lServerID As Long) As Long
Dim sSQL As String
Dim lID As Long
Dim rsResult As Recordset

  sSQL = "SELECT * "
  sSQL = sSQL & "FROM [" & sFormTableName & "] "
  sSQL = sSQL & "WHERE ([" & fldName & "] LIKE '" & sFormName & "') "
  sSQL = sSQL & "AND ([" & fldParentServerID & "] = " & Trim(Str(lServerID)) & ");"
  
  wsAccess.BeginTrans
  Set rsResult = dbFile.OpenRecordset(sSQL, dbOpenDynaset)
  wsAccess.CommitTrans
  
  On Error Resume Next
  rsResult.MoveLast
  rsResult.MoveFirst
  On Error GoTo 0
  
  If rsResult.RecordCount > 0 Then
    lID = rsResult(fldID)
  Else
    lID = 0
  End If
  
  GetFormCacheID = lID

End Function


Public Function GetFormCacheModTime(lFormID As Long) As Long
Dim sSQL As String
Dim lModTime As Long
Dim rsResult As Recordset

  sSQL = "SELECT * "
  sSQL = sSQL & "FROM [" & sFormTableName & "] "
  sSQL = sSQL & "WHERE [" & fldID & "] = " & Trim(Str(lFormID)) & ";"
  
  wsAccess.BeginTrans
  Set rsResult = dbFile.OpenRecordset(sSQL, dbOpenDynaset)
  wsAccess.CommitTrans
  
  On Error Resume Next
  rsResult.MoveLast
  rsResult.MoveFirst
  On Error GoTo 0
  
  If rsResult.RecordCount > 0 Then
    lModTime = rsResult(fldModTime)
  Else
    lModTime = 0
  End If
  
  GetFormCacheModTime = lModTime

End Function


Public Sub DeleteCache(lServerID As Long)
Dim sSQL As String

  sSQL = "DELETE * "
  sSQL = sSQL & "FROM [" & sServerTableName & "] "
  sSQL = sSQL & "WHERE [" & fldID & "] = " & Trim(Str(lServerID)) & ";"
  
  wsAccess.BeginTrans
  dbFile.Execute sSQL, dbFailOnError
  wsAccess.CommitTrans
  
End Sub


'Can I make this easier by using Cascade Delete inside the DB?
'(simply deleting the Form will auto delete all related records?)
Public Sub DeleteFormFromCache(lFormID As Long)
Dim sSQL As String
Dim i As Long

  If lFormID <> 0 Then
    sSQL = "DELETE * "
    sSQL = sSQL & "FROM [" & sFormTableName & "] "
    sSQL = sSQL & "WHERE [" & fldID & "] = " & Trim(Str(lFormID)) & ";"
    
    wsAccess.BeginTrans
    dbFile.Execute sSQL, dbFailOnError
    wsAccess.CommitTrans
  End If

End Sub


Public Function ExecuteCacheSQL(sSQL As String) As Recordset
Dim rsResult As Recordset

  On Error GoTo ErrHandler

  wsAccess.BeginTrans
  Set rsResult = dbFile.OpenRecordset(sSQL, dbOpenDynaset)
  wsAccess.CommitTrans
  
  Set ExecuteCacheSQL = rsResult
  
  Exit Function
  
ErrHandler:
  Dim sMSG As String
  Dim i As Integer
  
  sMSG = "Could not execute a SQL query to the AR Explorer Cache system."
  i = MsgBox(sMSG, vbOKCancel + vbCritical, "Critical Error:")
  Set rsResult = Nothing
  Resume Next
  
End Function

