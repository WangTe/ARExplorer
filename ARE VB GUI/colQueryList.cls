VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colQueryList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"UserDefinedCollection"
Attribute VB_Ext_KEY = "Member0" ,"clsQueryItem"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Const AL_ONBUTTON_VAL = 1
Const AL_ONRETURN_VAL = 2
Const AL_ONSUBMIT_VAL = 4
Const AL_ONMODIFY_VAL = 8
Const AL_ONDISPLAY_VAL = 16
Const AL_MODIFYALL_VAL = 32
Const AL_MENUOPEN_VAL = 64
Const AL_MENUCHOICE_VAL = 128
Const AL_LOSEFOCUS_VAL = 256
Const AL_SETDEFAULT_VAL = 512
Const AL_ONQUERY_VAL = 1024
Const AL_AFTERMODIFY_VAL = 2048
Const AL_AFTERSUBMIT_VAL = 4096
Const AL_GAINFOCUS_VAL = 8192
Const AL_WINDOWOPEN_VAL = 16384
Const AL_WINDOWCLOSE_VAL = 32768
Const AL_NONE_VAL = 1073741824

Const FL_SUBMIT_VAL = 4
Const FL_MODIFY_VAL = 2
Const FL_GET_VAL = 1
Const FL_DELETE_VAL = 8
Const FL_MERGE_VAL = 16


'local variable to hold collection
Private mCol As Collection

Private lNumOfExecuteOn As Long
'local variable(s) to hold property value(s)
Private mvarExecuteOnValue As Long 'local copy
Private mvarSavePath As String 'local copy
Private mvarSaveName As String 'local copy
Private mvarSearchServerName As String 'local copy
'local variable(s) to hold property value(s)
Private mvarSearchType As String 'local copy
'local variable(s) to hold property value(s)
Private mvarCaseSensitive As Boolean 'local copy
'local variable(s) to hold property value(s)
Private mvarSaved As Boolean 'local copy
Private mvarDirty As Boolean 'local copy
'local variable(s) to hold property value(s)
Private mvarSavedID As Long 'local copy
'local variable(s) to hold property value(s)
Private mvarExecuteOnANDORValue As String 'local copy
'local variable(s) to hold property value(s)
Private mvarDialogIndex As Long 'local copy
Public Property Let DialogIndex(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DialogIndex = 5
    mvarDialogIndex = vData
End Property


Public Property Get DialogIndex() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DialogIndex
    DialogIndex = mvarDialogIndex
End Property



Public Property Let ExecuteOnANDORValue(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ExecuteOnANDORValue = 5
    mvarExecuteOnANDORValue = vData
End Property


Public Property Get ExecuteOnANDORValue() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ExecuteOnANDORValue
    ExecuteOnANDORValue = mvarExecuteOnANDORValue
End Property



Public Property Let SavedID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SavedID = 5
    mvarSavedID = vData
End Property


Public Property Get SavedID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SavedID
    SavedID = mvarSavedID
End Property



Public Property Let Dirty(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Dirty = 5
    mvarDirty = vData
End Property


Public Property Get Dirty() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Dirty
    Dirty = mvarDirty
End Property



Public Property Let Saved(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Saved = 5
    mvarSaved = vData
End Property


Public Property Get Saved() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Saved
    Saved = mvarSaved
End Property



Public Property Let CaseSensitive(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CaseSensitive = 5
    mvarCaseSensitive = vData
End Property


Public Property Get CaseSensitive() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CaseSensitive
    CaseSensitive = mvarCaseSensitive
End Property


Public Sub ResetCollection()

  RemoveAll
  SearchType = frmMain.sDefaultSearchType
  SearchServerName = ""
  SaveName = "<EMPTY>"
  SavedID = 0
  SavePath = ""
  Saved = False
  Dirty = False
  ExecuteOnANDORValue = "NEITHER"
'  ExecuteOnValue = 0
  lNumOfExecuteOn = 0
  CaseSensitive = False
  
End Sub

Public Sub RemoveAll()
Dim i As Integer

  For i = 1 To Me.Count
    Me.Remove (1)
  Next i

  'ResetCollection
End Sub

Public Property Let SearchType(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SearchType = 5
    mvarSearchType = vData
End Property


Public Property Get SearchType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SearchType
    SearchType = mvarSearchType
End Property



Public Property Let SearchServerName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SearchServerName = 5
    mvarSearchServerName = vData
End Property


Public Property Get SearchServerName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SearchServerName
    SearchServerName = mvarSearchServerName
End Property



Public Property Let SaveName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SaveName = 5
    mvarSaveName = vData
End Property


Public Property Get SaveName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SaveName
    SaveName = mvarSaveName
End Property



Public Property Let SavePath(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SavePath = 5
    mvarSavePath = vData
End Property


Public Property Get SavePath() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SavePath
    SavePath = mvarSavePath
End Property





Private Function GetButtonVal(sButtonType As String) As Long
Dim lValue As Long

  lValue = 0

  If Me.SearchType = TYPE_AL Then
    Select Case sButtonType
      Case AR_ONBUTTON
        lValue = AL_ONBUTTON_VAL
      Case AR_ONRETURN
        lValue = AL_ONRETURN_VAL
      Case AR_ONSUBMIT
        lValue = AL_ONSUBMIT_VAL
      Case AR_ONMODIFY
        lValue = AL_ONMODIFY_VAL
      Case AR_ONDISPLAY
        lValue = AL_ONDISPLAY_VAL
  '    Case AR_MODIFYALL
  '      lValue = AR_MODIFYALL_VAL
  '    Case AR_MENUOPEN
  '      lValue = AR_MENUOPEN_VAL
      Case AR_MENUCHOICE
        lValue = AL_MENUCHOICE_VAL
      Case AR_LOSEFOCUS
        lValue = AL_LOSEFOCUS_VAL
      Case AR_SETDEFAULT
        lValue = AL_SETDEFAULT_VAL
      Case AR_ONQUERY
        lValue = AL_ONQUERY_VAL
      Case AR_AFTERMODIFY
        lValue = AL_AFTERMODIFY_VAL
      Case AR_AFTERSUBMIT
        lValue = AL_AFTERSUBMIT_VAL
      Case AR_GAINFOCUS
        lValue = AL_GAINFOCUS_VAL
      Case AR_WINDOWOPEN
        lValue = AL_WINDOWOPEN_VAL
      Case AR_WINDOWCLOSE
        lValue = AL_WINDOWCLOSE_VAL
      Case AR_NONE
        lValue = AL_NONE_VAL
    End Select
  ElseIf Me.SearchType = TYPE_FILTER Then
    Select Case sButtonType
      Case AR_ONSUBMIT
        lValue = FL_SUBMIT_VAL
      Case AR_ONMODIFY
        lValue = FL_MODIFY_VAL
      Case AR_GET
        lValue = FL_GET_VAL
      Case AR_DELETE
        lValue = FL_DELETE_VAL
      Case AR_MERGE
        lValue = FL_MERGE_VAL
    End Select
  End If

  GetButtonVal = lValue

End Function


Public Sub AddExecuteOnValue(sValue As String)

  mvarExecuteOnValue = mvarExecuteOnValue + (GetButtonVal(sValue))

End Sub


Public Sub SubtractExecuteOnValue(sString As String)

  mvarExecuteOnValue = mvarExecuteOnValue - (GetButtonVal(sString))

End Sub


Public Property Get ExecuteOnValue() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ExecuteOnValue
    ExecuteOnValue = mvarExecuteOnValue
End Property
Public Property Let ExecuteOnValue(lVal As Long)
  mvarExecuteOnValue = lVal
End Property

Public Property Get ExecuteOnCount() As Long
  ExecuteOnCount = lNumOfExecuteOn
End Property
Public Property Let ExecuteOnCount(lVal As Long)
  lNumOfExecuteOn = lVal
End Property


Public Function Add(SearchType As String, SearchParam As Long, SearchValueString As String, SearchValueNum As Long, SearchCondition As String, SearchConditionString As String, Optional sKey As String) As clsQueryItem
    'create a new object
    Dim objNewMember As clsQueryItem
    
    Set objNewMember = New clsQueryItem

    If SearchType = AR_EXECUTEON Then
      lNumOfExecuteOn = lNumOfExecuteOn + 1
      AddExecuteOnValue (SearchValueString)
    End If

    'set the properties passed into the method
    objNewMember.SearchType = SearchType
    objNewMember.SearchParam = SearchParam
    objNewMember.SearchValueString = SearchValueString
    objNewMember.SearchValueNum = SearchValueNum
    objNewMember.SearchCondition = SearchCondition
    objNewMember.SearchConditionString = SearchConditionString
    objNewMember.tag = sKey
'    If IsObject(tag) Then
'        Set objNewMember.tag = tag
'    Else
'        objNewMember.tag = tag
'    End If
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As clsQueryItem
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
Dim TempItem As clsQueryItem
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)

    Set TempItem = mCol.Item(vntIndexKey)
    
    If TempItem.SearchType = AR_EXECUTEON Then
      lNumOfExecuteOn = lNumOfExecuteOn - 1
      SubtractExecuteOnValue (TempItem.SearchValueString)
    End If
    
    Set TempItem = Nothing

    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
    lNumOfExecuteOn = 0
    mvarExecuteOnValue = 0
    
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

