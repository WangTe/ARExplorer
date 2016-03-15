VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colQueryGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"UserDefinedCollection"
Attribute VB_Ext_KEY = "Member0" ,"colQueryList"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable to hold collection
Private mCol As Collection

'Public Property Let SavedID(ByVal vData As Long)
''used when assigning a value to the property, on the left side of an assignment.
''Syntax: X.SavedID = 5
'    mvarSavedID = vData
'End Property
'Public Property Get SavedID() As Long
''used when retrieving value of a property, on the right side of an assignment.
''Syntax: Debug.Print X.SavedID
'    SavedID = mvarSavedID
'End Property

'Public Sub ResetCollection()
'Dim i As Long
'
'  SearchType = ""
'  SearchServerName = ""
'  SaveName = "<EMPTY>"
'  SavedID = 0
'  SavePath = ""
'  Saved = False
'  Dirty = False
'  mvarExecuteOnValue = 0
'  lNumOfExecuteOn = 0
'  CaseSensitive = False
'  RemoveAll
'
'End Sub

Public Sub RemoveAll()
Dim i As Integer

  For i = 1 To Me.Count
    Me.Remove (i)
  Next i

  'ResetCollection
End Sub


Public Sub Add(ByRef colARQuery As colQueryList, Optional sKey As String)
'Public Function Add(SearchType As String, SearchParam As Long, SearchValueString As String, SearchValueNum As Long, SearchCondition As String, SearchConditionString As String, Optional sKey As String) As colQueryList
    'create a new object
'    Dim objNewMember As colQueryList
'
'    Set objNewMember = New colQueryList
'
'    'set the properties passed into the method
'    objNewMember.SearchType = SearchType
'    objNewMember.SearchParam = SearchParam
'    objNewMember.SearchValueString = SearchValueString
'    objNewMember.SearchValueNum = SearchValueNum
'    objNewMember.SearchCondition = SearchCondition
'    objNewMember.SearchConditionString = SearchConditionString
    If Len(sKey) = 0 Then
        mCol.Add colARQuery
    Else
        mCol.Add colARQuery, sKey
    End If
    
    'return the object created
'    Set Add = objNewMember
'    Set objNewMember = Nothing


End Sub

Public Property Get Item(vntIndexKey As Variant) As colQueryList
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
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)
    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
    
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub


