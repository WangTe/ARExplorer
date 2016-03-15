VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsQueryItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "This class contains information for 1 query item, it will be contained within a collection to ease adding and removing of items.  Each class will contain 5 pieces of data used in the search:  SearchType, SearchParam, SearchValueString, SearchValueNum (only one of the SearchValue params can be used at once), SearchCondition.  This class will contain an index and a name as well."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarSearchType As String 'local copy
Private mvarSearchParam As Long 'local copy
Private mvarSearchValueString As String 'local copy
Private mvarSearchValueNum As Long 'local copy
Private mvarSearchCondition As String 'local copy
Private mvartag As Variant 'local copy
'local variable(s) to hold property value(s)
Private mvarSearchConditionValue As Long 'local copy
'local variable(s) to hold property value(s)
Private mvarSearchConditionString As String 'local copy
Public Property Let SearchConditionString(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SearchConditionString = 5
    mvarSearchConditionString = vData
End Property


Public Property Get SearchConditionString() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SearchConditionString
    SearchConditionString = mvarSearchConditionString
End Property



Public Property Let SearchConditionValue(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SearchConditionValue = 5
    mvarSearchConditionValue = vData
End Property


Public Property Get SearchConditionValue() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SearchConditionValue
    SearchConditionValue = mvarSearchConditionValue
End Property



Public Property Let tag(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.tag = 5
    mvartag = vData
End Property


Public Property Set tag(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.tag = Form1
    Set mvartag = vData
End Property


Public Property Get tag() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.tag
    If IsObject(mvartag) Then
        Set tag = mvartag
    Else
        tag = mvartag
    End If
End Property



Public Property Let SearchCondition(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SearchCondition = 5
    mvarSearchCondition = vData
End Property


Public Property Get SearchCondition() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SearchCondition
    SearchCondition = mvarSearchCondition
End Property



Public Property Let SearchValueNum(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SearchValueNum = 5
    mvarSearchValueNum = vData
End Property


Public Property Get SearchValueNum() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SearchValueNum
    SearchValueNum = mvarSearchValueNum
End Property



Public Property Let SearchValueString(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SearchValueString = 5
    mvarSearchValueString = vData
End Property


Public Property Get SearchValueString() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SearchValueString
    SearchValueString = mvarSearchValueString
End Property



Public Property Let SearchParam(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SearchParam = 5
    mvarSearchParam = vData
End Property


Public Property Get SearchParam() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SearchParam
    SearchParam = mvarSearchParam
End Property



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



