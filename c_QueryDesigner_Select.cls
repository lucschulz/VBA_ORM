VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_QueryDesigner_Select"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private p_Columns As Collection
Private p_GroupByItems As Collection
Private p_OrderByItems As Collection
Private p_Select As String
Private p_TableName As String
Private p_Where As String
Private p_GroupBy As String
Private p_OrderBy As String

Private Sub Class_Initialize()
    Set p_Columns = New Collection
End Sub

'PARAM(Value): The name of the database table to use in the query.
Public Property Let fromTable(value As String)
    p_TableName = " FROM " & value
End Property

'PARAM(Value): The WHERE clause to be used in the query. Do not add the 'WHERE' keyword.
Public Property Let Where(value As String)
    p_Where = " WHERE " & value
End Property

Public Function SelectCol(ParamArray columns() As Variant)    
    Dim v As Variant
    For Each v In columns
        p_Columns.Add v
    Next
    p_Select = " SELECT " & Join(CollectionToArray(p_Columns), ", ")
End Function

Public Function GroupBy(ParamArray columns() As Variant)    
    Set p_GroupByItems = New Collection    
    Dim v As Variant
    For Each v In columns
        p_GroupByItems.Add v
    Next
    If p_GroupByItems.Count <> 0 Then
        p_GroupBy = " GROUP BY " & Join(CollectionToArray(p_GroupByItems), ", ")
    Else
        p_GroupBy = vbNullString
    End If
End Function

Public Function OrderBy(ParamArray columns() As Variant)    
    Set p_OrderByItems = New Collection    
    Dim v As Variant
    For Each v In columns
        p_OrderByItems.Add v
    Next    
    If p_OrderByItems.Count <> 0 Then
        p_OrderBy = " ORDER BY " & Join(CollectionToArray(p_OrderByItems), ", ")
    Else
        p_OrderBy = vbNullString
    End If    
End Function

''ADDS A COLUMN NAME TO THE QUERY
''PARAM (columnName): THE NAME OF THE COLUMN TO BE ADDED TO THE QUERY
'Public Sub SelectColumn(columnName As String)
'    p_Columns.Add columnName
'End Sub

'''''''''''''
'SELECT QUERY
'''''''''''''
'RETURNS THE FORMATTED QUERY WITH ANY OPTIONAL CLAUSES ATTACHED
Public Function GetQuery() As String
    Dim selectQuery As String
    selectQuery = p_Select & p_TableName & p_Where & p_GroupBy & p_OrderBy
    GetQuery = selectQuery
End Function

'''''''''''''''''
'HELPER FUNCTIONS
'''''''''''''''''
'CONVERTS A COLLECTION INTO AN ARRAY
Public Function CollectionToArray(myCol As Collection) As Variant 
    Dim Result  As Variant
    Dim cnt     As Long 
    ReDim Result(myCol.Count - 1) 
    For cnt = 0 To myCol.Count - 1
        Result(cnt) = myCol(cnt + 1)
    Next cnt 
    CollectionToArray = Result
End Function
