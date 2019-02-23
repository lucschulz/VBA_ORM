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
Private p_TableName As String
Private p_Where As String
Private p_GroupBy As String
Private p_OrderBy As String


''''''''''''
'CONSTRUCTOR
''''''''''''
Private Sub Class_Initialize()
    Set p_Columns = New Collection
End Sub


'''''''''''
'PROPERTIES
'''''''''''

'PARAM(Value): The name of the database table to use in the query.
Public Property Let fromTable(value As String)
    p_TableName = " FROM " & value
End Property

'PARAM(Value): The WHERE clause to be used in the query. Do not add the 'WHERE' keyword.
Public Property Let Where(value As String)
    p_Where = " WHERE " & value
End Property

Public Property Let GroupBy(value As String)
    p_GroupBy = " GROUP BY " & value
End Property

Public Property Let OrderBy(value As String)
    p_OrderBy = " ORDER BY " & value
End Property





'ADDS A COLUMN NAME TO THE QUERY
'PARAM (columnName): THE NAME OF THE COLUMN TO BE ADDED TO THE QUERY
Public Sub SelectColumn(columnName As String)
    p_Columns.Add columnName
End Sub


'''''''''''''
'SELECT QUERY
'''''''''''''
'RETURNS THE FORMATTED QUERY WITH ANY OPTIONAL CLAUSES ATTACHED
Public Function GetQuery() As String
    Dim selectQuery As String
    selectQuery = "SELECT " & Join(CollectionToArray(p_Columns), ", ") & p_TableName & p_Where & p_GroupBy & p_OrderBy
    
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
