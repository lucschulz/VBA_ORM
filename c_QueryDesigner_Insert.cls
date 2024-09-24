VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_QueryDesigner_Insert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private p_Columns As Collection
Private p_Values As Collection
Private p_TableName As String

Private Sub Class_Initialize()
    Set p_Columns = New Collection
    Set p_Values = New Collection
End Sub

'Value: The name of the database table to insert the values into.
Public Property Let TableName(value As String)
    p_TableName = value
End Property

'PARAM (columnName): THE NAME OF THE COLUMN TO BE ADDED TO THE QUERY
'PARAM (columnValue): THE MATCHING VALUE TO BE INSERTED
Public Sub InsertColumn(columnName As String, columnValue As String)
    p_Columns.Add columnName
    p_Values.Add columnValue
End Sub

'RETURNS THE FORMATTED QUERY WITH ANY OPTIONAL CLAUSES ATTACHED
Public Function GetInsertCommandText() As String
    Dim insert As String
    insert = "INSERT INTO " & p_TableName & "(" & Join(CollectionToArray(p_Columns), ", ") & ")"
    insert = insert & " VALUES(" & Join(CollectionToArray(p_Values), ", ") & ");"    
    GetInsertCommandText = insert
End Function

'CONVERTS A COLLECTION INTO AN ARRAY
Private Function CollectionToArray(myCol As Collection) As Variant 
    Dim Result  As Variant
    Dim cnt     As Long 
    ReDim Result(myCol.Count - 1) 
    For cnt = 0 To myCol.Count - 1
        Result(cnt) = myCol(cnt + 1)
    Next cnt 
    CollectionToArray = Result    
End Function