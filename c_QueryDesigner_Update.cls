VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_QueryDesigner_Update"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private p_Columns As Collection
Private p_TableName As String
Private p_Where As String
Private p_JoinExpression As String

''''''''''''
'CONSTRUCTOR
''''''''''''
Private Sub Class_Initialize()
    Set p_Columns = New Collection
End Sub

'''''''''''
'PROPERTIES
'''''''''''
'Value: The name of the database table to insert the values into.
Public Property Let TableName(value As String)
    p_TableName = value
End Property

'PARAM (columnName): THE NAME OF THE COLUMN TO BE ADDED TO THE QUERY
'PARAM (columnValue): THE MATCHING VALUE TO BE INSERTED/UPDATED
Public Sub SetColumn(columnName As String, equalsValue As String)
    p_Columns.Add columnName & " = " & equalsValue
End Sub

Public Sub SetColumnWithTablePrefix(TableName As String, columnName As String, equalsValue As String)
    p_Columns.Add TableName & "." & columnName & " = " & equalsValue
End Sub

Public Property Let whereClause(value As String)
    p_Where = value
End Property

Public Function GetUpdateCommand() As String    
    Dim update As String
    Dim table As String
    If p_JoinExpression <> vbNullString Then
        update = "UPDATE " & p_JoinExpression
    Else
        update = "UPDATE " & p_TableName
    End If
    update = update & " SET " & Join(CollectionToArray(p_Columns), ", ") & " " & p_Where
    GetUpdateCommand = update
End Function

'''''''''''''''''
'HELPER FUNCTIONS
'''''''''''''''''
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
