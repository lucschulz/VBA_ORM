VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_QueryDesigner_Join"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private p_LeftTable As String
Private p_RightTable As String
Private p_LeftValue As String
Private p_RightValue As String
Private p_JoinType As String

Public Enum JoinTypeEnum
    Inner
    Left
    Right
End Enum


Public Property Let LeftTableName(value As String)
    p_LeftTable = value
End Property

Public Property Let RightTableName(value As String)
    p_RightTable = value
End Property

Public Property Let leftValue(value As String)
    p_LeftValue = value
End Property

Public Property Let rightValue(value As String)
    p_RightValue = value
End Property


Public Property Get JoinText() As String
    JoinText = "(" & p_LeftTable & p_JoinType & p_RightTable & " ON " & p_LeftValue & " = " & p_RightValue & ")"
End Property


Public Sub SetJoinType(jt As JoinTypeEnum)
    
    If jt = Inner Then
        p_JoinType = " INNER JOIN "
    
    ElseIf jt = Left Then
        p_JoinType = " LEFT JOIN "
    
    ElseIf jt = Right Then
        p_JoinType = " RIGHT JOIN "
    End If
    
End Sub
