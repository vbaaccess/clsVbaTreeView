'---------------------------------------------------------------------------------------
' Module    : clsTreeViewProperty
' Author    : paweld
' Date      : 2021-02-16
' Purpose   : Instead of a User Definition Type
' Licence   : Free
'
' Changes:
'----------------------------------------------------------------------------------------------------------------------------------
' 2021-02-16        , PD : Create
'----------------------------------------------------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

Private Const CurrentModName = "clsTreeViewProperty"

Private m_FormCaption As String
Private m_WhereTag As String
Private m_ArrraySql() As String

Private Sub Class_Initialize()
    ReDim m_ArrraySql(0)
End Sub

Public Property Let FormCaption(strValue As String)
    m_FormCaption = strValue
End Property
Public Property Get FormCaption() As String
    FormCaption = m_FormCaption
End Property

Public Property Let WhereTag(strValue As String)
    m_WhereTag = strValue
End Property
Public Property Get WhereTag() As String
    WhereTag = m_WhereTag
End Property

Public Property Get ArrraySql(Index As Long) As String
    ArrraySql = m_ArrraySql(Index)
End Property

Public Property Let ArrraySql(Index As Long, NewSqlString As String)
    If Len(Trim(NewSqlString)) > 0 Then
        If Index > UBound(m_ArrraySql) Then
            ReDim Preserve m_ArrraySql(Index)
        End If
        
        m_ArrraySql(Index) = NewSqlString
    End If
End Property

Public Property Get AllSql() As String()
    AllSql = m_ArrraySql()
End Property