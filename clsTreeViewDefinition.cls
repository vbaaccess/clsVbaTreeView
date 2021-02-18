'---------------------------------------------------------------------------------------
' Module    : clsTreeViewDefinition
' Author    : paweld
' Date      : 2021-02-01
' Purpose   : Tree View class; tree definitions
' Licence   : Free
'
' Changes:
'----------------------------------------------------------------------------------------------------------------------------------
' 2021-02-01        , PD : Definition (Create some sample)
'----------------------------------------------------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

Private Const CurrentModName = "clsTreeViewDefinition"
Private Const DefaultWhereTag = "/*{{WHERE}}*/"
Private m_DefinitionStatus As Dictionary

Private Sub Class_Initialize()
    Set m_DefinitionStatus = New Dictionary
End Sub

Public Property Get Property(TreeViewKey As vbaTreeViewDefinitionEnum) As Object
    Dim cTV As New clsTreeViewProperty
    
    If m_DefinitionStatus.Exists(TreeViewKey) Then
        
        Set cTV = m_DefinitionStatus(TreeViewKey)
        
        Set Property = cTV
    Else
        Set Property = Nothing
    End If

End Property

Public Property Get DefinitionStatus(mode As vbaTreeViewDefinitionEnum) As Boolean
    DefinitionStatus = m_DefinitionStatus.Exists(mode)
End Property

Public Function InitDefinition(DefinitioTreeView As vbaTreeViewDefinitionEnum) As Boolean
    
    Select Case DefinitioTreeView
    Case vbaTreeViewDefinitionEnum.PlanuKont
        Call InitTreeView_PlanKont(DefinitioTreeView)
    Case Else
        Debug.Print "DEFINITION UNDER CONSTRUCTION  !"
    End Select
    
    InitDefinition = True
End Function

Private Function InitTreeView_Sample(TreeViewDefinitionKey As vbaTreeViewDefinitionEnum) As vbaTreeViewProperty
Const Level1 = 1
Const Level2 = 2
Const LevelN = 3

    Dim SqlLevel(1 To 3) As String
    Dim TV As vbaTreeViewProperty

    SqlLevel(Level1) = "SELECT LEVEL 1"
    
    SqlLevel(Level2) = "SELECT LEVEL 2"
    
    SqlLevel(LevelN) = "SELECT LEVEL 3"

    TV.ArrraySql = Array(SqlLevel(Level1), SqlLevel(Level2), SqlLevel(LevelN))
    
    '--- add definition ------------------------------------------------------------
    
    If m_DefinitionStatus.Exists(TreeViewDefinitionKey) Then
        Call m_DefinitionStatus.Remove(TreeViewDefinitionKey)
    End If

    Dim cTV As clsTreeViewProperty
    Set cTV = New clsTreeViewProperty

    cTV.FormCaption = TV.FormCaption
    cTV.WhereTag = TV.WhereTag
    cTV.ArrraySql(0) = TV.ArrraySql(0)
    cTV.ArrraySql(1) = TV.ArrraySql(1)
    cTV.ArrraySql(2) = TV.ArrraySql(2)

End Function