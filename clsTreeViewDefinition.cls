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
Private m_DefinitionStatus As Dictionary

Private Sub Class_Initialize()
    Set m_DefinitionStatus = New Dictionary
End Sub

Private Function InitTreeView_Sample(TreeViewDefinitionKey As vbaTreeViewWybor) As TreeViewProperty
Const Level1 = 1
Const Level2 = 2
Const LevelN = 3

    Dim SqlLevel(1 To 3) As String
    Dim TV As TreeViewProperty

    SqlLevel(Level1) = "SELECT LEVEL 1"
    
    SqlLevel(Level2) = "SELECT LEVEL 2"
    
    SqlLevel(Level3) = "SELECT LEVEL 3"

    TV.ArrraySql = Array(SqlLevel(Grupa), SqlLevel(OBSZAR), SqlLevel(PLAN_KONT))
    
    Call UpdateDictionaryDefinition(TreeViewDefinitionKey, TV)

End Function

Private Sub UpdateDictionaryDefinition(TreeViewDefinitionKey As vbaTreeViewWybor, TreeViewDefinitionProperty As TreeViewProperty)
    
    If m_DefinitionStatus.Exists(TreeViewDefinitionKey) Then
        Call m_DefinitionStatus.Remove(TreeViewDefinitionKey)
    End If
    
    Call m_DefinitionStatus.Add(TreeViewDefinitionKey, TreeViewDefinitionProperty)
End Sub