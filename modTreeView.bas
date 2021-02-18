Option Compare Database
Option Explicit

Private Const CurrentModName = "modTreeView"

Public TreeView As New clsTreeViewMain

Public Enum vbaTreeViewDefinitionEnum
    Sample = 1
End Enum

Public Type vbaTreeViewProperty
  FormCaption As String
  WhereTag As String
  ArrraySql() As Variant
End Type