Option Compare Database
Option Explicit

Private Const CurrentModName = "modTreeView"

Public TreeView As New clsTreeViewMain

Public Enum vbaTreeViewWybor
    LokacjiZakladu = 1
    LokacjiWydzial = 2
    LokacjiRejonu = 3
    Maszyny = 4
    PlanuKont = 5
End Enum

Private Type TreeViewProperty
  FormCaption As String
  WhereTag As String
  ArrraySql() As Variant
End Type