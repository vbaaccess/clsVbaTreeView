'---------------------------------------------------------------------------------------
' Module    : clsTreeViewMain
' Author    : paweld
' Date      : 2019-09-11
' Purpose   : Universal form tree view, without ActiveX or other external controls
' Licence   : Free
'
' Changes:
'----------------------------------------------------------------------------------------------------------------------------------
' 2019-09-11        , PD : Create
' 2021-01-28        , PD : Refaktoring; modification of the management class
' 2021-02-15        , PD : Development; add option to call function from global object
'----------------------------------------------------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

Private Const CurrentModName = "clsTreeViewMain"
Private Const TreeViewFormName = "frmTreeView"

Private WithEvents objOknoTreeView As clsTreeViewWithEvent

Private Sub Class_Initialize()
    Set objOknoTreeView = New clsTreeViewWithEvent
    Call objOknoTreeView.Init

End Sub

Private Sub objOknoTreeView_PrzypiszUchwytDoObiektu(uchwytDoObiektu As Long)
    uchwytDoObiektu = ObjPtr(objOknoTreeView)
End Sub

Public Function Wybierz(mode As vbaTreeViewDefinitionEnum, Optional RefreshData As Boolean = True) As Long
    If RefreshData Then
        Call objOknoTreeView.RefreshData(mode)
    End If
    
    Wybierz = objOknoTreeView.Wybierz(mode)
End Function