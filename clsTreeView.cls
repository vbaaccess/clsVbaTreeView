'--------------------------- --------------------------- ---------------------------
' Module  : clsTreeView
' Author  : paweld
' Date    : 2019-09-11
' Purpose : Universal form tree view, without ActiveX or other external controls
' Licence : Free
'
' Changes :
'--------------------------- --------------------------- ---------------------------
' 2019-09-11 , PD : Create
'--------------------------- --------------------------- ---------------------------
Option Compare Database
Option Explicit

Private Const CurrentModName = "clsTreeView"
Private Const TreeViewFormName = "frmTreeView"

Private objHandle                                               ' object handle
Public Event PrzypiszUchwytDoObiektu(uchwytDoObiektu As Long)   ' Event for catching object

Private m_bWindowModeDialog As Boolean
Private m_HeaderOfSelectionWindow As String
Private m_ArraySql As Variant
Private m_SqlWhereTag As String
Private m_selectedID As Long

Private Sub Class_Initialize()
    m_bWindowModeDialog = True
End Sub

Private Sub Class_Terminate()
    Debug.Print "Terminate (" & VBA.Now & ")"
End Sub

Public Sub INIT(Optional bDialog As Boolean = True)
    m_bWindowModeDialog = bDialog
  'init parameters
End Sub

Public Sub ShowTree
    Dim sOpenArgs
    Dim WindowMode As AcWindowMode
    
                                WindowMode = acWindowNormal
    If m_bWindowModeDialog Then WindowMode = acDialog

    sOpenArgs = m_HeaderOfSelectionWindow & ";" & m_SqlWhereTag & ";" & Join(m_ArraySql, ";")

    DoCmd.OpenForm TreeViewFormName, , , , , WindowMode, sOpenArgs
    
    m_selectedID = fTreeView_ReturnID
    fTreeView_ReturnID = 0
End Sub
