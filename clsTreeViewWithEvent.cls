'---------------------------------------------------------------------------------------
' Module    : clsTreeViewWithEvent
' Author    : paweld
' Date      : 2021-01-28
' Purpose   : Tree View class to help manage events
' Licence   : Free
'
' Changes:
'----------------------------------------------------------------------------------------------------------------------------------
' 2021-01-28        , PD : Extract Events from main class (create class)
'----------------------------------------------------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

Private Const CurrentModName = "clsTreeViewWithEvent"
Private Const TreeViewFormName = "frmTreeView"

Private m_ReturnID As Long

'--- klasa z deinifjami TreeView ---
Private objTreeViewDefinition As New clsTreeViewDefinition

'--- do wykorzystania w formulazu z drzewkiem ---
Private m_FormCaption As String
Private m_FormWhereTag As String
Private m_TrybTreeView As vbaTreeViewDefinitionEnum
'--- ---------------------------------------- ---

Private objHandle                                               ' uchwyt do obiektu
Public Event PrzypiszUchwytDoObiektu(uchwytDoObiektu As Long)   ' do samego siebie

Public Sub RefreshData(mode As vbaTreeViewDefinitionEnum)
    Call objTreeViewDefinition.InitDefinition(mode)
End Sub

Public Function Wybierz(mode As vbaTreeViewDefinitionEnum) As Long
    m_TrybTreeView = mode
    DoCmd.OpenForm TreeViewFormName, , , , , acDialog, CStr(objHandle)
    Wybierz = m_ReturnID
End Function

Public Function Property() As Object
    Dim vProperty As New clsTreeViewProperty
    
    Set vProperty = objTreeViewDefinition.Property(m_TrybTreeView)

    Set Property = vProperty
    
End Function

Public Sub Init()
    Dim l_objHandle As Long
    
    If objHandle = 0 Then
        RaiseEvent PrzypiszUchwytDoObiektu(l_objHandle)
        objHandle = CLng(l_objHandle)
    End If
End Sub

Public Property Get ReturnID() As Long
    ReturnID = m_ReturnID
End Property

Public Property Let ReturnID(NewID As Long)
    m_ReturnID = NewID
    fTreeView_ReturnID = NewID ' comaptybile with old wersion
End Property

Public Property Get FormCaption() As String
    FormCaption = m_FormCaption
End Property

Public Property Let FormCaption(NewFormCaption As String)
    m_FormCaption = NewFormCaption
End Property

Public Property Get FormWhereTag() As String
    m_FormWhereTag = m_FormWhereTag
End Property

Public Property Let FormWhereTag(NewFormWhereTag As String)
    m_FormWhereTag = NewFormWhereTag
End Property
