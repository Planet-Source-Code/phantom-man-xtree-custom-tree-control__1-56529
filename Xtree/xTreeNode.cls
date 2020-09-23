VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "xTreeNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'//---------------------------------------------------------------------------------------
'xTreeX
'//---------------------------------------------------------------------------------------
' Module    : xTreeNode
' DateTime  : 04/10/2004 15:51
' Author    : Gary Noble
' Purpose   : Node Object
' Assumes   : xNodes.cls
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  04/10/2004
'//---------------------------------------------------------------------------------------
Option Explicit


Private mvarKey As String
Private mvarCaption As String
Private mvarParent As xTreeNode
Private mvarChildCount As Long
Private mvarChildren As xNodes
Private mChhildren As xNodes
Private mvarExpanded As Boolean
Private mvarLevel As Long
Private mvarRectData As RECT
Private mvarBold As Boolean
Private mvarID As Long
Private mvarNodeFont As StdFont
Private mvarCheVronRect As RECT
Private mvarForeColor As OLE_COLOR
Private mvarItemData As String
Private mvarItemDataColor As OLE_COLOR
Private mvarItemDataBold As Boolean
Private mvarParentPTR As Long
'local variable(s) to hold property value(s)
Private mvarIconIndex As Long 'local copy
Public Property Let IconIndex(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IconIndex = 5
    mvarIconIndex = vData
End Property


Public Property Get IconIndex() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IconIndex
    IconIndex = mvarIconIndex
End Property



Friend Property Let ParentPTR(ByVal vData As Long)
    mvarParentPTR = vData
End Property


Friend Property Get ParentPTR() As Long
    ParentPTR = mvarParentPTR
End Property

Public Property Let ItemDataBold(ByVal vData As Boolean)
    mvarItemDataBold = vData
End Property

Public Property Get ItemDataBold() As Boolean
    ItemDataBold = mvarItemDataBold
End Property

Public Property Let ItemDataColor(ByVal vData As OLE_COLOR)
    Let mvarItemDataColor = vData
End Property
Public Property Get ItemDataColor() As OLE_COLOR
    Let ItemDataColor = mvarItemDataColor
End Property

Public Property Let ItemData(ByVal vData As String)
    mvarItemData = vData
End Property


Public Property Get ItemData() As String
    ItemData = mvarItemData
End Property



Public Property Let ForeColor(ByVal vData As OLE_COLOR)
    mvarForeColor = vData
End Property


Public Property Get ForeColor() As OLE_COLOR
    ForeColor = mvarForeColor
End Property

Friend Property Get CheVronRect() As RECT
    Let CheVronRect = mvarCheVronRect
End Property


Public Property Set NodeFont(ByVal vData As StdFont)
    Set mvarNodeFont = vData
End Property


Public Property Get NodeFont() As StdFont
    Set NodeFont = mvarNodeFont
End Property

Friend Property Let id(ByVal vData As Long)
    mvarID = vData
End Property


Friend Property Get id() As Long
    id = mvarID
End Property

Public Property Let Bold(ByVal vData As Boolean)
    mvarBold = vData
End Property


Public Property Get Bold() As Boolean
    Bold = mvarBold
End Property


Friend Sub SetRCChevRect(x As Long, y As Long, right As Long, bottom As Long)
    mvarCheVronRect.top = y
    mvarCheVronRect.left = x
    mvarCheVronRect.right = right
    mvarCheVronRect.bottom = bottom

End Sub

Friend Sub SetRect(x As Long, y As Long, right As Long, bottom As Long)
    mvarRectData.top = y
    mvarRectData.left = x
    mvarRectData.right = right
    mvarRectData.bottom = bottom

End Sub


Friend Property Get RectData() As RECT
    Let RectData = mvarRectData
End Property



Friend Property Let Level(ByVal vData As Long)
    mvarLevel = vData
End Property


Friend Property Get Level() As Long
    Level = mvarLevel
End Property



Public Property Let Expanded(ByVal vData As Boolean)
    mvarExpanded = vData
End Property


Public Property Get Expanded() As Boolean
    Let Expanded = mvarExpanded
End Property




Public Property Get Children() As xNodes
    If mvarChildren Is Nothing Then Set mvarChildren = New xNodes

    If IsObject(mvarChildren) Then
        Set Children = mvarChildren
    Else
        Set Children = mvarChildren
    End If
End Property



Public Property Let ChildCount(ByVal vData As Long)
    mvarChildCount = vData
End Property


Public Property Get ChildCount() As Long
    ChildCount = mvarChildCount
End Property



Public Property Set Parent(ByVal vData As xTreeNode)
    Set mvarParent = vData
End Property


Public Property Get Parent() As xTreeNode
    Set Parent = mvarParent
End Property



Public Property Let Caption(ByVal vData As String)
    mvarCaption = vData
End Property


Public Property Get Caption() As String
    Caption = mvarCaption
End Property



Public Property Let Key(ByVal vData As String)
    mvarKey = vData
End Property


Public Property Get Key() As String
    Let Key = mvarKey
End Property



Private Sub Class_Initialize()
    Set mvarChildren = New xNodes


End Sub

Private Sub Class_Terminate()
    Set mvarChildren = Nothing
End Sub
