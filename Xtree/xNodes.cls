VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "xNodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarNodes As Collection
Private mvarTotalNodeCount As Long
Private m_lOwner   As Long
Private m_hWnd     As Long
Private Declare Function IsWindow Lib "user32" (ByVal HWND As Long) As Long

Private m_Tree As xTree



Friend Sub Init(ByVal lPtr As Long, _
                ByVal lHwnd As Long)

   m_lOwner = lPtr
   m_hWnd = lHwnd

End Sub
Private Property Get TreeObject() As xTree
   Dim oT As Object

   If IsWindow(m_hWnd) Then
      If Not (m_lOwner = 0) Then
         CopyMemory oT, m_lOwner, 4
         Set TreeObject = oT
         CopyMemory oT, 0&, 4
      End If
   End If

End Property

Friend Property Let TotalNodeCount(ByVal vData As Long)
    mvarTotalNodeCount = vData
End Property


Friend Property Get TotalNodeCount() As Long
    TotalNodeCount = mvarTotalNodeCount
End Property

Public Property Set Nodes(ByVal vData As Collection)
    Set mvarNodes = vData
End Property


Public Property Get Nodes() As Collection
    Set Nodes = mvarNodes
End Property


'//---------------------------------------------------------------------------------------
' Procedure : AddNode
' Type      : Function
' DateTime  : 04/10/2004 15:53
' Author    : Gary Noble
' Purpose   : Adds A Node To The Collection
' Returns   : xTreeNode
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  04/10/2004
'//---------------------------------------------------------------------------------------
Public Function AddNode(ByVal Caption As String, ByVal Key As String, Optional ByVal IconIndex As Long, Optional ByRef Parent As xTreeNode) As xTreeNode
On Error GoTo errAdd

    Dim x As xTreeNode

    Set x = New xTreeNode
    Dim xParent As xTreeNode

    '-- bail
    If mvarNodes Is Nothing Then Set mvarNodes = New Collection

    '-- If The Parent is Not Nothing The We Add To The Parent Collection
    If Not Parent Is Nothing Then
        rCurNode(ObjPtr(Parent)).ChildCount = rCurNode(ObjPtr(Parent)).ChildCount + 1
        Set x = Parent.Children.AddNode(Caption, Key, IconIndex)
        x.Level = Parent.Level + 1
        x.IconIndex = IconIndex
        x.ParentPTR = ObjPtr(Parent)

    Else
        '-- Add new Top Level Node
        With x
            .Caption = Caption
            .IconIndex = IconIndex
            
            '-- Add It To The collection
            mvarNodes.Add x, Key
        End With
    End If

    TotalNodeCount = TotalNodeCount + 1

    x.Key = Key
    CopyMemory AddNode, x, 4
    CopyMemory AddNode, 0&, 4
    
CleanExit:
    
    
    '-- Return the node
    Set AddNode = x

Exit Function

errAdd:
    
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    Set x = Nothing
    Resume CleanExit
    
End Function

Private Sub Class_Initialize()
    Set mvarNodes = New Collection
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Set mvarNodes = Nothing
End Sub


'//---------------------------------------------------------------------------------------
' Procedure : rCurNode
' Type      : Property
' DateTime  : 04/10/2004 15:55
' Author    : Gary Noble
' Purpose   : Returns A Node Object From its Pointer
' Returns   : xTreeNode
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  04/10/2004
'//---------------------------------------------------------------------------------------
Friend Property Get rCurNode(ptr As Long) As xTreeNode
    Dim xNode As xTreeNode

    CopyMemory xNode, ptr, 4

    Set rCurNode = xNode

    CopyMemory xNode, 0&, 4

    ZeroMemory ObjPtr(xNode), 4&

End Property

