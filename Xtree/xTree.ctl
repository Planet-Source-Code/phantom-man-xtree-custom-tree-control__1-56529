VERSION 5.00
Begin VB.UserControl xTree 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4155
   ForeColor       =   &H000080FF&
   ForwardFocus    =   -1  'True
   KeyPreview      =   -1  'True
   ScaleHeight     =   386
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   277
End
Attribute VB_Name = "xTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'//---------------------------------------------------------------------------------------
'xTreeX
'//---------------------------------------------------------------------------------------
' Module    : xTree
' DateTime  : 04/10/2004 15:27
' Author    : Gary Noble
' Purpose   : Simple Tree Control
' Assumes   : You Have A Brain!!!
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  04/10/2004
'             Beta 0.000001.
'//---------------------------------------------------------------------------------------
Option Explicit
Option Base 0

Private m_hIml                               As Long
Private m_lptrVb6ImageList                   As Long
Private m_lIconWidth                         As Long
Private m_lIconHeight                        As Long
Private m_vImageList                         As Variant
Private playImage                            As Long
Private sourceWidth                          As Long
Private sourceHeight                         As Long
Dim m_bLostFocus                             As Boolean

'-- Used For Holding The Text Widths For The Horizontal Scrollbar
Dim marr_Widths() As Long

'-- Are We Drawing The Seleted Node
Private m_bDrawingSelectedNode As Boolean

'-- Lock Redraw Calls
Private Const WM_SETREDRAW = &HB
Private Declare Function SendMessage Lib "user32" Alias _
        "SendMessageA" (ByVal HWND As Long, ByVal wMsg As _
        Long, ByVal wParam As Long, lParam As Any) As Long

'-- Background Picture Painter
Private m_Tiler As IAPP_BitMapTiler

Private bDrawingSelItem As Boolean
Private lLocate As Long
Private bEnsure As Boolean
Private lL As Long
Private llastdrawn As Long
Private lFirstdrawn As Long
Private lButton As Long
Private lID As Long
Private bFound As Boolean
Private m_SelItem As String
Private mB_chev As Boolean
Private m_SelectedID As Long

'-- Api Scrollbars
'-- This Was Taken From One Of My projects - No Need For Manifest File
'-- This Class Will Auto Draw The Scrollbar Depending On the Theme
Private WithEvents m_cScrollBar              As IAPP_ScrollBars
Attribute m_cScrollBar.VB_VarHelpID = -1

Dim m_SelItemHover As String

'-- Search Node Item
Dim XSearchNode As xTreeNode
'-- Selected node Item
Dim mSelectedNode As xTreeNode

'-- Copy Memory For Ptr
#If Win32 Then
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
            lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
#Else
    Private Declare Sub CopyMemory Lib "KERNEL" Alias "hmemcpy" ( _
            lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
#End If

'-- Node Collection
Dim m_Nodes As xNodes

'-- Node Pointers Used For Painting Display Nodes
Dim m_NodesS() As Long    'As xTreeNode

'-- x and Y Co-Ordinates
Dim msng_X As Single, msng_y As Single

'Default Property Values:
Const m_def_ShowHorizontalScrollbar = False
'Const m_def_NodeChange = 0
'Const m_def_Cleared = 0
Const m_def_LostFocusSelectedBackColor = vbButtonFace
Const m_def_DisplayHorizontalScrollbar = False
Const m_def_SelectedColorBorder = vbBlack
Const m_def_SelectedColor = vbWindowText
Const m_def_SelectedBackColorTwo = vbHighlight
Const m_def_FullRowSelect = False
Const m_def_BackGradientOne = vbWhite
Const m_def_BackGradientTwo = vbWhite
Const m_def_SelectedBackColor = vbHighlight

'Property Variables:
Dim m_ShowHorizontalScrollbar As Variant
'Dim m_NodeChange As Variant
'Dim m_Cleared As Variant
Dim m_LostFocusSelectedBackColor As OLE_COLOR
Dim m_DisplayHorizontalScrollbar As Boolean
Dim m_BackGroundPicture As Picture
Dim m_SelectedColorBorder As OLE_COLOR
Dim m_SelectedColor As OLE_COLOR
Dim m_SelectedBackColorTwo As OLE_COLOR
Dim m_FullRowSelect As Boolean
Dim m_BackGradientOne As OLE_COLOR
Dim m_BackGradientTwo As OLE_COLOR
Dim m_SelectedBackColor As OLE_COLOR
'Event Declarations:
Event BeforeExpand(xNode As xTreeNode, bExpanding As Boolean)
Event NodeChange(xNode As xTreeNode)
Event Cleared()
Event NodeSelected(xNode As xTreeNode)
Event BeforeNodeChange(xNode As xTreeNode)



Public Property Get BackGradientOne() As OLE_COLOR
    BackGradientOne = m_BackGradientOne
End Property

Public Property Let BackGradientOne(ByVal New_BackGradientOne As OLE_COLOR)
    m_BackGradientOne = New_BackGradientOne
    PropertyChanged "BackGradientOne"
    Redraw
    Refresh
End Property
Public Property Get BackGradientTwo() As OLE_COLOR
    BackGradientTwo = m_BackGradientTwo
End Property

Public Property Let BackGradientTwo(ByVal New_BackGradientTwo As OLE_COLOR)
    m_BackGradientTwo = New_BackGradientTwo
    PropertyChanged "BackGradientTwo"
    Redraw
    Refresh
End Property

Public Property Get BackGroundPicture() As Picture
    Set BackGroundPicture = m_BackGroundPicture
End Property

Public Property Set BackGroundPicture(ByVal New_BackGroundPicture As Picture)
    Set m_BackGroundPicture = New_BackGroundPicture
    PropertyChanged "BackGroundPicture"
    'If New_BackGroundPicture Is Nothing Then Exit Property
    Redraw
    Refresh
End Property


'//---------------------------------------------------------------------------------------
' Procedure : Clear
' Type      : Function
' DateTime  : 04/10/2004 15:28
' Author    : Gary Noble
' Purpose   : Clears The Nodes
' Returns   : Variant
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  04/10/2004
'//---------------------------------------------------------------------------------------
Public Function Clear()
    On Error Resume Next


    '-- Erase Everything And Redraw The control
    Set m_Nodes = Nothing
    Set m_Nodes = New xNodes

    Set mSelectedNode = Nothing
    m_SelItem = ""

    '-- Hide The ScrollBars
    m_cScrollBar.Visible(efsHorizontal) = False
    m_cScrollBar.Visible(efsVertical) = False

    Erase m_NodesS()
    Erase marr_Widths()
        
    RaiseEvent Cleared
    
    Me.RefreshData
    Me.Redraw


    On Error GoTo 0
End Function

'//---------------------------------------------------------------------------------------
' Procedure : DeleteNode
' Type      : Function
' DateTime  : 04/10/2004 15:29
' Author    : Gary Noble
' Purpose   : Deletes A Given Node
' Returns   : Variant
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  04/10/2004
'//---------------------------------------------------------------------------------------
Public Function DeleteNode(xNode As xTreeNode)
    On Error Resume Next

    Dim xParent As xTreeNode
    Dim bSelected As Boolean
    
    If xNode.Key = mSelectedNode.Key Then bSelected = True
    
    If Not rCurNode(xNode.ParentPTR) Is Nothing Then
        Set xParent = rCurNode(xNode.ParentPTR)
        If Not xParent Is Nothing Then
            xParent.Children.Nodes.Remove xNode.Key
            Set mSelectedNode = xParent
            xParent.ChildCount = xParent.ChildCount - 1
            m_SelItem = xParent.Key
        Else
            m_Nodes.Nodes.Remove xNode.Key
            Set mSelectedNode = rCurNode(m_Nodes(1))
            m_SelItem = mSelectedNode.Key
        End If
    Else
        m_Nodes.Nodes.Remove xNode.Key
        pvPutNodesToArray
        If LBound(m_NodesS) > 0 Then
            Set mSelectedNode = rCurNode(m_NodesS(LBound(m_NodesS)))
            m_SelItem = mSelectedNode.Key
        Else
            Set mSelectedNode = Nothing
            m_SelItem = ""
        End If
        Me.RefreshData
    End If

    Me.RefreshData
    Refresh
    Me.Redraw

    Me.Selectednode.Expanded = Me.Selectednode.Expanded

On Error GoTo 0
End Function

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
    Redraw
End Property


'//---------------------------------------------------------------------------------------
' Procedure : EnsureVisible
' Type      : Function
' DateTime  : 04/10/2004 15:30
' Author    : Gary Noble
' Purpose   : Like The Treeview Ensure Selected Property
' Returns   : Variant
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  04/10/2004
'//---------------------------------------------------------------------------------------
Public Function EnsureVisible(xNode As xTreeNode)

    On Error Resume Next
    Dim fNode As xTreeNode
    Dim i As Long
    Dim s As xTreeNode
    Dim Parent As xTreeNode
    
    If xNode Is Nothing Then Exit Function

    '-- Set The Drawing Mode
    m_bDrawingSelectedNode = False

    If xNode.Level = 0 Then GoTo sExit

    Set fNode = xNode

    '-- Loop Throught The parent Nodes And Make Sure That It's Expanded
    Set Parent = rCurNode(xNode.ParentPTR)
    For Each fNode In Parent.Children.Nodes
        
        If xNode.Level = 0 Then GoTo sExit
        
        If Not xNode.Expanded Then xNode.Expanded = True

        If Not rCurNode(xNode.ParentPTR) Is Nothing Then
            
            '-- Do The Same For Each Level
            For Each s In rCurNode(xNode.ParentPTR).Children.Nodes
                rCurNode(xNode.ParentPTR).Expanded = True
                If s.Key = xNode.Key Then xEnsureVisible xNode: GoTo sExit
            Next

        Else
            xNode.Expanded = True
        End If
    
    Next
sExit:

    Set mSelectedNode = xNode
    m_SelItem = xNode.Key

    pvPutNodesToArray
    Me.RefreshData
    
'    Locate xNode
    
    RaiseEvent NodeSelected(xNode)
    
    
End Function


'//---------------------------------------------------------------------------------------
' Procedure : ExpandAllNodes
' Type      : Sub
' DateTime  : 04/10/2004 15:34
' Author    : Gary Noble
' Purpose   : As It Says, Expands All The Nodes
' Returns   :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  04/10/2004
'//---------------------------------------------------------------------------------------
Public Sub ExpandAllNodes(bln As Boolean)
    pvExpand bln
    Me.RefreshData
End Sub

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"

End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Get FullRowSelect() As Boolean
    FullRowSelect = m_FullRowSelect
End Property

Public Property Let FullRowSelect(ByVal New_FullRowSelect As Boolean)
    m_FullRowSelect = New_FullRowSelect
    PropertyChanged "FullRowSelect"
    Redraw
End Property


'//---------------------------------------------------------------------------------------
' Procedure : GetNode
' Type      : Function
' DateTime  : 04/10/2004 15:34
' Author    : Gary Noble
' Purpose   : Returns A Node by The Key
' Returns   : xTreeNode
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  04/10/2004
'//---------------------------------------------------------------------------------------
Public Function GetNode(strKey As String) As xTreeNode
    bFound = False
    pvGetNode strKey
    Set XSearchNode = Nothing
    Set GetNode = XSearchNode
End Function


'//---------------------------------------------------------------------------------------
' Procedure : GetNodebyCaption
' Type      : Function
' DateTime  : 04/10/2004 15:35
' Author    : Gary Noble
' Purpose   : Returns A node By The Caption
' Returns   : xTreeNode
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  04/10/2004
'//---------------------------------------------------------------------------------------
Public Function GetNodebyCaption(ByVal Caption As String) As xTreeNode
    On Error Resume Next
    bFound = False
    pvGetNodeByCaption Caption
    Set GetNodebyCaption = XSearchNode
    On Error GoTo 0
End Function


'//---------------------------------------------------------------------------------------
' Procedure : HitTest
' Type      : Function
' DateTime  : 04/10/2004 15:35
' Author    : Gary Noble
' Purpose   : Private Hitest
' Returns   : xTreeNode
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  04/10/2004
'//---------------------------------------------------------------------------------------
Private Function HitTest(x As Long, y As Long, Optional isChevron As Boolean, Optional IsHover As Boolean = False) As xTreeNode
    On Error Resume Next

    Dim i As Long
    Dim CurNode As xTreeNode
    Dim r As RECT
    mB_chev = False

    For i = m_cScrollBar.Value(efsVertical) + 1 To m_cScrollBar.Value(efsVertical) + Round(((ScaleHeight / (TextHeight("Q")))))
        Set CurNode = rCurNode(m_NodesS(i))
        LSet r = CurNode.RectData
        r.left = IIf(CurNode.ChildCount > 0, 0, m_lIconWidth) + IIf(Me.FullRowSelect, 0, (r.left + -m_cScrollBar.Value(efsHorizontal)))
        r.left = r.left - (m_lIconWidth)
        r.right = IIf(Me.FullRowSelect, ScaleWidth, r.right + -m_cScrollBar.Value(efsHorizontal))
        If PtInRect(r, x, y) Then
            If CurNode.ChildCount > 0 Then
            If x > CurNode.RectData.left - (m_lIconWidth + 2) And x < (CurNode.RectData.left - (m_lIconWidth + 2)) + (12) Then
                mB_chev = True
                Set HitTest = CurNode
                GoTo CleanExit
            End If
            End If
            If Not IsHover Then
                m_SelItem = CurNode.Key
                Set HitTest = CurNode
                If CurNode.Key <> mSelectedNode.Key Then RaiseEvent BeforeNodeChange(mSelectedNode)
                
                Set mSelectedNode = HitTest
                 RaiseEvent NodeSelected(HitTest)
                m_SelectedID = i
                
            Else
                m_SelItemHover = CurNode.Key
                'Set HitTest = CurNode
                m_SelectedID = i
            End If

            Exit For
        End If
    Next


CleanExit:
On Error GoTo 0
End Function

'//---------------------------------------------------------------------------------------
' Procedure : Locate
' Type      : Sub
' DateTime  : 04/10/2004 15:16
' Author    : Gary Noble
' Purpose   : Locates And Displays the Required Node
' Returns   :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  04/10/2004
'//---------------------------------------------------------------------------------------
Private Sub Locate(xNode As xTreeNode)

' NOT QUIT WORKING YET SO I HAVE LEFT IT OUT!!
Exit Sub


    If xNode Is Nothing Then Exit Sub

    '-- Lock The Screen For Updates
    SendMessage HWND, WM_SETREDRAW, 0, 0

    '-- Reset The Scrollbar
    m_cScrollBar.Value(efsVertical) = 0

    '-- Loop Through The Nodes And Wait until The seleted node Has Been Drawn
    Do Until m_bDrawingSelectedNode
        Exit Do
        DoEvents
        bEnsure = False
        '-- just in Case ( We Don't Want To Get In To A Continuous Loop!)
        If llastdrawn >= UBound(m_NodesS) Then Exit Do
        If m_bDrawingSelectedNode Then Exit Do
        m_cScrollBar.Value(efsVertical) = m_cScrollBar.Value(efsVertical) + ((llastdrawn - lFirstdrawn))
    Loop
    
    '-- Reset
    m_bDrawingSelectedNode = False
    bEnsure = False
    
    '-- Redraw The Control
    SendMessage HWND, WM_SETREDRAW, 1, 0
    Redraw
    
End Sub

Private Sub m_cScrollBar_Change(eBar As EFSScrollBarConstants)
    DoEvents
    Redraw
End Sub

Private Sub m_cScrollBar_MouseWheel(eBar As EFSScrollBarConstants, lAmount As Long)
    DoEvents
    Redraw
End Sub

Private Sub m_cScrollBar_Scroll(eBar As EFSScrollBarConstants)
    DoEvents
    Redraw
End Sub

Private Sub m_cScrollBar_ScrollClick(eBar As EFSScrollBarConstants, eButton As MouseButtonConstants)
    DoEvents
    Redraw
End Sub

'//---------------------------------------------------------------------------------------
' Procedure : MakeScrollVisible
' Type      : Sub
' DateTime  : 04/10/2004 15:36
' Author    : Gary Noble
' Purpose   : Displays Or Hides The Scrollbars
' Returns   :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  04/10/2004
'//---------------------------------------------------------------------------------------
Friend Sub MakeScrollVisible()
    Dim bShown As Boolean
    Dim i As Long
    Dim xNode As xTreeNode
    Dim sLongestNode As String


    On Error Resume Next
    Dim tHeight As Long
    Dim xHeight As Long
    
    m_cScrollBar.Max(efsVertical) = (UBound(m_NodesS)) - ((ScaleHeight \ (m_lIconWidth + 2)))
    m_cScrollBar.Visible(efsVertical) = m_cScrollBar.Max(efsVertical) > 1

    If Not m_cScrollBar.Visible(efsVertical) Then
        m_cScrollBar.Value(efsVertical) = 0
        m_cScrollBar.Max(efsVertical) = 0
        m_cScrollBar.SmallChange(efsVertical) = 0
        m_cScrollBar.LargeChange(efsVertical) = 0
    Else
        bShown = True
        m_cScrollBar.SmallChange(efsVertical) = 5
        m_cScrollBar.LargeChange(efsVertical) = 30
    End If

    If lL > ScaleWidth + -m_cScrollBar.Value(efsHorizontal) Then
        m_cScrollBar.Visible(efsHorizontal) = True
        m_cScrollBar.Max(efsHorizontal) = m_lIconHeight + lL + 30 - ScaleWidth
        m_cScrollBar.SmallChange(efsHorizontal) = 1
        m_cScrollBar.LargeChange(efsHorizontal) = 5
    Else
        m_cScrollBar.Visible(efsHorizontal) = False
        m_cScrollBar.Max(efsHorizontal) = 0
        m_cScrollBar.Value(efsHorizontal) = 0
        m_cScrollBar.SmallChange(efsHorizontal) = 0
        m_cScrollBar.LargeChange(efsHorizontal) = 0
    End If

End Sub

'//---------------------------------------------------------------------------------------
' Procedure : MoveFirst
' Type      : Sub
' DateTime  : 04/10/2004 15:36
' Author    : Gary Noble
' Purpose   : Moves To The First node
' Returns   :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  04/10/2004
'//---------------------------------------------------------------------------------------
Public Sub MoveFirst()
    On Error Resume Next
    Set mSelectedNode = rCurNode(m_NodesS(LBound(m_NodesS) + 1))
    m_SelItem = mSelectedNode.Key
    m_SelectedID = 1
    m_cScrollBar.Value(efsVertical) = 0
    On Error GoTo 0
End Sub

'//---------------------------------------------------------------------------------------
' Procedure : MoveLast
' Type      : Sub
' DateTime  : 04/10/2004 15:37
' Author    : Gary Noble
' Purpose   : Moves To The Last Node
' Returns   :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  04/10/2004
'//---------------------------------------------------------------------------------------
Public Sub MoveLast()
    On Error Resume Next
    Set mSelectedNode = rCurNode(m_NodesS(UBound(m_NodesS)))
    m_SelItem = mSelectedNode.Key
    m_SelectedID = 1
    m_cScrollBar.Value(efsVertical) = m_cScrollBar.Max(efsVertical)

    On Error GoTo 0
End Sub

'//---------------------------------------------------------------------------------------
' Procedure : Nodes
' Type      : Property
' DateTime  : 04/10/2004 15:24
' Author    : Gary Noble
' Purpose   : Node Collection
' Returns   : xNodes
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  04/10/2004
'//---------------------------------------------------------------------------------------
Public Property Get Nodes() As xNodes
    If m_Nodes Is Nothing Then Set m_Nodes = New xNodes

    m_Nodes.Init ObjPtr(Me), UserControl.HWND
    Set Nodes = m_Nodes
    
End Property


'//---------------------------------------------------------------------------------------
' Procedure : pvExpand
' Type      : Sub
' DateTime  : 04/10/2004 15:37
' Author    : Gary Noble
' Purpose   : Used For Expanding All Nodes
' Returns   :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  04/10/2004
'//---------------------------------------------------------------------------------------
Private Sub pvExpand(bln As Boolean)
If m_Nodes Is Nothing Then Exit Sub

If m_Nodes.Nodes.Count <= 0 Then Exit Sub

    Dim ss As xTreeNode
    
    If m_Nodes.Nodes.Count > 0 Then
        For Each ss In m_Nodes.Nodes
            Set ss = rCurNode(ObjPtr(ss))
                ss.Expanded = bln
            If ss.ChildCount > 0 Then
                '-- Expand the Children
                pvExpandX ss, bln
            End If
        Next
    End If

End Sub

'//---------------------------------------------------------------------------------------
' Procedure : pvExpandX
' Type      : Sub
' DateTime  : 04/10/2004 15:38
' Author    : Gary Noble
' Purpose   : used For Expanding Nested Nodes
' Returns   :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  04/10/2004
'//---------------------------------------------------------------------------------------
Private Sub pvExpandX(xNode As xTreeNode, bln As Boolean)

    Dim ss As xTreeNode
    Dim xcNode As xTreeNode
    
    For Each ss In xNode.Children.Nodes
        Set xcNode = rCurNode(ObjPtr(ss))
            If xcNode.ChildCount > 0 Then
            xcNode.Expanded = bln
            pvExpandX xcNode, bln
        End If
    Next


End Sub

'//---------------------------------------------------------------------------------------
' Procedure : pvGetNode
' Type      : Sub
' DateTime  : 04/10/2004 15:38
' Author    : Gary Noble
' Purpose   : Used For Searching For A Node
' Returns   :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  04/10/2004
'//---------------------------------------------------------------------------------------
Private Sub pvGetNode(Key As String)

    Dim ss As xTreeNode
    
    If m_Nodes Is Nothing Then Exit Sub
    
    If m_Nodes.Nodes.Count > 0 Then

        For Each ss In m_Nodes.Nodes

            If ss.Key = Key Then Set XSearchNode = ss: bFound = True: Exit Sub
            If bFound Then Exit Sub
            If ss.ChildCount > 0 Then
                pvGetNodeX ss, Key
            End If

        Next

    End If

End Sub


'//---------------------------------------------------------------------------------------
' Procedure : pvGetNodeByCaption
' Type      : Sub
' DateTime  : 04/10/2004 15:39
' Author    : Gary Noble
' Purpose   : used For Searching To get A node By Its Caption
' Returns   :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  04/10/2004
'//---------------------------------------------------------------------------------------
Private Sub pvGetNodeByCaption(Caption As String)

    Dim ss As xTreeNode
    If m_Nodes.Nodes.Count > 0 Then

        For Each ss In m_Nodes.Nodes
            If ss.Caption = Caption Then Set XSearchNode = ss: bFound = True: Exit Sub
            If bFound Then Exit Sub
            If ss.ChildCount > 0 Then
                pvGetNodeByCaptionX ss, Caption
            End If
        Next

    End If

End Sub


'//---------------------------------------------------------------------------------------
' Procedure : pvGetNodeByCaptionX
' Type      : Sub
' DateTime  : 04/10/2004 15:39
' Author    : Gary Noble
' Purpose   : Used For Nested Nodes Searching
' Returns   :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  04/10/2004
'//---------------------------------------------------------------------------------------
Private Sub pvGetNodeByCaptionX(xNode As xTreeNode, Caption As String)

    Dim ss As xTreeNode

    For Each ss In xNode.Children.Nodes
        If ss.Caption = Caption Then Set XSearchNode = ss: bFound = True: Exit Sub
        If bFound Then Exit Sub
        If ss.ChildCount > 0 Then
            pvGetNodeByCaptionX ss, Caption
        End If
    Next


End Sub

Private Sub pvGetNodeX(xNode As xTreeNode, Key As String)

    Dim ss As xTreeNode

    For Each ss In xNode.Children.Nodes
        If ss.Key = Key Then Set XSearchNode = ss: bFound = True: Exit Sub
        If bFound Then Exit Sub
        If ss.ChildCount > 0 And ss.Expanded Then
            pvGetNodeX ss, Key
        End If
    Next


End Sub


'//---------------------------------------------------------------------------------------
' Procedure : pvPutNodesToArray
' Type      : Sub
' DateTime  : 04/10/2004 15:40
' Author    : Gary Noble
' Purpose   : The Main Point
'             For Drawing To The Screen We Make An Array Of Pointers To The node Objects.
'             This Enables us To Draw Faster And Not Take Up So Much Resource.
' Returns   :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  04/10/2004
'//---------------------------------------------------------------------------------------
Private Sub pvPutNodesToArray()
    On Error Resume Next
    
    Dim ss As xTreeNode
    Erase m_NodesS
    ReDim m_NodesS(0) As Long
    Erase marr_Widths
    ReDim marr_Widths(0)
    
    '-- Scroll Mask
    lL = 0

    '-- Bail
    If m_Nodes.Nodes.Count <= 0 Then Exit Sub

    '-- Loop Through The Nodes And Attach The Pointer To The Array
    
    For Each ss In m_Nodes.Nodes
        ReDim Preserve m_NodesS(UBound(m_NodesS) + 1)
        
        If Me.ShowHorizontalScrollbar Then
            '-- Used For Seting The Horizonal Scrollbar
            ReDim Preserve marr_Widths(UBound(marr_Widths) + 1)
            marr_Widths(UBound(m_NodesS)) = TextWidth(ss.Caption & IIf(Len(ss.ItemData) > 0, " (" & ss.ItemData & "       )", ""))
            If marr_Widths(UBound(m_NodesS)) > lL Then lL = marr_Widths(UBound(m_NodesS))
        End If
        '-- Set The Pointer
        m_NodesS(UBound(m_NodesS)) = ObjPtr(ss)

        If ss.ChildCount > 0 And ss.Expanded Then
            '-- If The Node has Children The Do The Same
            pvPutNodesToArrayX ss
        End If

    Next

End Sub

'//---------------------------------------------------------------------------------------
' Procedure : pvPutNodesToArrayX
' Type      : Sub
' DateTime  : 04/10/2004 15:43
' Author    : Gary Noble
' Purpose   : Same As pvPutNodesToArray But Used For children
' Returns   :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  04/10/2004
'//---------------------------------------------------------------------------------------
Private Sub pvPutNodesToArrayX(xNode As xTreeNode)

    Dim ss As xTreeNode

    For Each ss In xNode.Children.Nodes
        ReDim Preserve m_NodesS(UBound(m_NodesS) + 1)
        m_NodesS(UBound(m_NodesS)) = ObjPtr(ss)
        
        If Me.ShowHorizontalScrollbar Then
            ReDim Preserve marr_Widths(UBound(marr_Widths) + 1)
            marr_Widths(UBound(m_NodesS)) = TextWidth(ss.Caption & IIf(Len(ss.ItemData) > 0, " (" & ss.ItemData & "       )", ""))
            If marr_Widths(UBound(m_NodesS)) > lL Then lL = marr_Widths(UBound(m_NodesS))
        End If
        If ss.ChildCount > 0 And ss.Expanded Then
            pvPutNodesToArrayX ss
        End If
    Next


End Sub

'//---------------------------------------------------------------------------------------
' Procedure : rCurNode
' Type      : Property
' DateTime  : 04/10/2004 15:43
' Author    : Gary Noble
' Purpose   : Return The Node Object From The Pointer
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



'//---------------------------------------------------------------------------------------
' Procedure : Redraw
' Type      : Function
' DateTime  : 04/10/2004 15:43
' Author    : Gary Noble
' Purpose   : Draws The Control
' Returns   : Variant
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  04/10/2004
'//---------------------------------------------------------------------------------------
Public Function Redraw()
    On Error Resume Next

    Dim ss As xTreeNode
    Dim i As Long
    Dim r As RECT
    Dim CurNode As xTreeNode
    Dim xx As Long
    Dim EX As RECT
    Dim fntOrig As StdFont
    Dim rcText As RECT
    Dim xY As Long
    Dim tHeightX As Long
    Dim tWidthX As Long
    Dim lOffset As Long
    Dim bRectsOnly As Boolean
    Dim Parent As xTreeNode
    Dim lForecolor As Long

    
    lForecolor = UserControl.ForeColor


    '-- Scrollbar ??
    m_cScrollBar.Enabled(efsHorizontal) = Me.Enabled
    m_cScrollBar.Enabled(efsVertical) = Me.Enabled

    '-- Draw The Background Picture if Any
    If Not Me.BackGroundPicture Is Nothing Then
        m_Tiler.Picture = Me.BackGroundPicture
        m_Tiler.TileArea hdc, 0, 0, ScaleWidth, ScaleHeight
    Else
        Cls
    End If

    
    '-- Stor The Original font
    Set fntOrig = UserControl.Font

    '-- Bail
    If UBound(m_NodesS) <= 0 Then GoTo CleanExit
    
    If mSelectedNode Is Nothing Then Set mSelectedNode = rCurNode(m_NodesS(1))
        m_SelItem = mSelectedNode.Key
        
    
    '-- Ids For Scrollbar And Drawing
    lOffset = -m_cScrollBar.Value(efsHorizontal)
    lFirstdrawn = m_cScrollBar.Value(efsVertical) + 1
    xY = 0
    
    '-- Loop Throug The Array And Draw The Nodes
    For i = m_cScrollBar.Value(efsVertical) + 1 To UBound(m_NodesS)

        Set CurNode = rCurNode(m_NodesS(i))
        Set Parent = rCurNode(CurNode.ParentPTR)

        UserControl.Font.Bold = CurNode.Bold

        '-- Calculate The Text Width
        If m_bIsNt Then
            DrawTextW UserControl.hdc, StrPtr(CurNode.Caption), -1, rcText, &H400&
        Else
            DrawTextA UserControl.hdc, CurNode.Caption, -1, rcText, &H400&
        End If

        tHeightX = rcText.bottom - rcText.top 'TextHeight("Q")
        If m_lIconHeight > tHeightX Then tHeightX = m_lIconHeight
        
            
        '-- Set The Rectangle
        If Not Parent Is Nothing Then
            CurNode.SetRect Parent.RectData.left + 15, xY, Parent.RectData.left + 15 + (rcText.right - rcText.left) + 15, (tHeightX + (xY))
        Else
            CurNode.SetRect xx + m_lIconWidth, xY, xx + m_lIconWidth + (rcText.right - rcText.left) + 15, (tHeightX + (xY))
        End If


        If Not bRectsOnly Then

            '-- Draw The Node Background
            If CurNode.Key = m_SelItem Then
                m_bDrawingSelectedNode = True

                LSet EX = CurNode.RectData
                bDrawingSelItem = True

                '-- Draw The Selected Node
                If Not Me.FullRowSelect Then
                    mGraphics.UtilDrawBackground UserControl.hdc, IIf(m_bLostFocus, Me.LostFocusSelectedBackColor, Me.SelectedBackColor), IIf(m_bLostFocus, Me.LostFocusSelectedBackColor, Me.SelectedBackColorTwo), lOffset + xx + 15 + EX.left - 1, EX.top, TextWidth(CurNode.Caption) + 3, tHeightX + 1
                    If Not m_bLostFocus Then mGraphics.UtilDrawBorderRectangle UserControl.hdc, Me.SelectedColorBorder, lOffset + xx + 15 + EX.left - 1, EX.top, TextWidth(CurNode.Caption) + 3, tHeightX + 1, False
                Else
                    mGraphics.UtilDrawBackground UserControl.hdc, IIf(m_bLostFocus, Me.LostFocusSelectedBackColor, Me.SelectedBackColor), IIf(m_bLostFocus, Me.LostFocusSelectedBackColor, Me.SelectedBackColorTwo), 1, EX.top, ScaleWidth - 2, tHeightX + 1
                    If Not m_bLostFocus Then mGraphics.UtilDrawBorderRectangle UserControl.hdc, Me.SelectedColorBorder, 1, EX.top, ScaleWidth - 2, tHeightX + 1, False
                End If

            End If

            '-- Draw The Open Close Glymph
            If CurNode.ChildCount > 0 And CurNode.Expanded Then
                LSet EX = CurNode.RectData
                EX.top = EX.top + ((((tHeightX \ 2) + ((EX.bottom - EX.top) \ 2))) \ 2) - 7
                EX.left = -1 + -m_lIconWidth + EX.left + lOffset
                mGraphics.DrawOpenCloseGlyph HWND, UserControl.hdc, EX, False

            ElseIf CurNode.ChildCount > 0 And Not CurNode.Expanded Then
                LSet EX = CurNode.RectData
                EX.left = -1 + -m_lIconWidth + EX.left + lOffset
                EX.top = EX.top + ((((tHeightX \ 2) + ((EX.bottom - EX.top) \ 2))) \ 2) - 7
                mGraphics.DrawOpenCloseGlyph HWND, UserControl.hdc, EX, True
            End If

            LSet EX = CurNode.RectData
            ImageListDrawIcon m_lptrVb6ImageList, hdc, m_hIml, CurNode.IconIndex, lOffset + -m_lIconWidth + EX.left + 13, CurNode.RectData.top + ((CurNode.RectData.bottom - CurNode.RectData.top) - (m_lIconHeight)) \ 2, False, , True
            
            EX.left = EX.left + lOffset
           
            '-- Draw The Text
            If CurNode.Key = m_SelItem Then
                mGraphics.UtilDrawText UserControl.hdc, CurNode.Caption, 15 + EX.left, CurNode.RectData.top + ((CurNode.RectData.bottom - CurNode.RectData.top) - (TextHeight("Q"))) \ 2, CurNode.RectData.right, CurNode.RectData.bottom, Me.Enabled, Me.SelectedColor, False
            Else
                mGraphics.UtilDrawText UserControl.hdc, CurNode.Caption, 15 + EX.left, CurNode.RectData.top + ((CurNode.RectData.bottom - CurNode.RectData.top) - (TextHeight("Q"))) \ 2, CurNode.RectData.right, CurNode.RectData.bottom, Me.Enabled, IIf(CurNode.ForeColor > 0, CurNode.ForeColor, lForecolor), False
            End If

            '-- Draw The ItemData
            If Len(CurNode.ItemData) > 0 Then
                LSet EX = CurNode.RectData
                EX.top = EX.top + ((((tHeightX \ 2) + ((EX.bottom - EX.top) \ 2))) \ 2) - 7
                UserControl.Font.Bold = CurNode.ItemDataBold
                mGraphics.UtilDrawText UserControl.hdc, "(" & CurNode.ItemData & ")", lOffset + EX.right + IIf(CurNode.ItemDataBold, 6, 4), CurNode.RectData.top + ((CurNode.RectData.bottom - CurNode.RectData.top) - (TextHeight("Q"))) \ 2, EX.right + TextWidth("(" & CurNode.ItemData & ")") + 32, CurNode.RectData.bottom, Me.Enabled, IIf(CurNode.ItemDataColor <> UserControl.ForeColor, CurNode.ItemDataColor, UserControl.ForeColor), False
                UserControl.Font.Bold = False
            End If

        End If
        Set UserControl.Font = fntOrig
 
        xY = xY + tHeightX + 2
        llastdrawn = i
        
        '-- Only Draw What We Can See
        If xY > ScaleHeight Then Exit For    'Then bRectsOnly = True

    Next


CleanExit:

On Error GoTo 0
End Function


'//---------------------------------------------------------------------------------------
' Procedure : RefreshData
' Type      : Sub
' DateTime  : 04/10/2004 15:48
' Author    : Gary Noble
' Purpose   : Refreshes The Data And Draws the control
' Returns   :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  04/10/2004
'//---------------------------------------------------------------------------------------
Public Sub RefreshData()
    pvPutNodesToArray
    MakeScrollVisible
    Redraw
End Sub

Public Property Get SelectedBackColor() As OLE_COLOR
    SelectedBackColor = m_SelectedBackColor
End Property

Public Property Let SelectedBackColor(ByVal New_SelectedBackColor As OLE_COLOR)
    m_SelectedBackColor = New_SelectedBackColor
    PropertyChanged "SelectedBackColor"
    Redraw
End Property

Public Property Get SelectedBackColorTwo() As OLE_COLOR
    SelectedBackColorTwo = m_SelectedBackColorTwo
End Property

Public Property Let SelectedBackColorTwo(ByVal New_SelectedBackColorTwo As OLE_COLOR)
    m_SelectedBackColorTwo = New_SelectedBackColorTwo
    PropertyChanged "SelectedBackColorTwo"
End Property

Public Property Get SelectedColor() As OLE_COLOR
    SelectedColor = m_SelectedColor
End Property

Public Property Let SelectedColor(ByVal New_SelectedColor As OLE_COLOR)
    m_SelectedColor = New_SelectedColor
    PropertyChanged "SelectedColor"
End Property

Public Property Get SelectedColorBorder() As OLE_COLOR
    SelectedColorBorder = m_SelectedColorBorder
End Property

Public Property Let SelectedColorBorder(ByVal New_SelectedColorBorder As OLE_COLOR)
    m_SelectedColorBorder = New_SelectedColorBorder
    PropertyChanged "SelectedColorBorder"
End Property

Public Function Selectednode() As xTreeNode
    Set Selectednode = mSelectedNode
End Function



'//---------------------------------------------------------------------------------------
' Procedure : SetCur
' Type      : Function
' DateTime  : 04/10/2004 15:49
' Author    : Gary Noble
' Purpose   : NOT USED. This Will Be Usesd For Hover Selection
' Returns   : Variant
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  04/10/2004
'//---------------------------------------------------------------------------------------
Private Function SetCur(x As Long, y As Long)
    On Error Resume Next
    Exit Function

    Dim i As Long
    Dim CurNode As xTreeNode
    Dim bGotcha As Boolean
    Dim r As RECT

    If m_Nodes.Nodes.Count = 0 Then bGotcha = False: GoTo CleanExit


    For i = m_cScrollBar.Value(efsVertical) + 1 To m_cScrollBar.Value(efsVertical) + Round(((ScaleHeight / (TextHeight("Q")))))
        Set CurNode = rCurNode(m_NodesS(i))
        LSet r = CurNode.RectData
        r.left = r.left + -m_cScrollBar.Value(efsHorizontal)
        r.right = r.right + -m_cScrollBar.Value(efsHorizontal)
        If PtInRect(r, x, y) Then
            bGotcha = True
            GoTo CleanExit
        End If
    Next
CleanExit:
    On Error Resume Next

    If Not CurNode Is Nothing Then
        If CurNode.ChildCount = 0 Then
            bGotcha = x >= r.left + 15
        End If
    End If
    UtilSetCursor bGotcha

    On Error GoTo 0
End Function


Private Sub UserControl_Click()
   ' SetCur CLng(msng_X), CLng(msng_y)

End Sub

Public Property Get HWND() As Long
    HWND = UserControl.HWND
End Property

Private Sub UserControl_DblClick()
    Dim s As xTreeNode

    If lButton = vbLeftButton Then
        Set s = HitTest(CLng(msng_X), CLng(msng_y))
      '  SetCur CLng(msng_X), CLng(msng_y)
        If Not s Is Nothing Then
            RaiseEvent BeforeExpand(s, Not s.Expanded)
            s.Expanded = Not s.Expanded
            Me.RefreshData
        End If
    End If
End Sub

Private Sub UserControl_EnterFocus()
    m_bLostFocus = False
    Me.Redraw
End Sub

Private Sub UserControl_Initialize()
    VerInitialise
    Set m_Tiler = New IAPP_BitMapTiler

    
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()


    Set UserControl.Font = Ambient.Font
    m_BackGradientOne = m_def_BackGradientOne
    m_BackGradientTwo = m_def_BackGradientTwo
    m_SelectedBackColor = m_def_SelectedBackColor
    m_FullRowSelect = m_def_FullRowSelect
    m_SelectedColor = m_def_SelectedColor
    m_SelectedBackColorTwo = m_def_SelectedBackColorTwo
    m_SelectedColorBorder = m_def_SelectedColorBorder
    Set m_BackGroundPicture = LoadPicture("")
    m_DisplayHorizontalScrollbar = m_def_DisplayHorizontalScrollbar
    m_LostFocusSelectedBackColor = m_def_LostFocusSelectedBackColor
'    m_NodeChange = m_def_NodeChange
'    m_Cleared = m_def_Cleared
    m_ShowHorizontalScrollbar = m_def_ShowHorizontalScrollbar
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Dim lLast As Long
    Dim lMaxMove As Long

    DoEvents

    lMaxMove = (llastdrawn - (lFirstdrawn + 1))

    If KeyCode = 38 Then
        lLast = m_SelectedID - 1
        If m_SelectedID - 1 <= LBound(m_NodesS) Then
            RaiseEvent BeforeNodeChange(mSelectedNode)
            Set mSelectedNode = rCurNode(m_NodesS(1))
            m_SelItem = mSelectedNode.Key
            m_SelectedID = 1
            RaiseEvent NodeSelected(mSelectedNode)
        Else
            RaiseEvent BeforeNodeChange(mSelectedNode)
            Set mSelectedNode = rCurNode(m_NodesS(m_SelectedID - 1))
            m_SelItem = mSelectedNode.Key
            m_SelectedID = m_SelectedID - 1
            RaiseEvent NodeSelected(mSelectedNode)
        End If

        If m_SelectedID < lFirstdrawn Then
            m_cScrollBar.Value(efsVertical) = m_cScrollBar.Value(efsVertical) - 1
        Else
            Redraw
        End If


    ElseIf KeyCode = 39 Then
        RaiseEvent BeforeExpand(mSelectedNode, Not mSelectedNode.Expanded)
        If Not mSelectedNode.Expanded Then mSelectedNode.Expanded = True
        Me.RefreshData
    ElseIf KeyCode = 37 Then
        RaiseEvent BeforeExpand(mSelectedNode, Not mSelectedNode.Expanded)
        If mSelectedNode.Expanded Then mSelectedNode.Expanded = False
        Me.RefreshData
    ElseIf KeyCode = 40 Then
        If m_SelectedID + 1 >= UBound(m_NodesS) Then
            RaiseEvent BeforeNodeChange(mSelectedNode)
            Set mSelectedNode = rCurNode(m_NodesS(UBound(m_NodesS)))
            m_SelItem = mSelectedNode.Key
            m_SelectedID = UBound(m_NodesS)
            RaiseEvent NodeSelected(mSelectedNode)
        Else
            RaiseEvent BeforeNodeChange(mSelectedNode)
            Set mSelectedNode = rCurNode(m_NodesS(m_SelectedID + 1))
            m_SelItem = mSelectedNode.Key
            m_SelectedID = m_SelectedID + 1
            RaiseEvent NodeSelected(mSelectedNode)
        End If

        If m_SelectedID + 1 > llastdrawn Then
            m_cScrollBar.Value(efsVertical) = m_cScrollBar.Value(efsVertical) + 1
        Else
            Redraw
        End If

        Redraw
    ElseIf KeyCode = 34 Then
        If m_SelectedID + lMaxMove > UBound(m_NodesS) Then
            RaiseEvent BeforeNodeChange(mSelectedNode)
            Set mSelectedNode = rCurNode(m_NodesS(UBound(m_NodesS)))
            m_SelItem = mSelectedNode.Key
            m_SelectedID = UBound(m_NodesS)
            RaiseEvent NodeSelected(mSelectedNode)
        Else
            RaiseEvent BeforeNodeChange(mSelectedNode)
            Set mSelectedNode = rCurNode(m_NodesS(m_SelectedID + lMaxMove))
            m_SelItem = mSelectedNode.Key
            m_SelectedID = m_SelectedID + lMaxMove
            RaiseEvent NodeSelected(mSelectedNode)
        End If

        If m_SelectedID > llastdrawn Then
            If m_SelectedID + lMaxMove > UBound(m_NodesS) Then
                m_cScrollBar.Value(efsVertical) = m_cScrollBar.Max(efsVertical)
            Else
                m_cScrollBar.Value(efsVertical) = m_cScrollBar.Value(efsVertical) + lMaxMove
            End If
        Else
            Redraw
        End If


    ElseIf KeyCode = 33 Then

        If m_SelectedID - lMaxMove <= LBound(m_NodesS) Then
            RaiseEvent BeforeNodeChange(mSelectedNode)
            Set mSelectedNode = rCurNode(m_NodesS(1))
            m_SelItem = mSelectedNode.Key
            m_SelectedID = 1
            RaiseEvent NodeSelected(mSelectedNode)
        Else
            RaiseEvent BeforeNodeChange(mSelectedNode)
            Set mSelectedNode = rCurNode(m_NodesS(m_SelectedID - lMaxMove))
            m_SelItem = mSelectedNode.Key
            m_SelectedID = m_SelectedID - lMaxMove
            RaiseEvent NodeSelected(mSelectedNode)
        End If
        If m_SelectedID < lFirstdrawn Then
            m_cScrollBar.Value(efsVertical) = m_cScrollBar.Value(efsVertical) - lMaxMove
        Else
            Redraw
        End If

    End If

    On Error GoTo 0

End Sub


Private Sub UserControl_LostFocus()
    m_bLostFocus = True
    Me.Redraw
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next



    Dim s As xTreeNode
    Dim bChev As Boolean
    msng_y = y
    msng_X = x
    lButton = Button
    If Button = vbLeftButton Then
        Set s = HitTest(CLng(x), CLng(y))
        If Not s Is Nothing Then
            'SetCur CLng(msng_X), CLng(msng_y)
            If mB_chev Then
                RaiseEvent BeforeExpand(s, Not s.Expanded)
                s.Expanded = Not s.Expanded: Me.RefreshData
            End If
            Redraw
        End If
    End If



End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    bEnsure = False

    

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    
    SetCur CLng(x), CLng(y)
    'Redraw

End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    If UserControl.Ambient.UserMode Then
        Set m_cScrollBar = New IAPP_ScrollBars
        m_cScrollBar.Create UserControl.HWND
        m_cScrollBar.Orientation = efsoVertical
        m_cScrollBar.Visible(efsVertical) = False
        m_cScrollBar.Visible(efsHorizontal) = False
    End If

    m_BackGradientOne = PropBag.ReadProperty("BackGradientOne", m_def_BackGradientOne)
    m_BackGradientTwo = PropBag.ReadProperty("BackGradientTwo", m_def_BackGradientTwo)
    m_SelectedBackColor = PropBag.ReadProperty("SelectedBackColor", m_def_SelectedBackColor)
    'Redraw
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    m_FullRowSelect = PropBag.ReadProperty("FullRowSelect", m_def_FullRowSelect)
    m_SelectedColor = PropBag.ReadProperty("SelectedColor", m_def_SelectedBackColor)
    m_SelectedBackColorTwo = PropBag.ReadProperty("SelectedBackColorTwo", m_def_SelectedBackColorTwo)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80FF&)
    m_SelectedColorBorder = PropBag.ReadProperty("SelectedColorBorder", m_def_SelectedColorBorder)
    Set m_BackGroundPicture = PropBag.ReadProperty("BackGroundPicture", Nothing)
    m_DisplayHorizontalScrollbar = PropBag.ReadProperty("DisplayHorizontalScrollbar", m_def_DisplayHorizontalScrollbar)

    m_LostFocusSelectedBackColor = PropBag.ReadProperty("LostFocusSelectedBackColor", m_def_LostFocusSelectedBackColor)
'    m_NodeChange = PropBag.ReadProperty("NodeChange", m_def_NodeChange)
'    m_Cleared = PropBag.ReadProperty("Cleared", m_def_Cleared)
    m_ShowHorizontalScrollbar = PropBag.ReadProperty("ShowHorizontalScrollbar", m_def_ShowHorizontalScrollbar)
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    Me.RefreshData
    On Error GoTo 0
End Sub

Private Sub UserControl_Terminate()

    Set m_Tiler = Nothing
    Set m_cScrollBar = Nothing
    
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("BackGradientOne", m_BackGradientOne, m_def_BackGradientOne)
    Call PropBag.WriteProperty("BackGradientTwo", m_BackGradientTwo, m_def_BackGradientTwo)
    Call PropBag.WriteProperty("SelectedBackColor", m_SelectedBackColor, m_def_SelectedBackColor)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("FullRowSelect", m_FullRowSelect, m_def_FullRowSelect)
    Call PropBag.WriteProperty("SelectedColor", m_SelectedColor, m_def_SelectedColor)
    Call PropBag.WriteProperty("SelectedBackColorTwo", m_SelectedBackColorTwo, m_def_SelectedBackColorTwo)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80FF&)
    Call PropBag.WriteProperty("SelectedColorBorder", m_SelectedColorBorder, m_def_SelectedColorBorder)
    Call PropBag.WriteProperty("BackGroundPicture", m_BackGroundPicture, Nothing)
    Call PropBag.WriteProperty("DisplayHorizontalScrollbar", m_DisplayHorizontalScrollbar, m_def_DisplayHorizontalScrollbar)
    Call PropBag.WriteProperty("LostFocusSelectedBackColor", m_LostFocusSelectedBackColor, m_def_LostFocusSelectedBackColor)
    Call PropBag.WriteProperty("ShowHorizontalScrollbar", m_ShowHorizontalScrollbar, m_def_ShowHorizontalScrollbar)
End Sub

Private Sub VScroll1_Change()
    Redraw
    UserControl.SetFocus

End Sub

Private Sub VScroll1_Scroll()
    Redraw
    UserControl.SetFocus
End Sub

'//---------------------------------------------------------------------------------------
' Procedure : xEnsureVisible
' Type      : Function
' DateTime  : 04/10/2004 15:51
' Author    : Gary Noble
' Purpose   : Used In Conjuntion With EnsureVisible for Nested Children Nodes
' Returns   : Variant
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  04/10/2004
'//---------------------------------------------------------------------------------------
Private Function xEnsureVisible(xNode As xTreeNode)

    Dim ss As xTreeNode

    If Not xNode.Expanded Then xNode.Expanded = True
    If xNode.Level = 0 Then Exit Function
    If Not rCurNode(xNode.ParentPTR) Is Nothing Then
        xEnsureVisible rCurNode(xNode.ParentPTR)
    End If
End Function

Public Property Let ImageList(ByRef vImageList As Variant)

   Dim o  As Object
   Dim rc As RECT

   On Error Resume Next
   m_hIml = 0
   m_lptrVb6ImageList = 0
   m_lIconWidth = 0
   m_lIconHeight = 0
   m_hIml = 0
   m_lptrVb6ImageList = 0

   Set m_vImageList = vImageList
   If VarType(vImageList) = vbLong Then
      '-- Assume a handle to an image list:
      m_hIml = vImageList
   ElseIf (VarType(vImageList) = vbObject) Then
      '-- Assume a VB image list:
      On Error Resume Next
      '-- Get the image list initialised..
      vImageList.ListImages(1).Draw 0, 0, 0, 1
      m_hIml = vImageList.hImageList
      If Err.Number = 0 Then
         '-- Check for VB6 image list:
         If TypeName(vImageList) = "ImageList" Then
            If vImageList.ListImages.Count <> ImageList_GetImageCount(m_hIml) Then
               Set o = vImageList
               m_lptrVb6ImageList = ObjPtr(o)
            End If
         End If
      Else
         'debug.print "Failed to Get Image list Handle", "EDCTRL.ImageList"
      End If
      On Error GoTo 0
   End If
   If m_hIml <> 0 Then
      If m_lptrVb6ImageList <> 0 Then
         m_lIconWidth = vImageList.ImageWidth
         m_lIconHeight = vImageList.ImageHeight
         If (UserControl.Extender.Align = vbAlignLeft) Or (UserControl.Extender.Align = vbAlignRight) Then
            UserControl_Resize
         End If
      Else   'NOT (m_lptrVb6ImageList...
         ImageList_GetImageRect m_hIml, 0, rc
         m_lIconWidth = rc.right - rc.left
         m_lIconHeight = rc.bottom - rc.top
         If (UserControl.Extender.Align = vbAlignLeft) Or (UserControl.Extender.Align = vbAlignRight) Then
            UserControl_Resize
         End If
      End If
   End If
   On Error GoTo 0
   Redraw

End Property

Private Sub ImageListDrawIcon(ByVal ptrVb6ImageList As Long, _
                              ByVal lngHdc As Long, _
                              ByVal hIml As Long, _
                              ByVal iIconIndex As Long, _
                              ByVal lX As Long, _
                              ByVal lY As Long, _
                              Optional ByVal bSelected As Boolean = False, _
                              Optional ByVal bBlend25 As Boolean = False, _
                              Optional ByVal IsHeaderIcon As Boolean = False)


   Dim o          As Object
   Dim lFlags     As Long
   Dim lR         As Long
   Dim icoInfo    As ICONINFO
   Dim newICOinfo As ICONINFO
   Dim icoBMPinfo As BITMAP

   If Not Me.Enabled Then
      ImageListDrawIconDisabled ptrVb6ImageList, lngHdc, hIml, iIconIndex, lX, lY, m_lIconHeight, True
      Exit Sub
   End If
   lFlags = ILD_TRANSPARENT
   If bSelected Then
      lFlags = lFlags Or ILD_SELECTED
   End If
   If bBlend25 Then
      lFlags = lFlags Or ILD_BLEND25
   End If
   If ptrVb6ImageList <> 0 Then
      On Error Resume Next
      Set o = ObjectFromPtr(ptrVb6ImageList)
      If Not (o Is Nothing) Then
         If ((lFlags And ILD_SELECTED) = ILD_SELECTED) Then
            lFlags = 2   '-- best we can do in VB6
         End If
         GetIconInfo o.ListImages(iIconIndex + 1).ExtractIcon(), icoInfo
         If playImage Then
            DestroyIcon playImage
         End If
         '-- start a new icon structure
         CopyMemory newICOinfo, icoInfo, Len(icoInfo)
         '-- get the icon dimensions from the bitmap portion of the icon
         GetGDIObject icoInfo.hbmColor, Len(icoBMPinfo), icoBMPinfo
         sourceWidth = m_lIconWidth
         sourceHeight = m_lIconHeight
         playImage = CreateIconIndirect(newICOinfo)
         
         With o
            .ListImages(iIconIndex + 1).Draw hdc, lX * Screen.TwipsPerPixelX, lY * Screen.TwipsPerPixelY, lFlags
            '--  DrawIconEx lngHdc, lX, lY, playImage, sourceWidth, sourceHeight, 0, 0, &H3 Or ILD_BLEND25
            DeleteObject newICOinfo.hbmMask
            DeleteObject newICOinfo.hbmColor
         End With   'o
      End If
      On Error GoTo 0
   Else
      lR = ImageList_Draw(hIml, iIconIndex, lngHdc, lX, lY, lFlags)
      If lR = 0 Then
         'debug.print "Failed to draw Image: " & iIconIndex & " onto hDC " & hdc, "ImageListDrawIcon"
      End If
   End If

End Sub

Private Sub ImageListDrawIconDisabled(ByVal ptrVb6ImageList As Long, _
                                      ByVal lngHdc As Long, _
                                      ByVal hIml As Long, _
                                      ByVal iIconIndex As Long, _
                                      ByVal lX As Long, _
                                      ByVal lY As Long, _
                                      ByVal lSize As Long, _
                                      Optional ByVal asShadow As Boolean)

   Dim o     As Object
   Dim hBr   As Long
   Dim hIcon As Long

   'Dim lR    As Long
   hIcon = 0
   If ptrVb6ImageList <> 0 Then
      On Error Resume Next
      Set o = ObjectFromPtr(ptrVb6ImageList)
      If Not (o Is Nothing) Then
         hIcon = o.ListImages(iIconIndex + 1).ExtractIcon()
      End If
      On Error GoTo 0
   Else
      hIcon = ImageList_GetIcon(hIml, iIconIndex, 0)
   End If
   If hIcon <> 0 Then
      If asShadow Then
         hBr = GetSysColorBrush(vb3DShadow And &H1F)
         If lngHdc = hdc Then
            DrawState lngHdc, hBr, 0, hIcon, 0, lX, lY, lSize, lSize, DST_ICON Or DSS_MONO
         Else
            DrawState lngHdc, hBr, 0, hIcon, 0, lX, lY + 4, 16, 16, DST_ICON Or DSS_MONO
         End If
         DeleteObject hBr
      Else
         DrawState lngHdc, 0, 0, hIcon, 0, lX, lY, lSize, lSize, DST_ICON Or DSS_DISABLED
      End If
      DestroyIcon hIcon
   End If

End Sub



Public Property Get LostFocusSelectedBackColor() As OLE_COLOR
    LostFocusSelectedBackColor = m_LostFocusSelectedBackColor
End Property

Public Property Let LostFocusSelectedBackColor(ByVal New_LostFocusSelectedBackColor As OLE_COLOR)
    m_LostFocusSelectedBackColor = New_LostFocusSelectedBackColor
    PropertyChanged "LostFocusSelectedBackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,False
Public Property Get ShowHorizontalScrollbar() As Variant
    ShowHorizontalScrollbar = m_ShowHorizontalScrollbar
End Property

Public Property Let ShowHorizontalScrollbar(ByVal New_ShowHorizontalScrollbar As Variant)
    m_ShowHorizontalScrollbar = New_ShowHorizontalScrollbar
    PropertyChanged "ShowHorizontalScrollbar"
    Me.RefreshData
    
End Property

