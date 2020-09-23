VERSION 5.00
Begin VB.Form frm2DArray 
   AutoRedraw      =   -1  'True
   Caption         =   "Dynamic Multidimensional Array"
   ClientHeight    =   7695
   ClientLeft      =   12750
   ClientTop       =   2745
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   ScaleHeight     =   7695
   ScaleWidth      =   8415
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command5 
      Caption         =   "Delete Selected node"
      Height          =   495
      Left            =   5400
      TabIndex        =   5
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ExpandAll"
      Height          =   495
      Left            =   5400
      TabIndex        =   4
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Collapse"
      CausesValidation=   0   'False
      Height          =   495
      Left            =   5400
      TabIndex        =   3
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Make Font Bigger"
      Height          =   735
      Left            =   5400
      TabIndex        =   2
      Top             =   960
      Width           =   1455
   End
   Begin xTreeX.xTree xTree1 
      Align           =   3  'Align Left
      Height          =   7695
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   13573
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   5400
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "frm2DArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim k As Long

Private Sub Command1_Click()


    Dim s As xTreeNode
    Dim ss As xTreeNode
    Dim sss As xTreeNode
    Dim ssss As xTreeNode
    Dim i As Long
    Dim ii As Long
    Dim iii As Long
    Dim iiii As Long
    Dim id As Long
    Dim std As New StdFont

    std.Name = "Times New Roman"
    std.Size = 18
    std.Italic = True

Screen.MousePointer = vbHourglass
DoEvents


    For i = 1 To 50
        k = k + 1
        Set s = xTree1.Nodes.AddNode("Cap: " & k, "Sq" & k)
        If i = 2 Then
            s.Expanded = True
            s.Bold = True
            s.ForeColor = vbRed

        End If

        For ii = 0 To 5
            k = k + 1

            Set ss = xTree1.Nodes.AddNode("Cap: " & k, "SA" & k, s)
            If ii = 2 Then
                ss.Expanded = True
                ss.Bold = True
                ss.ForeColor = vbBlue

            End If

            For iii = 0 To 2
                k = k + 1
                If iii = 2 Then ss.Expanded = True
                Set sss = xTree1.Nodes.AddNode("Cap: " & k, "SP" & k, ss)
                For iiii = 0 To 2
                    k = k + 1
                    If iii = 2 Then ss.Expanded = True
                    Set ssss = xTree1.Nodes.AddNode("Cap: " & k, "SL" & k, sss)
                    If iiii = 2 Then
                        ssss.Expanded = True
                        ssss.Bold = True
                        ssss.ForeColor = vbGreen

                    End If

                Next
                sss.ItemData = s.Children.Nodes.Count
                sss.ItemDataBold = False
                sss.ItemDataColor = vbBlue
                

            Next
            ss.ItemData = ss.Children.Nodes.Count
            ss.ItemDataBold = True
            'ss.ItemDataColor = vbRed
            
            
        Next
            s.ItemData = s.Children.Nodes.Count
            s.ItemDataBold = True
            'ss.ItemDataColor = vbRed

    Next
    Debug.Print "Nodes Added: " & k
    On Error Resume Next
    Cls
    Set s = xTree1.Nodes(2).Children.AddNode("SSSSSSS", "SS")
    xTree1.RefreshData

    xTree1.Redraw

Screen.MousePointer = vbDefault
DoEvents


End Sub

Private Sub Command2_Click()
    Me.xTree1.Font.Size = Me.xTree1.Font.Size + 2
    Me.xTree1.RefreshData


    Me.xTree1.Redraw




End Sub

Private Sub Command3_Click()
    Dim s As xTreeNode
    On Error Resume Next

    Set s = xTree1.GetNode("Sq1")
    '  s.Caption = "This Has Changed"
    ' s.Expanded = Not s.Expanded

    'xTree1.RefreshData
    Me.xTree1.ExpandAllNodes False



    xTree1.Redraw

End Sub

Private Sub Command4_Click()
    Me.xTree1.ExpandAllNodes True
    xTree1.Redraw
End Sub

Private Sub Command5_Click()
    Me.xTree1.DeleteNode Me.xTree1.Selectednode







End Sub
