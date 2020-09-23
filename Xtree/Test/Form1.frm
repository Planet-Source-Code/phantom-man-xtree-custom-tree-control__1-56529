VERSION 5.00
Object = "*\A..\xtree.vbp"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "XTree Version 1.0 Demo"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8250
   LinkTopic       =   "Form1"
   ScaleHeight     =   6060
   ScaleWidth      =   8250
   StartUpPosition =   3  'Windows Default
   Begin xTreeX.xTree xTree1 
      Align           =   3  'Align Left
      Height          =   6060
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   3015
      _extentx        =   5318
      _extenty        =   10689
      font            =   "Form1.frx":0000
      selectedcolor   =   -2147483635
      forecolor       =   8388608
   End
   Begin VB.Frame Frame1 
      Caption         =   " Options "
      Height          =   5295
      Left            =   3240
      TabIndex        =   2
      Top             =   720
      Width           =   4935
      Begin VB.CommandButton Command16 
         Caption         =   "Background Picture"
         Height          =   495
         Left            =   240
         TabIndex        =   20
         Top             =   3360
         Width           =   1455
      End
      Begin VB.PictureBox Picture2 
         Height          =   975
         Left            =   3480
         Picture         =   "Form1.frx":002C
         ScaleHeight     =   915
         ScaleWidth      =   1155
         TabIndex        =   19
         Top             =   3840
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Show Horiz Scrollbar"
         Height          =   495
         Left            =   3360
         TabIndex        =   18
         Top             =   2760
         Width           =   1455
      End
      Begin MSComctlLib.ImageList LargeImages 
         Left            =   1920
         Top             =   3600
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   13
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":9A6BE
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":9CE70
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":9F622
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":9FA74
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":9FEC6
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":A0318
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":A076A
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":A0A84
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":A0D9E
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":A0EF8
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":A1052
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":A2554
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":A325E
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList SmallImages 
         Left            =   1080
         Top             =   3600
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   13
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":A3578
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":A5D2A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":A84DC
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":A892E
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":A8D80
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":A91D2
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":A9624
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":A993E
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":A9C58
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":A9DB2
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":A9F0C
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":AB40E
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":AC118
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton Command14 
         Caption         =   "No Icons"
         Height          =   495
         Left            =   3360
         TabIndex        =   16
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Small Icons"
         Height          =   495
         Left            =   3360
         TabIndex        =   15
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Delete Selected node"
         Height          =   495
         Left            =   240
         TabIndex        =   14
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         Caption         =   "ExpandAll"
         Height          =   495
         Left            =   1800
         TabIndex        =   13
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Collapse"
         CausesValidation=   0   'False
         Height          =   495
         Left            =   1800
         TabIndex        =   12
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Make Font Bigger"
         Height          =   495
         Left            =   3360
         TabIndex        =   11
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Load 1000 Nodes"
         Height          =   495
         Left            =   1800
         TabIndex        =   10
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Enable / Disable"
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Clear"
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   2760
         Width           =   1455
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Full Row Select"
         Height          =   495
         Left            =   1800
         TabIndex        =   7
         Top             =   2760
         Width           =   1455
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Last Node"
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton Command10 
         Caption         =   "First Node"
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Ensure Selected"
         Height          =   495
         Left            =   1800
         TabIndex        =   4
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Large Icons"
         Height          =   495
         Left            =   3360
         TabIndex        =   3
         Top             =   960
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   3240
      ScaleHeight     =   615
      ScaleWidth      =   5055
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "XTree Version 1.0 Demo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   5055
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim k As Long

Private Sub Command1_Click()
    On Error GoTo ErrAdd

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
 
    Me.xTree1.ImageList = Me.SmallImages
    
    
    '-- Add The Nodes
    For i = 1 To 10

        k = k + 1
        Set s = xTree1.Nodes.AddNode("Cap: " & k, "Sq" & k, 1)
        If i = 2 Then
            s.Expanded = True
            s.Bold = True
            s.ForeColor = vbRed

        End If

        For ii = 0 To 5
            k = k + 1

            Set ss = xTree1.Nodes.AddNode("Cap: " & k, "SA" & k, 2, s)
            If ii = 2 Then
                ss.Expanded = True
                ss.Bold = True
                ss.ForeColor = vbBlue

            End If

            For iii = 0 To 2
                k = k + 1
                If iii = 2 Then ss.Expanded = True
                Set sss = xTree1.Nodes.AddNode("Cap: " & k, "SP" & k, 3, ss)
                For iiii = 0 To 2
                    k = k + 1
                    If iii = 2 Then ss.Expanded = True
                    Set ssss = xTree1.Nodes.AddNode("Cap: " & k, "SL" & k, 5, sss)
                    If iiii = 2 Then
                        ssss.Expanded = True
                        ssss.Bold = True
                        ssss.ForeColor = vbGreen

                    End If

                Next
                sss.ItemData = sss.Children.Nodes.Count
                sss.ItemDataBold = False
                sss.ItemDataColor = vbBlue


            Next
            ss.ItemData = ss.Children.Nodes.Count
            ss.ItemDataBold = True
            ss.ItemDataColor = vbRed


        Next
        If i = 3 Then
            s.Caption = "This is A Long String Test"
            s.ItemData = "Child Count " & s.Children.Nodes.Count
            s.ItemDataBold = True
            s.ItemDataColor = vbRed
        End If
    Next
    Debug.Print "Nodes Added: " & k
    
    On Error Resume Next
    xTree1.RefreshData
    Set s = xTree1.GetNode("Sq1")
    xTree1.EnsureVisible s

CleanExit:

    Screen.MousePointer = vbDefault
    DoEvents
Exit Sub
ErrAdd:
        MsgBox Err.Description
        Resume CleanExit

End Sub

Private Sub Command10_Click()
    Me.xTree1.MoveFirst



End Sub

Private Sub Command11_Click()
    Dim x As xTreeNode

    Set x = Me.xTree1.GetNodebyCaption("Cap: 136")
    xTree1.EnsureVisible x


End Sub

Private Sub Command12_Click()

    Me.xTree1.ImageList = Me.LargeImages

End Sub

Private Sub Command13_Click()
    Me.xTree1.ImageList = Me.SmallImages

End Sub

Private Sub Command14_Click()
    Me.xTree1.ImageList = ""
    
    
End Sub

Private Sub Command15_Click()
Me.xTree1.ShowHorizontalScrollbar = Not Me.xTree1.ShowHorizontalScrollbar


End Sub

Private Sub Command16_Click()
    If xTree1.BackGroundPicture Is Nothing Then
        Set Me.xTree1.BackGroundPicture = Me.Picture2.Picture
    Else
        Set Me.xTree1.BackGroundPicture = Nothing
    End If
    
End Sub

Private Sub Command2_Click()
    Me.xTree1.Font.Size = Me.xTree1.Font.Size + 2
    Me.xTree1.RefreshData
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

Private Sub Command6_Click()
    Me.xTree1.Enabled = Not Me.xTree1.Enabled


End Sub

Private Sub Command7_Click()
    Me.xTree1.Clear

End Sub

Private Sub Command8_Click()
    Me.xTree1.FullRowSelect = Not Me.xTree1.FullRowSelect


End Sub

Private Sub Command9_Click()
    Me.xTree1.MoveLast


End Sub

Private Sub xTree1_BeforeExpand(xNode As xTreeX.xTreeNode, bExpanding As Boolean)
    Debug.Print "Expanding: " & xNode.Caption & "      -  " & bExpanding
    
End Sub

Private Sub xTree1_BeforeNodeChange(xNode As xTreeX.xTreeNode)
    Debug.Print "Before Node Change: " & xNode.Caption
End Sub

Private Sub xTree1_Cleared()
    Debug.Print "Nodes Cleared"
End Sub

Private Sub xTree1_NodeChange(xNode As xTreeX.xTreeNode)
    Debug.Print "Node Changed: " & xNode.Caption
    
End Sub

Private Sub xTree1_NodeSelected(xNode As xTreeX.xTreeNode)
    Debug.Print "Selected Node: " & xNode.Caption
    
End Sub
