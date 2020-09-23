VERSION 5.00
Begin VB.Form Form_main 
   Caption         =   "Pure VB diagram monitors"
   ClientHeight    =   5175
   ClientLeft      =   1800
   ClientTop       =   1965
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   6090
   Begin VB.Frame Frame2 
      Caption         =   "Point"
      Height          =   2235
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   5835
      Begin VB.PictureBox Picture_point2 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         ClipControls    =   0   'False
         Height          =   1695
         Left            =   2940
         ScaleHeight     =   113
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   181
         TabIndex        =   6
         Top             =   300
         Width           =   2715
      End
      Begin VB.PictureBox Picture_point 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         ClipControls    =   0   'False
         Height          =   1695
         Left            =   180
         ScaleHeight     =   113
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   181
         TabIndex        =   4
         Top             =   300
         Width           =   2715
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   3720
      Top             =   4680
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   315
      Left            =   4380
      TabIndex        =   2
      Top             =   4740
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Line"
      Height          =   2235
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5835
      Begin VB.PictureBox Picture_line2 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         ClipControls    =   0   'False
         Height          =   1695
         Left            =   2940
         ScaleHeight     =   113
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   181
         TabIndex        =   5
         Top             =   300
         Width           =   2715
      End
      Begin VB.PictureBox Picture_line 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         ClipControls    =   0   'False
         Height          =   1695
         Left            =   180
         ScaleHeight     =   113
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   181
         TabIndex        =   1
         Top             =   300
         Width           =   2715
      End
   End
End
Attribute VB_Name = "Form_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''
'Simple diagrams
'Just add new reference to cls_diagram!
'By Sami Riihilahti
'Free to use at any case!
''''''''''''''''''''''''''''''''''''

Public linediagram As New Cls_diagram
Public linediagram2 As New Cls_diagram
Public pointdiagram As New Cls_diagram
Public pointdiagram2 As New Cls_diagram
Public tancounter As Single 'We are using this just to create tan diagrams

Private Sub Command1_Click()

    'Create all 4 diagrams. See what options we have when calling .InitDiagram()
    
    '.InitDiagram picturebox, linecolor, showgrid, gridcolor, movinggrid
    'picturebox = the picturebox in which to add the diagram
    'linecolor = line color
    'showgrid = grid ON/OFF
    'gridcolor = [optional] grid color (default=dark green)
    'movinggrid = [optional] boolean true/false of moving grid (default=false)
    
    With linediagram
        .InitDiagram Picture_line, RGB(0, 255, 0), True
        .Max = 10
        .HorzSplits = 9
        .VertSplits = 9
        .DiagramType = TYPE_LINE
        .RePaint
    End With
    With pointdiagram
        .InitDiagram Picture_point, RGB(0, 255, 0), True
        .Max = 20
        .HorzSplits = 9
        .VertSplits = 9
        .DiagramType = TYPE_POINT
        .RePaint
    End With
    With linediagram2
        .InitDiagram Picture_line2, RGB(255, 255, 0), True, , True
        .Max = 100
        .HorzSplits = 9
        .VertSplits = 9
        .DiagramType = TYPE_LINE
        .RePaint
    End With
    With pointdiagram2
        .InitDiagram Picture_point2, RGB(0, 255, 255), True, RGB(100, 0, 0), True
        .Max = 10
        .HorzSplits = 9
        .VertSplits = 9
        .DiagramType = TYPE_POINT
        .RePaint
    End With
    
End Sub

Private Sub Picture_line_Paint()
    linediagram.RePaint
End Sub

Private Sub Picture_line2_Click()
    linediagram2.RePaint
End Sub

Private Sub Picture_point_Paint()
    pointdiagram.RePaint
End Sub

Private Sub Picture_point2_Click()
    pointdiagram.RePaint
End Sub

Private Sub Timer1_Timer()

    'Just randomize a new value in this sample
    Dim value As Single
    tancounter = tancounter + 0.1
    value = Tan(tancounter) + 2
    
    linediagram.AddValue value
    linediagram2.AddValue value
    pointdiagram.AddValue value
    pointdiagram2.AddValue value
    linediagram.RePaint
    linediagram2.RePaint
    pointdiagram.RePaint
    pointdiagram2.RePaint
End Sub
