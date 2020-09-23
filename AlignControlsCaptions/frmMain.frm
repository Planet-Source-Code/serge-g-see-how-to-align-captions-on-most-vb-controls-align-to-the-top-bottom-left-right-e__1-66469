VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   4995
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   1020
      Left            =   3180
      TabIndex        =   19
      Top             =   2490
      Width           =   1695
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   825
      Left            =   3180
      TabIndex        =   17
      Top             =   315
      Width           =   1725
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3255
      TabIndex        =   16
      Top             =   3855
      Width           =   1455
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Command15"
      Height          =   495
      Left            =   360
      TabIndex        =   15
      Top             =   3720
      Width           =   2415
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Command14"
      Height          =   495
      Left            =   1560
      TabIndex        =   14
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Command13"
      Height          =   495
      Left            =   360
      TabIndex        =   13
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Command12"
      Height          =   495
      Left            =   1560
      TabIndex        =   12
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Command11"
      Height          =   495
      Left            =   360
      TabIndex        =   11
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Command10"
      Height          =   495
      Left            =   1560
      TabIndex        =   9
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Command9"
      Height          =   495
      Left            =   1560
      TabIndex        =   8
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Command8"
      Height          =   495
      Left            =   1560
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   495
      Left            =   1560
      TabIndex        =   6
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   495
      Left            =   1560
      TabIndex        =   5
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame"
      Height          =   4335
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   2895
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   645
      Left            =   3180
      TabIndex        =   18
      Top             =   1485
      Width           =   1710
   End
   Begin VB.Shape Shape2 
      Height          =   930
      Left            =   3135
      Top             =   270
      Width           =   1800
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   3120
      Top             =   1455
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Thanks to AllApi.net
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const BS_LEFT As Long = &H100
Private Const BS_RIGHT As Long = &H200
Private Const BS_CENTER As Long = &H300
Private Const BS_TOP As Long = &H400
Private Const BS_BOTTOM As Long = &H800
Private Const BS_VCENTER As Long = &HC00

Private Const BS_ALLSTYLES = BS_LEFT Or BS_RIGHT Or BS_CENTER Or BS_TOP Or BS_BOTTOM Or BS_VCENTER
Private Const GWL_STYLE& = (-16)

Public Enum bsHorizontalAlignments
    bsLeft = BS_LEFT
    bsright = BS_RIGHT
    bsCenter = BS_CENTER
End Enum

Public Enum bsVerticalAlignments
    bsTop = BS_TOP
    bsBottom = BS_BOTTOM
    bsVcenter = BS_VCENTER
End Enum

Public Sub AlignButtonText(cmd As Control, _
Optional ByVal HStyle As bsHorizontalAlignments = _
bsCenter, Optional ByVal VStyle As _
bsVerticalAlignments = bsVcenter)

    Dim oldStyle As Long
    ' retrieve the current style of the control
    oldStyle = GetWindowLong(cmd.hWnd, GWL_STYLE)
    ' change the style
    oldStyle = oldStyle And (Not BS_ALLSTYLES)
    ' set the style of the control to the new style
    Call SetWindowLong(cmd.hWnd, GWL_STYLE, _
    oldStyle Or HStyle Or VStyle)
    cmd.Refresh
End Sub

Private Sub Command19_Click()

    End

End Sub

Private Sub Command19_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call AlignButtonText(Command19, bsCenter, bsBottom)

End Sub

Private Sub Command19_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call AlignButtonText(Command19, bsCenter, bsTop)

End Sub

Private Sub Form_Load()
    Call AlignButtonText(Command1, bsCenter, bsTop)
    Call AlignButtonText(Command2, bsCenter, bsBottom)
    Call AlignButtonText(Command3, bsCenter, bsVcenter)
    Call AlignButtonText(Command4, bsLeft, bsTop)
    Call AlignButtonText(Command5, bsLeft, bsBottom)
    Call AlignButtonText(Command6, bsLeft, bsVcenter)
    Call AlignButtonText(Command7, bsright, bsTop)
    Call AlignButtonText(Command8, bsright, bsBottom)
    Call AlignButtonText(Command9, bsright, bsVcenter)
    Call AlignButtonText(Command10, , bsTop)
    Call AlignButtonText(Command11, , bsBottom)
    Call AlignButtonText(Command12, , bsVcenter)
    Call AlignButtonText(Command13, bsCenter)
    Call AlignButtonText(Command14, bsLeft)
    Call AlignButtonText(Command15, bsright)
    Call AlignButtonText(Frame2, bsCenter, bsVcenter)
    Call AlignButtonText(Command19, bsCenter, bsTop)
    
    Call AlignButtonText(Frame1, bsCenter, bsTop)
    Call AlignButtonText(Check1, bsLeft, bsBottom)
    Call AlignButtonText(Option1, bsright, bsTop)
    
    
End Sub


