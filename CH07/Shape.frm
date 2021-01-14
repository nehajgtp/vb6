VERSION 5.00
Begin VB.Form frmShape 
   Caption         =   "The Shape Program"
   ClientHeight    =   5100
   ClientLeft      =   1200
   ClientTop       =   1635
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   ScaleHeight     =   5100
   ScaleWidth      =   7245
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdRectangle 
      Caption         =   "&Rectangle"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   6
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton cmdSquare 
      Caption         =   "&Square"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   5
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdOval 
      Caption         =   "&Oval"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   4
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmdCircle 
      Caption         =   "&Circle"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   3
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdRndSqr 
      Caption         =   "Rounded S&quare"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   3960
      Width           =   2295
   End
   Begin VB.CommandButton cmdRndRect 
      Caption         =   "Rounded Rectan&gle"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   1
      Top             =   3960
      Width           =   2295
   End
   Begin VB.HScrollBar hsbWidth 
      Height          =   255
      Left            =   1560
      Max             =   10
      Min             =   1
      TabIndex        =   0
      Top             =   3000
      Value           =   1
      Width           =   4935
   End
   Begin VB.Shape shpAllShapes 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   2640
      Shape           =   3  'Circle
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblInfo 
      Caption         =   "Change Width:"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   8
      Top             =   2640
      Width           =   1695
   End
End
Attribute VB_Name = "frmShape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' All variables MUST be declared.
Option Explicit

Private Sub cmdCircle_Click()
    ' The user clicked the Circle button,
    ' so set the Shape property of shpAllShapes to
    ' cirlce (3).
    shpAllShapes = 3
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdOval_Click()
    ' The user clicked the Oval button,
    ' so set the Shape property of shpAllShapes to
    ' oval (2).
    shpAllShapes.Shape = 2
End Sub

Private Sub cmdRectangle_Click()
    ' The user clicked the Rectangle button,
    ' so set the Shape property of shpAllShapes to
    ' rectangle (0).
    shpAllShapes.Shape = 0
End Sub

Private Sub cmdRndRect_Click()
    ' The user clicked the Rounded Rectangle button,
    ' so set the Shape property of shpAllShapes to
    ' rounded square (5).
    shpAllShapes.Shape = 5
End Sub

Private Sub cmdRndSqr_Click()
    ' The user clicked the Square button,
    ' so set the Shape property of shpAllShapes to
    ' rounded sqaure (5).
    shpAllShapes.Shape = 5
End Sub

Private Sub cmdSquare_Click()
    ' The user clicked the Square button,
    ' so set the Shape property of shpAllShapes to
    ' square (1).
    shpAllShapes.Shape = 1
End Sub

Private Sub hsbWidth_Change()
    ' The user changed the scroll bar position,
    ' so set the BorderWidth propoerty of
    ' shpAllShapes to the new Value of the scroll
    ' bar.
    shpAllShapes.BorderWidth = hsbWidth.Value
    
End Sub
