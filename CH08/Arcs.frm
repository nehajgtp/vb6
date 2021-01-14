VERSION 5.00
Begin VB.Form frmArcs 
   Caption         =   "The Arcs Program"
   ClientHeight    =   5145
   ClientLeft      =   1200
   ClientTop       =   1635
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   ScaleHeight     =   5145
   ScaleWidth      =   7320
   Begin VB.VScrollBar vsbRadius 
      Height          =   3135
      Left            =   360
      Max             =   100
      Min             =   1
      TabIndex        =   3
      Top             =   960
      Value           =   10
      Width           =   255
   End
   Begin VB.HScrollBar hsbTo 
      Height          =   255
      Left            =   3720
      Max             =   360
      TabIndex        =   2
      Top             =   3840
      Width           =   2655
   End
   Begin VB.HScrollBar hsbFrom 
      Height          =   255
      Left            =   720
      Max             =   360
      TabIndex        =   1
      Top             =   3840
      Width           =   2655
   End
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
      Left            =   3000
      TabIndex        =   0
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label lblRadius 
      Caption         =   "Radius:"
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
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label lblTo 
      Caption         =   "To:"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   5
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label lblFrom 
      Caption         =   "From:"
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
      Left            =   720
      TabIndex        =   4
      Top             =   3480
      Width           =   1215
   End
End
Attribute VB_Name = "frmArcs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' All variables MUST be declared.
Option Explicit

Public Sub DrawArc()
    Dim X, Y
    Const PI = 3.14159265
    
    ' Calculate the center of the form.
    X = frmArcs.ScaleWidth / 2
    Y = frmArcs.ScaleHeight / 2
    
    ' Clear the form.
    frmArcs.Cls
    
    ' Draw an arc.
    Circle (X, Y), vsbRadius.Value * 20, , _
            -hsbFrom * 2 * PI / 260, -hsbTo * 2 * PI / 360
            
    ' Update the lblFrom label.
    lblFrom.Caption = "From: " + Str(hsbFrom.Value) + _
                            " degrees"
    
    ' Update the lblTo label.
    lblTo.Caption = "To: " + Str(hsbTo.Value) + _
                    " degrees"
    
    ' Update the lblRadius label.
    lblRadius.Caption = "Radius: " + Str(vsbRadius.Value * 20)
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub hsbFrom_Change()
    ' Execute the DrawArc procedure to draw the arc.
    DrawArc
End Sub

Private Sub hsbFrom_Scroll()
    ' Execute the DrawArc procedure to draw the arc.
    DrawArc
End Sub

Private Sub hsbTo_Change()
    ' Execute the DrawArc procedure to draw the arc.
    DrawArc
End Sub

Private Sub hsbTo_Scroll()
    ' Execute the DrawArc procedure to draw the arc.
    DrawArc
End Sub

Private Sub vsbRadius_Change()
    ' Execute the DrawArc procedure to draw the arc.
    DrawArc
End Sub

Private Sub vsbRadius_Scroll()
    ' Execute the DrawArc procedure to draw the arc.
    DrawArc
End Sub
