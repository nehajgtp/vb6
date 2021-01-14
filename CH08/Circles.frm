VERSION 5.00
Begin VB.Form frmCircles 
   Caption         =   "The Circles Program"
   ClientHeight    =   4800
   ClientLeft      =   1200
   ClientTop       =   1635
   ClientWidth     =   5505
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   5505
   Begin VB.CommandButton cmdDrawStyle 
      Caption         =   "&Draw Style"
      Height          =   1095
      Left            =   4320
      TabIndex        =   4
      Top             =   1320
      Width           =   1095
   End
   Begin VB.VScrollBar vsbRadius 
      Height          =   2295
      Left            =   240
      Max             =   100
      Min             =   1
      TabIndex        =   2
      Top             =   960
      Value           =   1
      Width           =   255
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
      Left            =   2160
      TabIndex        =   1
      Top             =   3840
      Width           =   1215
   End
   Begin VB.PictureBox picCircles 
      BackColor       =   &H00FFFFFF&
      Height          =   2295
      Left            =   480
      ScaleHeight     =   2235
      ScaleWidth      =   3675
      TabIndex        =   0
      Top             =   960
      Width           =   3735
   End
   Begin VB.Label lblRadius 
      Caption         =   "Radius"
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
      TabIndex        =   3
      Top             =   480
      Width           =   1695
   End
End
Attribute VB_Name = "frmCircles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' All variables MUST be declared
Option Explicit

Private Sub cmdDrawStyle_Click()
    Dim TheStyle
    ' Get a number from the user.
    TheStyle = InputBox$("Enter DrawStyle (0-6):")
    
    ' Is the number between 0 and 6?
    If Val(TheStyle) < 0 Or Val(TheStyle) > 6 Then
        ' The entered number is not within the valid
        ' range
        Beep
        MsgBox ("Invalid DrawStyle")
    Else
        ' The entered number is within the valid
        ' range, so change to DrawStyle property.
        picCircles.DrawStyle = Val(TheStyle)
    End If
    
End Sub

Private Sub cmdExit_Click()
    End
End Sub


Private Sub vsbRadius_Change()
Dim X, Y, Radius
Static LastValue
Dim R, G, B

' Generate random colors.
R = Rnd * 255
G = Rnd * 255
B = Rnd * 255

' Calculate the coordinate of the center of the
' picture control
X = picCircles.ScaleWidth / 2
Y = picCircles.ScaleHeight / 2

' If scroll bar was decrement, then clear the
' picture box.
If LastValue > vsbRadius.Value Then
    picCircles.Cls
End If

' Draw the circle.
picCircles.Circle (X, Y), vsbRadius.Value * 10, _
    RGB(R, G, B)

' Update LastValue for next time.
LastValue = vsbRadius.Value
End Sub

Private Sub vsbRadius_Scroll()
Dim X, Y, Radius
Static LastValue
Dim R, G, B

' Generate random colors.
R = Rnd * 255
G = Rnd * 255
B = Rnd * 255

' Calculate the coordinate of the center of the
' picture control.
X = picCircles.ScaleWidth / 2
Y = picCircles.ScaleHeight / 2

' If scroll bar was decremented, then clear the
' picture box.
If LastValue > vsbRadius Then
    picCircles.Cls
End If

' Draw the circle.
picCircles.Circle (X, Y), vsbRadius.Value * 10, _
            RGB(R, G, B)

' Update LastValue for next time.
LastValue = vsbRadius.Value
End Sub
