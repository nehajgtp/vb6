VERSION 5.00
Begin VB.Form frmPoints 
   BackColor       =   &H00FFFFFF&
   Caption         =   "The Points Program"
   ClientHeight    =   4545
   ClientLeft      =   1200
   ClientTop       =   1935
   ClientWidth     =   10215
   LinkTopic       =   "Form1"
   ScaleHeight     =   4545
   ScaleWidth      =   10215
   Begin VB.Timer tmrTimer 
      Interval        =   60
      Left            =   600
      Top             =   2400
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuGraphics 
      Caption         =   "&Graphics"
      Begin VB.Menu mnuDrawPoints 
         Caption         =   "&Draw Points"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "&Clear"
      End
      Begin VB.Menu mnuLines 
         Caption         =   "&Lines"
      End
   End
   Begin VB.Menu mnuDrawBox 
      Caption         =   "D&raw Box"
      Begin VB.Menu mnuRed 
         Caption         =   "R&ed"
      End
      Begin VB.Menu mnuGreen 
         Caption         =   "Gree&n"
      End
      Begin VB.Menu mnuBlue 
         Caption         =   "&Blue"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSetStyle 
         Caption         =   "&Set Style"
      End
   End
End
Attribute VB_Name = "frmPoints"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' All variables MUST be declared.
Option Explicit

' A flag that determines if points will be drawn.
Dim gDrawPoints

Private Sub Form_Load()
    ' Disable drawing.
    gDrawPoints = 0
End Sub

Private Sub mnuBlue_Click()
    ' Set the FillColor property of the form
    frmPoints.FillColor = RGB(255, 0, 0)
    
    ' Draw the box.
    frmPoints.Line (100, 80)-Step(5000, 3000), _
    RGB(0, 0, 255), B
End Sub

Private Sub mnuClear_Click()
    ' Disable drawing.
    gDrawPoints = 0
    
    ' Clear the form.
    frmPoints.Cls
End Sub

Private Sub mnuDrawPoints_Click()
    ' Enable drawing
    gDrawPoints = 1
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuGreen_Click()
    ' Set the FillColor property of the form.
    frmPoints.FillColor = RGB(255, 0, 0)
    
    ' Draw the box.
    frmPoints.Line (100, 80)-Step(5000, 3000), _
    RGB(0, 255, 0), B
End Sub

Private Sub mnuLines_Click()
    Dim Counter
    
    For Counter = 1 To 100 Step 1
        Line -(Rnd * frmPoints.ScaleWidth, _
        Rnd * frmPoints.ScaleHeight), RGB(0, 0, 0)
    Next
End Sub

Private Sub mnuRed_Click()
    ' Set the FillColor property of the form.
    frmPoints.FillColor = RGB(255, 0, 0)
    
    ' Draw the box
    frmPoints.Line (100, 80)-Step(5000, 3000), _
    RGB(255, 0, 0), B
End Sub

Private Sub mnuSetStyle_Click()
    Dim FromUser
    Dim Instruction
    Instruction = "Enter a number between 0 and 7 " + _
                "for the FillStyle"
    
    ' Get from the user the desired FillStyle.
    FromUser = InputBox$(Instruction, _
                "Setting the FillStyle")
    
    ' Clear the form
    frmPoints.Cls
    
    ' Did the user enter a valid FillStyle?
    If Val(FromUser) >= 0 And Val(FromUser) <= 7 Then
        frmPoints.FillStyle = Val(FromUser)
    Else
        Beep
        MsgBox ("Invalid FillStyle")
    End If
    
    ' Draw the box.
    frmPoints.Line (100, 80)-Step(5000, 3000), _
        RGB(0, 0, 0), B
End Sub

Private Sub tmrTimer_Timer()
    Dim R, G, B
    Dim X, Y
    Dim Counter
    
    ' Is it OK to draw?
    If gDrawPoints = 1 Then
        ' Draw 100 points.
        For Counter = 1 To 100 Step 1
            ' Get a random color.
            R = Rnd * 255
            G = Rnd * 255
            B = Rnd * 255
            frmPoints.PSet Step(1, 1), RGB(R, G, B)
            If CurrentX >= frmPoints.ScaleWidth Then
                CurrentX = Rnd * frmPoints.ScaleWidth
            End If
            If CurrentY >= frmPoints.ScaleHeight Then
                CurrentY = Rnd * frmPoints.ScaleHeight
            End If
        Next
    End If
End Sub
