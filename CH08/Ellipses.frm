VERSION 5.00
Begin VB.Form frmEllipses 
   Caption         =   "The Ellipses Program"
   ClientHeight    =   4320
   ClientLeft      =   1200
   ClientTop       =   1635
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   ScaleHeight     =   4320
   ScaleWidth      =   8460
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
      Left            =   3480
      TabIndex        =   4
      Top             =   3480
      Width           =   1215
   End
   Begin VB.HScrollBar hsbRadius 
      Height          =   255
      Left            =   1200
      Max             =   100
      Min             =   1
      TabIndex        =   1
      Top             =   3000
      Value           =   1
      Width           =   4575
   End
   Begin VB.HScrollBar hsbAspect 
      Height          =   255
      Left            =   1200
      Max             =   100
      Min             =   1
      TabIndex        =   0
      Top             =   2280
      Value           =   1
      Width           =   4575
   End
   Begin VB.Label lblInfo 
      Caption         =   "Aspect:"
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
      Left            =   6120
      TabIndex        =   5
      Top             =   120
      Width           =   1695
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
      Left            =   1200
      TabIndex        =   3
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label lblAspect 
      Caption         =   "Aspect:"
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
      Left            =   1200
      TabIndex        =   2
      Top             =   1920
      Width           =   1695
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmEllipses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' All variables MUST be declared.
Option Explicit

Private Sub cmdExit_Click()
 End
End Sub

Private Sub Form_Load()
    ' Initialize the radius and aspect scroll bars.
    hsbRadius.Value = 10
    hsbAspect.Value = 10
    
    ' Initialize the info label.
    lblInfo.Caption = "Aspect: 1"
    
    ' Set the DrawWidth property on the form.
    frmEllipses.DrawWidth = 2
End Sub

Private Sub hsbAspect_Change()
    Dim X, Y
    Dim Info
    ' Calculate the center of the form.
    X = frmEllipses.ScaleWidth / 2
    Y = frmEllipses.ScaleHeight / 2
    
    ' Clear the form
    frmEllipses.Cls
    
    ' Draw the ellipse.
    frmEllipses.Circle (X, Y), hsbRadius.Value * 10, _
                RGB(255, 0, 0), , , hsbAspect.Value / 10
    
    ' Prepare the Info string.
    Info = "Aspect: " + Str(hsbAspect.Value / 10)
    
    ' Display the value of the aspect.
    frmEllipses.lblInfo.Caption = Info
End Sub

Private Sub hsbAspect_Scroll()
    Dim X, Y
    Dim Info
    ' Calculate the center of the form.
    X = frmEllipses.ScaleWidth / 2
    Y = frmEllipses.ScaleHeight / 2
    
    ' Clear the form.
    frmEllipses.Cls
    
    ' Draw the ellipse.
    frmEllipses.Circle (X, Y), hsbRadius.Value * 10, _
            RGB(255, 0, 0), , , hsbAspect.Value / 10
            
    ' Prepare the Info string.
    Info = "Aspect: " + Str(hsbAspect.Value / 10)
    
    ' Display the value of the aspect
    frmEllipses.lblInfo.Caption = Info
End Sub

Private Sub hsbRadius_Change()
    Dim X, Y
    Dim Info
    
    X = frmEllipses.ScaleWidth / 2
    Y = frmEllipses.ScaleHeight / 2
    
    frmEllipses.Cls
    frmEllipses.Circle (X, Y), hsbRadius.Value * 10, _
                RGB(255, 0, 0), , , hsbAspect.Value / 10
    
    Info = "Aspect: " + Str(hsbAspect.Value / 10)
    
    frmEllipses.lblInfo.Caption = Info
End Sub

Private Sub hsbRadius_Scroll()
    Dim X, Y
    Dim Info
    
    X = frmEllipses.ScaleWidth / 2
    Y = frmEllipses.ScaleHeight / 2
    
    frmEllipses.Cls
    frmEllipses.Circle (X, Y), hsbRadius.Value * 10, _
                RGB(255, 0, 0), , , hsbAspect.Value / 10
            
    Info = "Aspect: " + Str(hsbAspect.Value / 10)
    
    frmEllipses.lblInfo.Caption = Info
End Sub
