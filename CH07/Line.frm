VERSION 5.00
Begin VB.Form frmLine 
   BackColor       =   &H00FFFFFF&
   Caption         =   "The Line Program"
   ClientHeight    =   4320
   ClientLeft      =   1200
   ClientTop       =   1635
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   ScaleHeight     =   4320
   ScaleWidth      =   6735
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
      Left            =   2640
      TabIndex        =   1
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start"
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
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Line linLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   10
      X1              =   2760
      X2              =   3960
      Y1              =   1800
      Y2              =   2280
   End
End
Attribute VB_Name = "frmLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
 End
End Sub

Private Sub cmdStart_Click()
    ' Set the start and endpoints of the line
    ' constrol to random values.
    linLine.BorderColor = RGB(Int(255 * Rnd), Int(255 * Rnd), Int(255 * Rnd))
    linLine.BorderWidth = Int(100 * Rnd) + 1
    
    linLine.X1 = Int(frmLine.Width * Rnd)
    linLine.Y1 = Int(frmLine.Height * Rnd)
    linLine.X2 = Int(frmLine.Width * Rnd)
    linLine.Y2 = Int(frmLine.Height * Rnd)
End Sub
