VERSION 5.00
Begin VB.Form frmSum 
   Caption         =   "The Sum Program"
   ClientHeight    =   4440
   ClientLeft      =   1200
   ClientTop       =   1635
   ClientWidth     =   5205
   LinkTopic       =   "Form1"
   ScaleHeight     =   4440
   ScaleWidth      =   5205
   Begin VB.CommandButton cmdSumIt 
      Caption         =   "&Sum It"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   1920
      TabIndex        =   3
      Top             =   1560
      Width           =   1455
   End
   Begin VB.VScrollBar vsbNum 
      Height          =   2775
      Left            =   840
      Max             =   500
      Min             =   1
      TabIndex        =   2
      Top             =   1080
      Value           =   1
      Width           =   255
   End
   Begin VB.TextBox txtResult 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   495
      Left            =   1680
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   720
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
      Left            =   2040
      TabIndex        =   0
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label lblNum 
      Caption         =   "Selected Number: 1"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   360
      TabIndex        =   4
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "frmSum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    Beep
    End
End Sub

Private Sub cmdSumIt_Click()
    Dim I
    Dim R
    For I = 1 To vsbNum.Value Step 1
        R = R + I
    Next
    
    txtResult.Text = Str(R)
End Sub

Private Sub vsbNum_Change()
    lblNum = "Selected number: " + Str(vsbNum.Value)
End Sub

Private Sub vsbNum_Scroll()
    lblNum = "Selected number: " + Str(vsbNum.Value)
End Sub
