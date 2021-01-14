VERSION 5.00
Begin VB.Form frmMultiply 
   Caption         =   "The Multiply Program"
   ClientHeight    =   4770
   ClientLeft      =   1200
   ClientTop       =   6690
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   ScaleHeight     =   4770
   ScaleWidth      =   6165
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "&Calculate"
      Height          =   1095
      Left            =   2280
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtResult 
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   5055
   End
   Begin VB.Label lblResult 
      Caption         =   "Result:"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "frmMultiply"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' All variables MUST be declared.
Option Explicit
Private Sub cmdCalculate_Click()
    txtResult.Text = Str(Multiply(2, 3))
End Sub

Private Sub Label1_Click()

End Sub

Private Sub cmdExit_Click()
 End
End Sub


Public Function Multiply(X As Integer, Y As Integer)
     Dim Z
     Z = X * Y
     Multiply = Z
End Function
