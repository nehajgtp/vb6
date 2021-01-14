VERSION 5.00
Begin VB.Form frmDrop 
   Caption         =   "The Drop Program"
   ClientHeight    =   3960
   ClientLeft      =   1200
   ClientTop       =   1635
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   ScaleHeight     =   3960
   ScaleWidth      =   6570
   Begin VB.PictureBox Picture1 
      DragMode        =   1  'Automatic
      Height          =   1080
      Left            =   2760
      ScaleHeight     =   1020
      ScaleWidth      =   1035
      TabIndex        =   2
      Tag             =   "Water image"
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox txtInfo 
      Alignment       =   2  'Center
      Enabled         =   0   'False
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
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Drop.frx":0000
      Top             =   240
      Width           =   4455
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
      Left            =   2520
      TabIndex        =   0
      Top             =   3240
      Width           =   1215
   End
End
Attribute VB_Name = "frmDrop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' All variables must be declared.
Option Explicit

Private Sub cmdExit_Click()
End
End Sub

Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
' Clear the text box.
txtInfo.Text = ""
' Move the control.
Source.Move X, Y
End Sub

Private Sub Form_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
Dim sInfo As String

' Display the dragging information
sInfo = "Now dragging "
sInfo = sInfo + Source.Tag
sInfo = sInfo + " over the Form."
sInfo = sInfo + " State = "
sInfo = sInfo + Str(State)
txtInfo.Text = sInfo
End Sub

Private Sub Form_Load()

End Sub
