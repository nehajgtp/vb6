VERSION 5.00
Begin VB.Form frmHello 
   BackColor       =   &H00C000C0&
   Caption         =   "The Hello Program"
   ClientHeight    =   6555
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13575
   LinkTopic       =   "Form1"
   ScaleHeight     =   6555
   ScaleWidth      =   13575
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDisplay 
      Alignment       =   2  'Center
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
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "Hello.frx":0000
      Top             =   1320
      Width           =   6855
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
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
      Left            =   7680
      TabIndex        =   2
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdHello 
      Caption         =   "&Display Hello"
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
      TabIndex        =   1
      Top             =   3000
      Width           =   1215
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
      Height          =   735
      Left            =   5520
      TabIndex        =   0
      Top             =   4320
      Width           =   1455
   End
End
Attribute VB_Name = "frmHello"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub cmdClear_Click()
txtDisplay.Text = ""
End Sub

Private Sub cmdExit_Click()
Beep
End
End Sub

Private Sub Command2_Click()

End Sub

Private Sub cmdHello_Click()
txtDisplay.Text = "Hello World!"
End Sub

