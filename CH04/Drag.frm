VERSION 5.00
Begin VB.Form frmDrag 
   Caption         =   "The Drag Program"
   ClientHeight    =   3960
   ClientLeft      =   1200
   ClientTop       =   1635
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   ScaleHeight     =   3960
   ScaleWidth      =   6570
   Begin VB.CommandButton cmdDragMe 
      Caption         =   "&Drag Me"
      DragMode        =   1  'Automatic
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
      Left            =   1080
      TabIndex        =   1
      Top             =   840
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
      Height          =   495
      Left            =   2880
      TabIndex        =   0
      Top             =   3360
      Width           =   1215
   End
End
Attribute VB_Name = "frmDrag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' All variables MUST be declared.
Option Explicit

Private Sub cmdExit_Click()
    End
End Sub

