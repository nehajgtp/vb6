VERSION 5.00
Begin VB.Form frmSpeed 
   Caption         =   "The Speed Program"
   ClientHeight    =   3960
   ClientLeft      =   1200
   ClientTop       =   1635
   ClientWidth     =   4155
   LinkTopic       =   "Form1"
   ScaleHeight     =   3960
   ScaleWidth      =   4155
   Begin VB.TextBox txtSpeed 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1440
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "Speed.frx":0000
      Top             =   240
      Width           =   1215
   End
   Begin VB.HScrollBar hsbSpeed 
      Height          =   255
      Left            =   600
      Max             =   100
      TabIndex        =   1
      Tag             =   "1440"
      Top             =   1800
      Value           =   50
      Width           =   2775
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   3000
      Width           =   1215
   End
End
Attribute VB_Name = "frmSpeed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
End
End Sub

Private Sub hsbSpeed_Change()
txtSpeed.Text = Str(hsbSpeed.Value) + ".mph"
End Sub
