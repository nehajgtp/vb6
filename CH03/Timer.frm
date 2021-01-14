VERSION 5.00
Begin VB.Form frmTimer 
   Caption         =   "The Timer Program"
   ClientHeight    =   3705
   ClientLeft      =   1200
   ClientTop       =   1635
   ClientWidth     =   3090
   LinkTopic       =   "Form1"
   ScaleHeight     =   3705
   ScaleWidth      =   3090
   Begin VB.CommandButton cmdEnableDisable 
      Caption         =   "&Enable"
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Timer tmrTimer 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   240
      Top             =   2040
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   3120
      Width           =   1215
   End
End
Attribute VB_Name = "frmTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' All variables MUST be declared.
Option Explicit

Private Sub cmdEnableDisable_Click()
If tmrTimer.Enabled = True Then
    tmrTimer.Enabled = False
    cmdEnableDisable.Caption = "&Enable"
Else
    tmrTimer.Enabled = True
    cmdEnableDisable.Caption = "&Disable"
End If
End Sub

Private Sub cmdExit_Click()
 End
End Sub

Private Sub tmrTimer_Timer()
' If the gKeepTrack variable is equal to 1,
' then beep.
If tmrTimer.Enabled = True Then
    Beep
End If
End Sub
