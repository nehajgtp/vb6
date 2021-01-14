VERSION 5.00
Begin VB.Form frmOption 
   BackColor       =   &H000000FF&
   Caption         =   "The Option Program"
   ClientHeight    =   4815
   ClientLeft      =   1200
   ClientTop       =   1635
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   8535
   Begin VB.OptionButton optLevel3 
      BackColor       =   &H000000FF&
      Caption         =   "Level &3"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   1440
      Width           =   1215
   End
   Begin VB.OptionButton optLevel1 
      BackColor       =   &H000000FF&
      Caption         =   "Level &1"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   480
      Width           =   1215
   End
   Begin VB.OptionButton optLevel2 
      BackColor       =   &H000000FF&
      Caption         =   "Level &2"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.CheckBox chkColors 
      BackColor       =   &H000000FF&
      Caption         =   "&Colors"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin VB.CheckBox chkMouse 
      BackColor       =   &H000000FF&
      Caption         =   "&Mouse"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.CheckBox chkSound 
      BackColor       =   &H000000FF&
      Caption         =   "&Sound"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Top             =   1440
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
      Height          =   1095
      Left            =   4800
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblChoice 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   360
      TabIndex        =   7
      Top             =   2280
      Width           =   3015
   End
End
Attribute VB_Name = "frmOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkColors_Click()
    UpdateLabel
End Sub

Private Sub chkMouse_Click()
    UpdateLabel
End Sub

Private Sub chkSound_Click()
    UpdateLabel
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub Form_Load()

End Sub

Private Sub optLevel1_Click()
    UpdateLabel
End Sub

Private Sub optLevel2_Click()
    UpdateLabel
End Sub

Private Sub optLevel3_Click()
    UpdateLabel
End Sub

Public Sub UpdateLabel()
' Declare the variables
Dim Info
Dim LFCR

LFCR = Chr(13) + Chr(10)

' Sound
If chkSound.Value = 1 Then
    Info = "Sound: ON"
Else
    Info = "Sound: OFF"
End If

' Mouse
If chkMouse.Value = 1 Then
    Info = Info + LFCR + "Mouse: ON"
Else
    Info = Info + LFCR + "Mouse: OFF"
End If

' Colors
If chkColors.Value = 1 Then
    Info = Info + LFCR + "Colors: ON"
Else
    Info = Info + LFCR + "Colors: OFF"
End If

' Level 1
If optLevel1.Value = True Then
    Info = Info + LFCR + "Level:1"
End If

If optLevel2.Value = True Then
    Info = Info + LFCR + "Level:2"
End If

' Level 3
If optLevel3.Value = True Then
    Info = Info + LFCR + "Level:3"
End If

lblChoice.Caption = Info

End Sub
