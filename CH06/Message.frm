VERSION 5.00
Begin VB.Form frmMessage 
   Caption         =   "The Message Program"
   ClientHeight    =   3960
   ClientLeft      =   1200
   ClientTop       =   1635
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   ScaleHeight     =   3960
   ScaleWidth      =   6570
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
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdMessage 
      Caption         =   "&Message"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1920
      TabIndex        =   0
      Top             =   1320
      Width           =   2775
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' All variables MUST be declared.
Option Explicit

Private Sub cmdExit_Click()
Dim Message As String
Dim DialogType As Integer
Dim Title As String
Dim Response As Integer

' The message of the dialog box.
Message = "Are you sure you want to quit?"

' The dialog box should have Yes and No buttons,
' and a question icon.
DialogType = vbYesNo + vbQuestion

' The title of the dialog box.
Title = "The Message Program"

' Display the dialog box and get user's response.
Response = MsgBox(Message, DialogType, Title)

' Evaluate the user's response.
If Response = vbYes Then
    End
End If
End Sub

Private Sub cmdMessage_Click()
Dim Message As String
Dim DialogType As Integer
Dim Title As String
Dim Response As Integer

' The message of the dialog box.
Message = "This is a sample message!"

' The dialog box should have an OK button and
' an exclamation icon.
DialogType = vbOKOnly + vbExclamation

' The title of the dialog box
Title = "Dialog Box Demonstration"

' Display the dialog box.
MsgBox Message, DialogType, Title
End Sub

