VERSION 5.00
Begin VB.Form frmDialogs 
   BackColor       =   &H00FFFFFF&
   Caption         =   "The Dialogs Program"
   ClientHeight    =   3015
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuDialogs 
      Caption         =   "Dialogs"
      Begin VB.Menu mnuOkCancel 
         Caption         =   "OK-Cancel dialog"
      End
      Begin VB.Menu mnuAbortRetryIgnore 
         Caption         =   "Abort-Retry-Ignore dialog"
      End
      Begin VB.Menu mnuYesNoCancel 
         Caption         =   "Yes-No-Cancel dialog"
      End
      Begin VB.Menu mnuYesNo 
         Caption         =   "Yes-No dialog"
      End
      Begin VB.Menu mnuRetryCancel 
         Caption         =   "Retry-Cancel dialog"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmDialogs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' All variables MUST be declared
Option Explicit

Private Sub mnuAbortRetryIgnore_Click()
Dim DialogType As Integer
Dim DialogTitle As String
Dim DialogMsg As String
Dim Response As Integer

' Dialog should have Abort, Retry, Ignore buttons,
' and an Exclamation icon.
DialogType = vbAbortRetryIgnore + vbExclamation

' The dialog title.
DialogTitle = "MsgBox Demonstration"

' The dialog message.
DialogMsg = "This is a sample message!"

' Display the dialog box, and get user's response.
Response = MsgBox(DialogMsg, DialogType, DialogTitle)

' Evaluate the user's response.
Select Case Response
    Case vbAbort
        MsgBox "You clicked the Abort button!"
    Case vbRetry
        MsgBox "You clicked the Retry button!"
    Case vbIgnore
        MsgBox "You clicked the Ignore button!"
End Select
End Sub

Private Sub mnuExit_Click()
Dim DialogType As Integer
Dim DialogTitle As String
Dim DialogMsg As String
Dim Response As Integer

' Dialog should have Yes & No buttons,
' and a Critical Message icon.
DialogType = vbYesNo + vbCritical

' The dialog message.
DialogMsg = "Are you sure you want to exit?"

' Display the dialog box, and get user's response.
Response = MsgBox(DialogMsg, DialogType, DialogTitle)

' Evaluate the user's response.
If Response = vbYes Then
    End
End If
End Sub

Private Sub mnuOkCancel_Click()
Dim DialogType As Integer
Dim DialogTitle As String
Dim DialogMsg As String
Dim Response As Integer

' Dialog should have OK & Cancel buttons,
' and an Exclamation icon.
DialogType = vbOKCancel + vbExclamation

' The dialog title.
DialogTitle = "MsgBox Demonstration"

' The dialog message.
DialogMsg = "This is a sample message!"

' Display the dialog box, and get user's response.
Response = MsgBox(DialogMsg, DialogType, DialogTitle)

' Evaluate the user's response.
If Response = vbOK Then
    MsgBox "You clicked the OK button!"
Else
    MsgBox "You clicked the Cancel button!"
End If
End Sub

Private Sub mnuRetryCancel_Click()
Dim DialogType As Integer
Dim DialogTitle As String
Dim DialogMsg As String
Dim Response As Integer

' Dialog should have Retry & Cancel buttons,
' and an Exclamation icon.
DialogType = vbRetryCancel + vbExclamation

' The dialog title.
DialogTitle = "MsgBox Demonstration"

' The dialog message.
DialogMsg = "This is a sample message!"

' Display the dialog box, and get user's response.
Response = MsgBox(DialogMsg, DialogType, DialogTitle)

' Evaluate the user's response.
If Response = vbRetry Then
    MsgBox "You clicked the Retry button!"
Else
    MsgBox "You clicked the Cancel button!"
End If
End Sub

Private Sub mnuYesNo_Click()
Dim DialogType As Integer
Dim DialogTitle As String
Dim DialogMsg As String
Dim Response As Integer

' Dialog should have Yes & No buttons,
' and a question mark icon.
DialogType = vbYesNo + vbQuestion

' The dialog title.
DialogTitle = "MsgBox Demonstration"

' The dialog message.
DialogMsg = "Is this a sample message?"

' Display the dialog box, and get user's response.
Response = MsgBox(DialogMsg, DialogType, DialogTitle)

' Evaluate the user's response.
If Response = vbYes Then
    MsgBox "You clicked the Yes button!"
Else
    MsgBox "You clicked the No button!"
End If
End Sub

Private Sub mnuYesNoCancel_Click()
Dim DialogType As Integer
Dim DialogTitle As String
Dim DialogMsg As String
Dim Response As Integer

' Dialog should have Yes, No, and Cancel buttons,
' and an Exclamation icon.
DialogType = vbYesNoCancel + vbExclamation

' The dialog title.
DialogTitle = "MsgBox Demonstration"

' The dialog message.
DialogMsg = "This is a sample message!"

' Display the dialog box, and get user's response.
Response = MsgBox(DialogMsg, DialogType, DialogTitle)

' Evaluate the user's response
Select Case Response
Case vbYes
    MsgBox "You clicked the Yes button!"
Case vbNo
    MsgBox "You clicked the No button!"
Case vbCancel
    MsgBox "You clicked the Cancel button!"
End Select
End Sub
