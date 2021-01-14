VERSION 5.00
Begin VB.Form frmGrow 
   BackColor       =   &H00FFFFFF&
   Caption         =   "The Grow Program"
   ClientHeight    =   3015
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuGrow 
      Caption         =   "&Grow"
      Begin VB.Menu mnuAdd 
         Caption         =   "&Add"
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "&Remove"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
      Begin VB.Menu mnuItems 
         Caption         =   "-"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmGrow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' All variables MUST be declared.
Option Explicit

'Declare the gLastElement variable.
Dim gLastElement As Integer

Private Sub Form_Load()
' Initally the last element in the
' mnuItems[] array is 0.
gLastElement = 0

' Initially no items are added to the
' Grow menu, so disable the Remove option.

mnuRemove.Enabled = False
End Sub

Private Sub mnuAdd_Click()
' Increment the gLastElement variable.
gLastElement = gLastElement + 1

' Add a new element to the mnuItems[] array.
Load mnuItems(gLastElement)

' Assign a caption to the item that
' was just added.
mnuItems(gLastElement).Caption = _
                        "Item " + Str(gLastElement)

' Because an element was just added to the
' mnuItems array, the Remove option should be enabled.
mnuRemove.Enabled = True

End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuItems_Click(Index As Integer)
    ' Display the item that was selected.
    MsgBox "You selected Item " + Str(Index)
End Sub

Private Sub mnuRemove_Click()
' Remove the last element of the mnuItems array.
Unload mnuItems(gLastElement)

'Decrement the gLastElement variable.
gLastElement = gLastElement - 1

' If only element 0 is left in the array,
' disable the Remove menu item.
If gLastElement = 0 Then
    mnuRemove.Enabled = False
End If
End Sub
