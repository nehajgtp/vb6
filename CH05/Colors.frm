VERSION 5.00
Begin VB.Form frmColors 
   BackColor       =   &H00FFFFFF&
   Caption         =   "The Colors Program"
   ClientHeight    =   3570
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   ScaleHeight     =   3570
   ScaleWidth      =   6195
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuColors 
      Caption         =   "Colors"
      Begin VB.Menu mnuSetColor 
         Caption         =   "&Set Color"
         Begin VB.Menu mnuRed 
            Caption         =   "&Red"
            Shortcut        =   ^R
         End
         Begin VB.Menu mnuBlue 
            Caption         =   "&Blue"
            Shortcut        =   ^B
         End
         Begin VB.Menu mnuWhite 
            Caption         =   "&White"
         End
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuSize 
      Caption         =   "&Size"
      Begin VB.Menu mnuSmall 
         Caption         =   "&Small"
      End
      Begin VB.Menu mnuLarge 
         Caption         =   "&Large"
      End
   End
End
Attribute VB_Name = "frmColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
' Because initially the window is white.
' disable the White menu item.
mnuWhite.Enabled = False

'Because initially the window is small,
'disable the Small menu itme.
mnuSmall.Enabled = False
End Sub

Private Sub mnuBlue_Click()
' Set the color of the form to blue.
frmColors.BackColor = QBColor(1)

' Disable the Blue menu item.
mnuBlue.Enabled = False

'Enable the Red and White menu items.
mnuRed.Enabled = True
mnuWhite.Enabled = True
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuLarge_Click()
' Set the size of the form to large.
frmColors.WindowState = 2

' Disable the Large menu item.
mnuLarge.Enabled = False

'Enable the Small menu item.
mnuSmall.Enabled = True

End Sub

Private Sub mnuRed_Click()
' Set the color of the form to red.
frmColors.BackColor = QBColor(4)

'Disable the Red menu item
mnuRed.Enabled = False

' Enable the Blue and White menu items.
mnuBlue.Enabled = True
mnuWhite.Enabled = True
End Sub

Private Sub mnuSmall_Click()
' Set the size of the form to small.
frmColors.WindowState = 0

' Disable the Small menu item.
mnuSmall.Enabled = False

' Enable the Large menu item
mnuLarge.Enabled = True

End Sub

Private Sub mnuWhite_Click()
' Set the color of the form to bright white.
frmColors.BackColor = QBColor(15)

'Disable the White menu item.
mnuWhite.Enabled = False

' Enable the Red and Blue menu items.
mnuRed.Enabled = True
mnuBlue.Enabled = True
End Sub
