VERSION 5.00
Begin VB.Form frmClip 
   Caption         =   "The Clip Program"
   ClientHeight    =   4245
   ClientLeft      =   1200
   ClientTop       =   1935
   ClientWidth     =   10575
   LinkTopic       =   "Form1"
   ScaleHeight     =   4245
   ScaleWidth      =   10575
   Begin VB.TextBox txtUserArea 
      Height          =   2775
      Left            =   2640
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   4215
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
   End
End
Attribute VB_Name = "frmClip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
    ' Make the text box cover the entire form area.
    txtUserArea.Width = frmClip.ScaleWidth
    txtUserArea.Height = frmClip.ScaleHeight
End Sub

Private Sub mnuCopy_Click()
    ' Clear the clipboard.
    Clipboard.Clear
    ' Transfer to the clipboard the currently
    ' selected text of the text box.
    Clipboard.SetText txtUserArea.SelText
End Sub

Private Sub mnuCut_Click()
    ' Clear the clipboard.
    Clipboard.Clear
    ' Transfer to the clipboard the currently
    ' selected text of the text box.
    Clipboard.SetText txtUserArea.SelText
    ' Replace the currently selected text of the
    ' text box with null.
    txtUserArea.SelText = " "
End Sub


Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuPaste_Click()
    ' Replace the currently selected area of the
    ' text box with the content of the clipboard.
    ' If nothing is selected in the text box,
    ' transfer the text of the clipboard to the text
    ' box at the current location of the cursor.
    txtUserArea.SelText = Clipboard.GetText()
End Sub
