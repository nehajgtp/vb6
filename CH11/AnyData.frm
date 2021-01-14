VERSION 5.00
Begin VB.Form frmAnyData 
   Caption         =   "The AnyData Program"
   ClientHeight    =   4245
   ClientLeft      =   1200
   ClientTop       =   1935
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   ScaleHeight     =   4245
   ScaleWidth      =   6570
   Begin VB.TextBox txtUserArea 
      Height          =   855
      Left            =   3840
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Tag             =   "The txtUserArea"
      Top             =   3000
      Width           =   2655
   End
   Begin VB.ComboBox cboList 
      Height          =   315
      Left            =   5160
      TabIndex        =   2
      Tag             =   "The cboList"
      Text            =   "Combo1"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.ListBox lstList 
      Height          =   1035
      Left            =   3720
      TabIndex        =   1
      Tag             =   "The lstList"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.PictureBox picMyPicture 
      Height          =   3125
      Left            =   120
      ScaleHeight     =   3060
      ScaleWidth      =   2715
      TabIndex        =   0
      Tag             =   "The picMyPicture"
      Top             =   720
      Width           =   2775
   End
   Begin VB.Menu menuExit 
      Caption         =   "Exit "
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
Attribute VB_Name = "frmAnyData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' All variables must be declared.
Option Explicit

Private Sub Form_Load()
    ' Fill three items inside the combo box.
    cboList.AddItem "Clock"
    cboList.AddItem "Cup"
    cboList.AddItem "Bell"
    
    ' Fill three items inside the list control.
    lstList.AddItem "One"
    lstList.AddItem "Two"
    lstList.AddItem "Three"
    
End Sub

Private Sub mnuCopy_Click()
    ' Clear the clipboard.
    Clipboard.Clear
    ' Find which is the currently active control, and
    ' copy its highlighted content to the clipboard.
    If TypeOf Screen.ActiveControl Is TextBox Then
        Clipboard.SetText Screen.ActiveControl.SelText
    ElseIf TypeOf Screen.ActiveControl Is ComboBox Then
        Clipboard.SetText Screen.ActiveControl.Text
    ElseIf TypeOf Screen.ActiveControl Is PictureBox Then
        Clipboard.SetData Screen.ActiveControl.Picture
    ElseIf TypeOf Screen.ActiveControl Is ListBox Then
        Clipboard.SetText Screen.ActiveControl.Text
    Else
        ' Do nothing
    End If
End Sub

Private Sub mnuCut_Click()
    ' Execute the mnuCopy_Click() procedure
    mnuCopy_Click
    ' Find which is the currently highlighted control,
    ' and remove its highlighted content.
    If TypeOf Screen.ActiveControl Is TextBox Then
        Screen.ActiveControl.SelText = ""
    ElseIf TypeOf Screen.ActiveControl Is ComboBox Then
        Screen.ActiveControl.Text = ""
    ElseIf TypeOf Screen.ActiveControl Is PictureBox Then
        Screen.ActiveControl.Picture = LoadPicture()
    ElseIf TypeOf Screen.ActiveControl Is ListBox Then
        If Screen.ActiveControl.ListIndex >= 0 Then
            Screen.ActiveControl.RemoveItem
                Screen.ActiveControl.ListIndex
        End If
    Else
        ' Do nothing
    End If
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuPaste_Click()
    ' Find which is the currently active control and
    ' paste the content of the clipboard to it.
    If TypeOf Screen.ActiveControl Is TextBox Then
        Screen.ActiveControl.SelText = Clipboard.GetText()
    ElseIf TypeOf Screen.ActiveControl Is ComboBox Then
        Screen.ActiveControl.Text = Clipboard.GetText()
    ElseIf TypeOf Screen.ActiveControl Is PictureBox Then
        Screen.ActiveControl.Picture = Clipboard.GetData()
    ElseIf TypeOf Screen.ActiveControl Is ListBox Then
        Screen.ActiveControl.AddItem Clipboard.GetText()
    Else
        ' Do nothing
    End If
End Sub

Private Sub picMyPicture_GotFocus()
    ' Change the BorderStyle so that user will be
    ' able to tell that the picture control got the
    ' focus (i.e., selected).
    picMyPicture.BorderStyle = 1
    
End Sub

Private Sub picMyPicture_LostFocus()
    ' Change the BorderStyle so that the user will be
    ' able to tell that the picture control lost the
    ' focus (i.e., not selected).
    picMyPicture.BorderStyle = 0
End Sub
