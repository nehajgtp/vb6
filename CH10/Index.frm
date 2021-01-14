VERSION 5.00
Begin VB.Form frmIndex 
   BackColor       =   &H00FFFFFF&
   Caption         =   "The Index Program"
   ClientHeight    =   4245
   ClientLeft      =   1200
   ClientTop       =   1935
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   ScaleHeight     =   4245
   ScaleWidth      =   6570
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuDisplayIndex 
         Caption         =   "&Display Index"
      End
      Begin VB.Menu mnuEraseCh2 
         Caption         =   "Erase Chapter 2"
      End
      Begin VB.Menu mnuClearText 
         Caption         =   "&Clear Text"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' All variables MUST be declared.
Option Explicit
Dim gDots

Private Sub Form_Load()
    ' Clear the form.
    frmIndex.Cls
End Sub

Private Sub mnuDisplayIndex_Click()
    frmIndex.Cls
    
    ' Heading should be displayed 100 twips from
    ' the top.
    CurrentY = 100
    
    ' Place the heading at the center of the row.
    CurrentX = (frmIndex.ScaleWidth - _
        frmIndex.TextWidth("Index")) / 2
    frmIndex.FontUnderline = True
    
    frmIndex.Print "Index"
    
    ' Display the chapters.
    frmIndex.FontUnderline = False
    CurrentY = frmIndex.TextHeight("VVV") * 2
    CurrentX = 100
    Print "Chapter 1" + gDots + "The world"
    CurrentY = frmIndex.TextHeight("VVV") * 3
    CurrentX = 100
    Print "Chapter 2" + gDots + "The chair"
    CurrentY = frmIndex.TextHeight("VVV") * 4
    CurrentX = 100
    Print "Chapter 3" + gDots + "The mouse"
    CurrentY = frmIndex.TextHeight("VVV") * 5
    CurrentX = 100
    Print "Chapter 4" + gDots + "The end"
    
    
End Sub

Private Sub mnuEraseCh2_Click()
Dim LengthOfLine
Dim HeightOfLine

' Erase the line by placing a box with white
' background over the line.
CurrentY = frmIndex.TextHeight("VVV") * 3
CurrentX = 100
LengthOfLine = frmIndex.TextWidth("Chapter 2" + _
               gDots + "The chair")
HeightOfLine = frmIndex.TextHeight("C")
frmIndex.Line -Step(LengthOfLine, HeightOfLine), _
              RGB(255, 255, 255), BF
End Sub

Private Sub mnuExit_Click()
    End
End Sub
