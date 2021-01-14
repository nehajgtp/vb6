VERSION 5.00
Begin VB.Form frmShowFont 
   Caption         =   "The ShowFont Program"
   ClientHeight    =   4245
   ClientLeft      =   1200
   ClientTop       =   1935
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   ScaleHeight     =   4245
   ScaleWidth      =   6570
   Begin VB.CheckBox chkUnderline 
      Caption         =   "&Underline"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CheckBox chkStrike 
      Caption         =   "&Strike"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CheckBox chkItalic 
      Caption         =   "&Italic"
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   4800
      TabIndex        =   2
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtTest 
      Height          =   2055
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Text            =   "ShowFont.frx":0000
      Top             =   240
      Width           =   5895
   End
   Begin VB.CheckBox chkBold 
      Caption         =   "&Bold"
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Menu mnuFonts 
      Caption         =   "&Fonts"
      Begin VB.Menu mnuCourier 
         Caption         =   "Courier"
      End
      Begin VB.Menu mnuMSSansSerif 
         Caption         =   "MS Sans Serif"
      End
   End
   Begin VB.Menu mnuSize 
      Caption         =   "&Size"
      Begin VB.Menu mnu10Points 
         Caption         =   "1&0 Points"
      End
      Begin VB.Menu mnu12Points 
         Caption         =   "1&2 Points"
      End
   End
End
Attribute VB_Name = "frmShowFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' All variables MUST be declared.
Option Explicit

Private Sub chkBold_Click()
' Update the FontBold property of the text
' box with the Value property of the
' chkBold check box.
txtTest.FontBold = chkBold.Value

End Sub

Private Sub chkItalic_Click()
' Update the FontItalic property of the
' text box with the Value property
' of the chkItalic check box.
txtTest.FontItalic = chkItalic.Value

End Sub

Private Sub chkStrike_Click()
' Update the FontStrikethru property
' of the text box with the Value property
' of the chkStrike check box.
txtTest.FontStrikethru = chkStrike.Value

End Sub

Private Sub chkUnderline_Click()
' Update the FontUnderline property
' of the text box with the Value
' property of the chkUnderline check box.
txtTest.FontUnderline = chkUnderline.Value

End Sub

Private Sub cmdExit_Click()
    
    End
    
End Sub

Private Sub mnu10Points_Click()
    
    ' Set the size of the font to 10 points
    txtTest.FontSize = 10
    
End Sub

Private Sub mnu12Points_Click()
    
    ' Set the size of the font to 12 points.
    txtTest.FontSize = 12
    
End Sub

Private Sub mnuCourier_Click()
    ' Set the font name to Courier.
    txtTest.FontName = "Courier"
    
End Sub

Private Sub mnuMSSansSerif_Click()
    ' Set the font name to MS Sans Serif.
    txtTest.FontName = "MS Sans Serif"
    
End Sub
