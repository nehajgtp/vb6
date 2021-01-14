VERSION 5.00
Begin VB.Form frmFonts 
   Caption         =   "The Fonts Program"
   ClientHeight    =   4845
   ClientLeft      =   1200
   ClientTop       =   1635
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   ScaleHeight     =   4845
   ScaleWidth      =   6675
   Begin VB.CommandButton cmdNumberOfFonts 
      Caption         =   "&Number of Fonts"
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
      Left            =   2280
      TabIndex        =   3
      Top             =   3120
      Width           =   2175
   End
   Begin VB.ComboBox cboFontsPrinter 
      Height          =   315
      Left            =   3960
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   720
      Width           =   2415
   End
   Begin VB.ComboBox cboFontsScreen 
      Height          =   315
      Left            =   480
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   720
      Width           =   2415
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
      Height          =   495
      Left            =   5160
      TabIndex        =   0
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label lblSampleInfo 
      Caption         =   "Sample:"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label lblSample 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Aa Bb Cc Dd Ee Ff"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   4080
      Width           =   6135
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblPrinter 
      Caption         =   "Available Printer Fonts:"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   5
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label lblScreen 
      Caption         =   "Available Screen Fonts:"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   480
      Width           =   2535
   End
End
Attribute VB_Name = "frmFonts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' All variables MUST be declared.
Option Explicit
Dim gNumOfScreenFonts
Dim gNumOfPrinterFonts

Private Sub cboFontsScreen_Click()
    ' User selected a new screen font. Change the
    ' font of the label in accordance with the
    ' user's font selection.
    lblSample.FontName = cboFontsScreen.Text
    
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdNumberOfFonts_Click()
    ' Display the number of screen fonts in the
    ' system.
    MsgBox "Number of Screen fonts: " + _
            Str(gNumOfScreenFonts)
End Sub

Private Sub Form_Load()
    Dim I
    
    ' Calculate the number of screen fonts.
    gNumOfScreenFonts = Screen.FontCount - 1
    
    ' Calculate the number of printer fonts.
    gNumOfPrinterFonts = Printer.FontCount - 1
    
    ' Fill the items of the combo box with the
    ' screen fonts.
    For I = 0 To gNumOfScreenFonts - 1 Step 1
        cboFontsScreen.AddItem Screen.Fonts(I)
    Next
    
    ' Fill the items of the combo box with the
    ' printer fonts.
    For I = 0 To gNumOfPrinterFonts - 1 Step 1
        cboFontsPrinter.AddItem Printer.Fonts(I)
    Next
    
    ' initialize the text of the combo box
    ' to item #0.
    cboFontsScreen.ListIndex = 0
    ' Initialize the label font to value of the
    ' combo box.
    lblSample.FontName = cboFontsScreen.Text
    
    
End Sub
