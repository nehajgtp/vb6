VERSION 5.00
Begin VB.Form frmMove 
   BackColor       =   &H0000FFFF&
   Caption         =   "The Move Program"
   ClientHeight    =   4860
   ClientLeft      =   1200
   ClientTop       =   1635
   ClientWidth     =   14280
   LinkTopic       =   "Form1"
   ScaleHeight     =   4860
   ScaleWidth      =   14280
   Begin VB.OptionButton optBell 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Bell"
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
      Left            =   360
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.OptionButton optClub 
      BackColor       =   &H0000FFFF&
      Caption         =   "C&lub"
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
      Left            =   360
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin VB.OptionButton optCup 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Cup"
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
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Value           =   -1  'True
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
      Height          =   495
      Left            =   2520
      TabIndex        =   0
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Image imgCup 
      Height          =   330
      Left            =   2400
      Top             =   1800
      Width           =   1215
   End
End
Attribute VB_Name = "frmMove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
