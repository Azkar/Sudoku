VERSION 5.00
Begin VB.Form frmAbout 
   Caption         =   "About Sudoku"
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   ScaleHeight     =   5640
   ScaleWidth      =   5055
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   615
      Left            =   1440
      TabIndex        =   0
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label lblInfo 
      Caption         =   $"frmAbout.frx":0000
      Height          =   855
      Left            =   480
      TabIndex        =   3
      Top             =   2400
      Width           =   3615
   End
   Begin VB.Label lblCreated 
      Alignment       =   2  'Center
      Caption         =   "Created By:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4815
   End
   Begin VB.Label lblMe 
      Alignment       =   2  'Center
      Caption         =   "Sean Madden"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   840
      Width           =   3735
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    Unload Me
End Sub
