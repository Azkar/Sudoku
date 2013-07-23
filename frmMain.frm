VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Sudoku"
   ClientHeight    =   3435
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3465
   LinkTopic       =   "Form3"
   ScaleHeight     =   3435
   ScaleWidth      =   3465
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHow 
      Caption         =   "How to Play"
      Height          =   615
      Left            =   600
      TabIndex        =   3
      Top             =   2640
      Width           =   2055
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      Height          =   615
      Left            =   600
      TabIndex        =   2
      Top             =   1920
      Width           =   2055
   End
   Begin VB.ComboBox cboDiff 
      Height          =   315
      ItemData        =   "frmMain.frx":0000
      Left            =   600
      List            =   "frmMain.frx":0010
      TabIndex        =   1
      Text            =   "Difficulty"
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label lblName 
      Caption         =   "Sean's Sudoku"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdHow_Click()
    frmHow.Show
End Sub

Private Sub cmdPlay_Click()
    If cboDiff <> "Difficulty" Then
        Me.Hide
        frmSud.Show
        frmSplash.Show
        frmSplash.Left = frmSud.Left
        frmSplash.Top = frmSud.Top
        
        Select Case LCase(cboDiff.Text)
            Case "easy"
                frmSud.diff = 1
            Case "medium"
                frmSud.diff = 2
            Case "hard"
                frmSud.diff = 3
            Case "insanity"
                frmSud.diff = 4
        End Select
        
    End If
End Sub

