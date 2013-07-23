VERSION 5.00
Begin VB.Form frmHow 
   Caption         =   "How to Play"
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9915
   LinkTopic       =   "Form1"
   ScaleHeight     =   6555
   ScaleWidth      =   9915
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblHow 
      Height          =   5175
      Left            =   840
      TabIndex        =   0
      Top             =   600
      Width           =   8295
   End
End
Attribute VB_Name = "frmHow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Open App.Path & "/how.txt" For Input As #1
    Input #1, x
    lblHow.Caption = x
End Sub
