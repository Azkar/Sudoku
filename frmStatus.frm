VERSION 5.00
Begin VB.Form frmStatus 
   ClientHeight    =   1770
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3255
   LinkTopic       =   "Form1"
   ScaleHeight     =   1770
   ScaleWidth      =   3255
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   735
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Me.Hide
End Sub

Private Sub Form_Load()
    Me.Left = frmSud.Left
    Me.Top = frmSud.Top
    
    Me.Left = frmSud.Width / 2
    Me.Top = frmSud.Height / 2
End Sub

