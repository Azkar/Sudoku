VERSION 5.00
Begin VB.Form frmSud 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sudoku"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4950
   FillColor       =   &H80000012&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   4950
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSolve 
      Caption         =   "Solution"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   6000
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3840
      Top             =   5880
   End
   Begin VB.Label lblNum 
      Alignment       =   2  'Center
      DragMode        =   1  'Automatic
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
      Index           =   0
      Left            =   3960
      TabIndex        =   1
      Top             =   4440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblGrid 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   0
      Left            =   3240
      TabIndex        =   0
      Top             =   4440
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "frmSud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim row(1 To 9, 1 To 9) As Boolean
Dim col(1 To 9, 1 To 9) As Boolean
Dim box(1 To 9, 1 To 9) As Boolean

Dim solvedSud(1 To 9, 1 To 9) As Integer

Dim z As Integer, acount As Integer
Dim er, ec As Integer

Dim found As Boolean

Public diff As Integer '1 to 4; 1 is easiest


Private Sub cmdSolve_Click()
    For i = 1 To 9
        For j = 1 To 9
            k = k + 1
            lblGrid(k).Caption = solvedSud(i, j)
            lblGrid(k).Enabled = False
            'lblGrid(k).ForeColor = vbBlack
        Next j
    Next i
End Sub

Private Sub Form_Load()
    Randomize '(100)
    Call createboard
    'frmMon.Show
    Timer1.Enabled = True
    cmdSolve.Top = frmSud.Height - cmdSolve.Height * 2 - 200
    cmdSolve.Left = frmSud.Width / 2 - cmdSolve.Width / 2
    
End Sub

Sub ATd(ByVal txt As String)
    frmMon.txtDebug.Text = frmMon.txtDebug.Text & vbCrLf & txt
    frmMon.txtDebug.SelStart = Len(frmMon.txtDebug) - 1
End Sub

Sub resetAllArray()
    k = 1
    For i = 1 To 9
        For j = 1 To 9
            row(i, j) = False
            col(i, j) = False
            box(i, j) = False
            lblGrid(k).Caption = ""
            k = k + 1
        Next j
    Next i
End Sub

Private Sub fillAll()
    ATd ("Try number: " & acount)
    
    Dim x As Integer, y As Integer, qq As Integer
    x = 1
    y = 1
    z = 1
    For i = 1 To 9
        For j = 1 To 9
            
            q = Int(Rnd * 9 + 1)
            Do While row(i, q) = True Or col(j, q) = True Or box(x, q) = True Or z = 82
                q = Int(Rnd * 9 + 1)
                qq = qq + 1
                If qq > 1000 Then
                    'MsgBox ("starting over")
                   'found = False
                   Exit Sub
                End If
            Loop

            row(i, q) = True
            col(j, q) = True
            box(x, q) = True
            solvedSud(i, j) = q
            If j Mod 3 = 0 Then
                x = x + 1
            End If
            'lblGrid(z) = q
            z = z + 1
            If z = 82 Then
                found = True
                Exit Sub
            End If
        Next j
        If i Mod 3 = 0 Then
            y = y + 3
            x = y
        Else
            x = y
        End If
    Next i
            
End Sub

Sub printSud()
    For i = 1 To 9
        For j = 1 To 9
            k = k + 1
            lblGrid(k) = solvedSud(i, j)
        Next j
    Next i
End Sub




Private Sub lblGrid_DblClick(Index As Integer)
    lblGrid(Index).Caption = ""
End Sub

Private Sub lblGrid_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)
    lblGrid(Index).Caption = Source.Caption
    'solvedSud(Index, Index) = Source.Caption
    Call checkboard
End Sub

Sub resetMainArray()
    For i = 1 To 9
        For j = 1 To 9
            solvedSud(i, j) = 0
        Next j
    Next i
        
End Sub


Sub checkboard()
    For i = 1 To 81
        If lblGrid(i).Caption = "" Then
            Exit Sub
        End If
    Next i
    
    Call doublecheck
End Sub


Sub doublecheck()
    k = 1
    For i = 1 To 9
        For j = 1 To 9
            If lblGrid(k).Caption <> solvedSud(i, j) Then
                frmStatus.Caption = "Try again!"
                frmStatus.lblStatus = "Try again!"
                frmStatus.Show
                Exit Sub
            End If
            k = k + 1
        Next j
    Next i
    
    frmStatus.Caption = "You win!"
    frmStatus.lblStatus = "You win!!"
    frmStatus.Show
    For i = 1 To 81
        lblGrid(i).Enabled = False
    Next i
End Sub




'Private Sub lblGrid_Click(Index As Integer)
'
'    Do While found = False
'        Call fillAll
'        If found = False Then
'            Call resetAllArray
'            Call resetMainArray
'        End If
'
'
'        acount = acount + 1'
'
'    Loop
'
'    Call printSud
'    frmSplash.Hide
'    'Call colorNumbers
'
'End Sub

Sub colorNumbers()
    For i = 1 To 81
        Select Case lblGrid(i).Caption
            Case 1
                lblGrid(i).ForeColor = vbGreen
            Case 2
                lblGrid(i).ForeColor = vbBlue
            Case 3
                lblGrid(i).ForeColor = vbCyan
            Case 4
                lblGrid(i).ForeColor = vbMagenta
            Case 5
                lblGrid(i).ForeColor = vbYellow
            Case 6
                lblGrid(i).ForeColor = vbRed
            Case 7
                lblGrid(i).ForeColor = &H80FF& 'orange
            Case 8
                lblGrid(i).ForeColor = &H8080FF 'salmon
            Case 9
                lblGrid(i).ForeColor = &HFF8080 'light blue
        End Select
    Next i
End Sub

Sub findEmptyRows(ByVal rn As Integer)
    Call checkRowArray(rn)
    er = 0
    For i = 1 To 9
        If row(i) = False Then
            er = er + 1
        End If
    Next i
End Sub


Sub createboard()
    l = 100
    t = 100
    For i = 1 To 81
        Load lblGrid(i)
        lblGrid(i).Left = l
        lblGrid(i).Top = t
        l = l + lblGrid(0).Width
        lblGrid(i).Visible = True
        If i Mod 3 = 0 Then
            l = l + 50
        End If
        
        If i Mod 9 = 0 Then
            t = t + lblGrid(0).Height
            l = 100
        End If
        
        If i Mod 27 = 0 Then
            t = t + 50
        End If
    Next i
    
    l = 100
    t = lblGrid(81).Top + lblGrid(0).Height + 250
    
    For i = 1 To 9
        Load lblNum(i)
        lblNum(i).Visible = True
        lblNum(i).Left = l
        lblNum(i).Top = t
        l = l + lblNum(i).Width
        lblNum(i).Caption = i
    Next i
    
    Me.Width = lblGrid(9).Left + lblGrid(0).Width + 200
    Me.Height = lblNum(1).Top + lblNum(1).Height * 2 + cmdSolve.Height

End Sub

Private Sub Timer1_Timer()

    Do While found = False
        Call fillAll
        If found = False Then
            Call resetAllArray
            Call resetMainArray
        End If
        

        acount = acount + 1

    Loop
    
    'Call printSud
    Call printHints
    'Call printSud
    frmSplash.Hide
    Timer1.Enabled = False
    
End Sub

Sub printHints()
    Select Case diff
        Case 1
            'give them 60 to start with
            k = 1
            hg = 0
            For i = 1 To 9
                For j = 1 To 9
                    If Int(Rnd * 100 + 1) > 36 Then 'give it to them
                        lblGrid(k).Caption = solvedSud(i, j)
                        lblGrid(k).Enabled = False
                        hg = hg + 1
                    End If
                    
                    If hg >= 60 Then
                        Exit Sub
                    End If
                    k = k + 1
                Next j
            Next i
            
        Case 2
            'give them 40 to start with
            k = 1
            hg = 0
            For i = 1 To 9
                For j = 1 To 9
                    If Int(Rnd * 100 + 1) > 50 Then 'give it to them
                        lblGrid(k).Caption = solvedSud(i, j)
                        lblGrid(k).Enabled = False
                        hg = hg + 1
                    End If
                    
                    If hg >= 40 Then
                        Exit Sub
                    End If
                    k = k + 1
                Next j
            Next i
            
            
        Case 3
            'give them 35 to start with
            k = 1
            hg = 0
            For i = 1 To 9
                For j = 1 To 9
                    If Int(Rnd * 100 + 1) > 50 Then 'give it to them
                        lblGrid(k).Caption = solvedSud(i, j)
                        lblGrid(k).Enabled = False
                        hg = hg + 1
                    End If
                    
                    If hg >= 35 Then
                        Exit Sub
                    End If
                    k = k + 1
                Next j
            Next i
            
        Case 4
            'give them 27 to start with
            k = 1
            hg = 0
            For i = 1 To 9
                For j = 1 To 9
                    If Int(Rnd * 100 + 1) > 60 Then 'give it to them
                        lblGrid(k).Caption = solvedSud(i, j)
                        lblGrid(k).Enabled = False
                        hg = hg + 1
                    End If
                    
                    If hg >= 27 Then
                        Exit Sub
                    End If
                    k = k + 1
                Next j
            Next i
    End Select


End Sub
