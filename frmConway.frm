VERSION 5.00
Begin VB.Form frmConway 
   Caption         =   "Conway"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer timerRun 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   840
      Top             =   720
   End
   Begin VB.CheckBox chkCell 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   0
      MaskColor       =   &H0000C000&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   375
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuGameRun 
         Caption         =   "&Run"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuGameQuit 
         Caption         =   "&Quit"
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "frmConway"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const cells = 20
Dim futureState(0 To cells * cells) As Boolean

Private Sub Form_Load()

    Dim idx As Integer
    
    For idx = 1 To cells * cells
        Load chkCell(idx)
    Next idx
    
    Dim marginLeft As Integer
    Dim marginTop As Integer
    Dim iColWidth As Integer
    Dim iRowHeight As Integer
    
    Dim lineWidth As Integer
    lineWidth = 0
    
    marginLeft = 0
    marginTop = 0
    iColWidth = chkCell(0).Width
    iRowHeight = chkCell(0).Height
    
    frmConway.Width = iColWidth * cells + (frmConway.Width - frmConway.ScaleWidth)
    frmConway.Height = iRowHeight * cells + (frmConway.Height - frmConway.ScaleHeight)
    
    Dim iRow As Integer
    Dim iCol As Integer
    
    For iRow = 0 To cells - 1
        For iCol = 0 To cells - 1
            idx = iCol + cells * iRow
            chkCell(idx).Left = marginLeft + iCol * iColWidth + lineWidth
            chkCell(idx).Top = marginTop + iRow * iRowHeight + lineWidth
            chkCell(idx).Visible = True
            
            futureState(idx) = toBool(chkCell(idx).Value)
        Next iCol
    Next iRow
End Sub

Private Function toBool(v As Integer) As Boolean
    If v = 1 Then
        toBool = True
    Else
        toBool = False
    End If
End Function

Private Sub mnuGameQuit_Click()
    Unload frmConway
    Set frmConway = Nothing
End Sub

Private Sub mnuGameRun_Click()
    mnuGameRun.Checked = Not mnuGameRun.Checked
    timerRun.Enabled = mnuGameRun.Checked
End Sub

Private Function liveNeighbourCount(row As Integer, col As Integer) As Integer
    Dim count As Integer
    count = 0
    Dim i As Integer, j As Integer
    Dim inspectCellX As Integer, inspectCellY As Integer
    Dim inspectIdx As Integer
    
    For i = -1 To 1
        For j = -1 To 1
            If (Not i = 0) Or (Not j = 0) Then
                
                inspectCellX = (col + i + cells) Mod cells
                inspectCellY = (row + j + cells) Mod cells
                
                inspectIdx = inspectCellX + cells * inspectCellY
                
                count = count + chkCell(inspectIdx).Value
            End If
        Next j
    Next i
    
    liveNeighbourCount = count
End Function

Private Sub timerRun_Timer()
    ' Do a game step!
    Dim idx As Integer
    Dim iRow As Integer
    Dim iCol As Integer
    Dim liveNeighbours As Integer
    
    For iRow = 0 To cells - 1
        For iCol = 0 To cells - 1
            
            idx = iCol + cells * iRow
            liveNeighbours = liveNeighbourCount(iRow, iCol)
            
            ' copy state:
            futureState(idx) = toBool(chkCell(idx).Value)

            If chkCell(idx).Value = 0 Then
                ' it's currently dead
                If liveNeighbours >= 3 Then
                    futureState(idx) = True
                End If
            Else
                ' it's currently alive
                If liveNeighbours = 2 Or liveNeighbours = 3 Then
                    ' it stays alive
                Else
                    ' it dies
                    futureState(idx) = False
                End If
            End If
        Next iCol
    Next iRow
    
    Call applyState
End Sub

Private Function fromBool(b As Boolean) As Integer
    If b Then
        fromBool = 1
    Else
        fromBool = 0
    End If
End Function

Private Sub applyState()
    Dim iRow As Integer, iCol As Integer
    Dim idx As Integer
    
    For iRow = 0 To cells - 1
        For iCol = 0 To cells - 1
            idx = iCol + cells * iRow
            chkCell(idx).Value = fromBool(futureState(idx))
        Next iCol
    Next iRow
End Sub
