VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LocalStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' TODO
#If 0 Then
Option Explicit
Implements IStorageProvider

Private Const emptyGrid = "!....!....!....!...."
Private Const rowSep = "!"
Private Const tileSep = "."

Private Sub IStorageProvider_Save(GameState As tGameState)
Dim strRow As String, strGrid As String, I As Integer, J As Integer
Dim isSaved As Boolean
isSaved = ThisWorkbook.Saved

SaveSetting "xl2048", "2048", "Score", CStr(Board.Range("Score"))
SaveSetting "xl2048", "2048", "Best", CStr(Board.Range("BestScore"))

If Game.GameOver Then
    Game.Clear
    ThisWorkbook.Saved = isSaved
    SaveSetting "xl2048", "2048", "Grid", ""
    Exit Sub
End If

For I = 1 To 4
    strGrid = strGrid & rowSep
    For J = 1 To 4
        strGrid = strGrid & tileSep & CStr(Board.Range("Playground").Cells(I, J))
    Next J
Next I

SaveSetting "xl2048", "2048", "Grid", strGrid

End Sub

Private Function IStorageProvider_Load() As tGameState
Dim Row, Grid, I As Integer, J As Integer
Dim isSaved As Boolean
isSaved = ThisWorkbook.Saved

Application.ScreenUpdating = False
Game.Clear
Board.Unprotect

Board.Range("Score") = GetSetting("xl2048", "2048", "Score")
Board.Range("BestScore") = GetSetting("xl2048", "2048", "Best")
Grid = GetSetting("xl2048", "2048", "Grid")

If Animate Then Board.updateScreen 500

If Grid = "" Or Grid = emptyGrid Then
    Board.Protect
    ThisWorkbook.Saved = isSaved
    Application.OnTime Now, "Game.newGame"
    Exit Function
End If

Grid = Split(Grid, rowSep)
For I = 1 To 4
    Row = Split(Grid(I), tileSep)
    For J = 1 To 4
        If Row(J) <> "" Then
            If Animate Then Board.updateScreen 75
            Board.Range("Playground").Cells(I, J) = CLng(Row(J))
            Style.Apply Board.Range("Playground").Cells(I, J)
            If CLng(Row(J)) >= 2048 Then
                Application.OnTime Now, "Game.Continue"
            End If
        End If
    Next J
Next I

If Animate Then Board.updateScreen

Board.Protect
ThisWorkbook.Saved = isSaved
End Function

#End If
