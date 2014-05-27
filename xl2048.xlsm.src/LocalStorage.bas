Attribute VB_Name = "LocalStorage"
Option Explicit

Private Const emptyGrid = "!....!....!....!...."
Private Const rowSep = "!"
Private Const tileSep = "."

Sub Save(Optional xlHidden As Object)
Dim strRow As String, strGrid As String, I As Integer, J As Integer
Dim isSaved As Boolean
isSaved = ThisWorkbook.Saved

SaveSetting "xl2048", "2048", "Score", CStr(Game.Range("Score"))
SaveSetting "xl2048", "2048", "Best", CStr(Game.Range("BestScore"))

If Game.GameOver Then
    Game.Clear
    ThisWorkbook.Saved = isSaved
    SaveSetting "xl2048", "2048", "Grid", ""
    Exit Sub
End If

For I = 1 To 4
    strGrid = strGrid & rowSep
    For J = 1 To 4
        strGrid = strGrid & tileSep & CStr(Game.Range("Playground").Cells(I, J))
    Next J
Next I

SaveSetting "xl2048", "2048", "Grid", strGrid

End Sub

Sub Load(Optional Animate As Boolean)
Dim Row, Grid, I As Integer, J As Integer
Dim isSaved As Boolean
isSaved = ThisWorkbook.Saved

Application.ScreenUpdating = False
Game.Clear
Game.Unprotect

Game.Range("Score") = GetSetting("xl2048", "2048", "Score")
Game.Range("BestScore") = GetSetting("xl2048", "2048", "Best")
Grid = GetSetting("xl2048", "2048", "Grid")

If Animate Then ThisWorkbook.screenActuate 500

If Grid = "" Or Grid = emptyGrid Then
    Game.Protect
    ThisWorkbook.Saved = isSaved
    Application.OnTime Now, "Game.newGame"
    Exit Sub
End If

Grid = Split(Grid, rowSep)
For I = 1 To 4
    Row = Split(Grid(I), tileSep)
    For J = 1 To 4
        If Row(J) <> "" Then
            If Animate Then ThisWorkbook.screenActuate 75
            Game.Range("Playground").Cells(I, J) = CLng(Row(J))
            Style.Apply Game.Range("Playground").Cells(I, J)
            If CLng(Row(J)) >= 2048 Then
                Application.OnTime Now, "Game.Continue"
            End If
        End If
    Next J
Next I

If Animate Then ThisWorkbook.screenActuate

Game.Protect
ThisWorkbook.Saved = isSaved
End Sub
