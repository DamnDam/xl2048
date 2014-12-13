VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Board"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Public Sub updateBoard(GameState As tGameState)
Unprotect

Application.ScreenUpdating = False

Range("Score") = GameState.Score
Range("BestScore") = GameState.BestScore
Range("Playground").Value = GameState.Grid.getBoard

Dim Diff
Diff = GameState.Grid.getDiff
If Len(Diff) > 0 Then
    Style.Apply Range("Playground").Range(Diff)
End If

Application.ScreenUpdating = True
forceRefresh

GameOver = GameState.GameOver
GameWon = GameState.GameWon

Protect
End Sub

Public Sub Clear(Optional xlHidden As Object)
Unprotect
Range("Playground").Interior.Color = 14277081
Range("Playground").ClearContents
Range("Score") = ""


GameOver = False
GameWon = False
Protect
End Sub

Private Property Let GameOver(Value As Boolean)
If Shapes("lblGameOver").Visible = Value Then Exit Property
Dim I As Integer

Shapes("lblGameOver").TextFrame2.TextRange.Font.Fill.Transparency = 1
Shapes("lblGameOver").Fill.Transparency = 1
Shapes("lblGameOver").Visible = Value

If Value Then
    For I = 100 To 0 Step -1
        forceRefresh
        Shapes("lblGameOver").Fill.Transparency = 0.33 + (2 * I / 300)
        Shapes("lblGameOver").TextFrame2.TextRange.Font.Fill.Transparency = I / 100
    Next I
End If
    
Shapes("btnAgain").Visible = Value
End Property

Private Property Let GameWon(Value As Boolean)
If Shapes("lblGameWon").Visible = Value Then Exit Property
Dim I As Integer

Application.ScreenUpdating = False
Shapes("lblGameWon").Visible = Value

If Value Then
    Shapes("lblGameWon").TextFrame2.TextRange.Font.Fill.Transparency = 1
    Shapes("lblGameWon").Fill.Transparency = 1
    
    Application.ScreenUpdating = True
    For I = 100 To 0 Step -1
        forceRefresh
        Shapes("lblGameWon").Fill.Transparency = 0.33 + (2 * I / 300)
        Shapes("lblGameWon").TextFrame2.TextRange.Font.Fill.Transparency = I / 100
    Next I
End If

Shapes("btnContinue").Visible = Value
End Property

Private Sub forceRefresh()
Range("Playground").Borders.Weight = 2
End Sub

Private Sub Worksheet_Activate()
Unprotect
Range("XL2048").Select
Protect
End Sub
