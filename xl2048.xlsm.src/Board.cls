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

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub updateScreen(Optional Delay As Integer)
Activate
Application.ScreenUpdating = True
Application.EnableEvents = False
DoEvents
Application.EnableEvents = True
Application.ScreenUpdating = False
Sleep Delay
Unprotect
Range("XL2048").Select
Protect
End Sub

Public Sub updateBoard(GameState As tGameState)
Application.ScreenUpdating = False
Unprotect
Range("Score") = GameState.Score
Range("BestScore") = GameState.BestScore
Range("Playground").Value = GameState.Grid.getBoard

Dim Diff
Diff = GameState.Grid.getDiff
If Len(Diff) > 0 Then Style.Apply Range("Playground").Range(Diff)

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

Application.ScreenUpdating = False
Shapes("lblGameOver").Visible = Value

If Value Then
    Shapes("lblGameOver").TextFrame2.TextRange.Font.Fill.Transparency = 1
    Shapes("lblGameOver").Fill.Transparency = 1
    
    For I = 100 To 0 Step -10
        updateScreen 0
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
    
    For I = 100 To 0 Step -10
        updateScreen 0
        Shapes("lblGameWon").Fill.Transparency = 0.33 + (2 * I / 300)
        Shapes("lblGameWon").TextFrame2.TextRange.Font.Fill.Transparency = I / 100
    Next I
End If

Shapes("btnContinue").Visible = Value
End Property
