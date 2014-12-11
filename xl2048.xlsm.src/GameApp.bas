Attribute VB_Name = "GameApp"
Option Explicit

Public Const GRID_SIZE = 4

Public Type tGameState
    Grid As Grid
    Score As Long
    BestScore As Long
    GameOver As Boolean
    GameWon As Boolean
End Type

Public Enum tDirection
    toUp = xlUp
    toDown = xlDown
    toleft = xlToLeft
    toRight = xlToRight
End Enum

Public Type tCoordinates
    Top As Integer
    Left As Integer
End Type

Dim mManager As IGameManager

Public Property Get Manager() As IGameManager
If mManager Is Nothing Then
    Set mManager = New GameManager
End If
Set Manager = mManager
End Property

Sub newGame()
Manager.newGame
End Sub

Sub Continue()
Manager.Continue
End Sub

Sub doMove(Direction As tDirection)
Manager.doMove Direction
End Sub

Sub Clear()
Manager.Clear
End Sub
