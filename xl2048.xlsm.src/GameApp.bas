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
    toLeft = xlToLeft
    toRight = xlToRight
End Enum

Public Type tCoordinates
    Top As Integer
    Left As Integer
End Type

Dim mManager As IGameManager
Dim mKBController As KeyboardControl

Public Property Get Manager() As IGameManager
Dim Control As IControlProvider
If mManager Is Nothing Then
    Set mManager = New GameManager
    Set Control = KBController
    Control.Register mManager
End If
Set Manager = mManager
End Property

Private Property Get KBController() As KeyboardControl
If mKBController Is Nothing Then
    Set mKBController = New KeyboardControl
End If
Set KBController = mKBController
End Property

Sub newGame()
Manager.newGame
End Sub

Sub Continue()
Manager.Continue
End Sub

Sub Clear()
Manager.Clear
End Sub

Sub KBdoMove(Direction As tDirection)
KBController.Callback_DoMove Direction
End Sub
