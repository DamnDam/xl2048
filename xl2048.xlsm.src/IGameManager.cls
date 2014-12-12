VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IGameManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Control As IControlProvider
Attribute Control.VB_VarHelpID = -1

Public Sub newGame()

End Sub

Public Sub Continue()

End Sub

Public Sub doMove(Direction As tDirection)

End Sub

Public Sub Load(GameState As tGameState)

End Sub

Public Function Save() As tGameState

End Function

Public Sub Clear()

End Sub

