VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IControlProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Manager As IGameManager

Public Sub Register(GameManager As IGameManager)
Set GameManager.Control = Me
Set Manager = GameManager
End Sub

Public Property Get Enable() As Boolean

End Property

Public Property Let Enable(Enable As Boolean)

End Property
