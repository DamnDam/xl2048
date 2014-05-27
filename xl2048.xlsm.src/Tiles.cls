VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim CStack As New Collection

Public Function getStack(Key As String)
Set getStack = CStack(Key)
End Function

Public Sub delStack(aTile As Tile)
On Error Resume Next
CStack.Remove aTile.Cell.Address
End Sub

Public Sub addStack(aTile As Tile)
On Error Resume Next
CStack.Add aTile, aTile.Cell.Address

If Err = 0 Then Exit Sub

On Error GoTo 0
CStack.Remove aTile.Cell.Address
CStack.Add aTile, aTile.Cell.Address
End Sub

Public Function Go(Direction As tDirection) As Boolean
Dim aTile As Tile
For Each aTile In CStack
    Go = Go Or aTile.Move(Direction)
    Go = Go Or aTile.Fuse(Direction)
Next aTile
End Function

Public Property Get Movable() As Boolean
Dim aTile As Tile
Movable = True
For Each aTile In CStack
    If aTile.Movable Then Exit Function
Next aTile
Movable = False
End Property

Private Sub Class_Terminate()
Dim Dummy
For Each Dummy In CStack
    Set Dummy = Nothing
Next Dummy
End Sub
