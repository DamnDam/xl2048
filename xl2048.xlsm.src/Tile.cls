VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private pCell As Range
Private pFusable As Boolean

Public Function setCell(Cell As Range)
Set pCell = Cell
pFusable = True
End Function

Public Property Get Cell() As Range
Set Cell = pCell
End Property

Public Property Get Fusable() As Boolean
Fusable = pFusable
End Property

Public Property Get Movable() As Boolean
Movable = True
If Not Neighbor(toUp) Is Nothing Then _
    If Neighbor(toUp).Value = pCell Then Exit Property
If Not Neighbor(toDown) Is Nothing Then _
    If Neighbor(toDown).Value = pCell Then Exit Property
If Not Neighbor(toRight) Is Nothing Then _
    If Neighbor(toRight).Value = pCell Then Exit Property
If Not Neighbor(toLeft) Is Nothing Then _
    If Neighbor(toLeft).Value = pCell Then Exit Property
    
Movable = False
End Property

Public Function Move(Direction As tDirection) As Boolean
Do
If Neighbor(Direction) Is Nothing Then Exit Do
If Neighbor(Direction) = "" Then
    Move = True
    Neighbor(Direction) = pCell
    pCell.ClearContents
    Style.Apply pCell
    Game.myTiles.delStack Me
    Set pCell = Neighbor(Direction)
    Style.Apply pCell
    Game.myTiles.addStack Me
Else
    Exit Do
End If
Loop
End Function

Public Function Fuse(Direction As tDirection) As Boolean
Dim canFuse As Boolean
On Error Resume Next
canFuse = Game.myTiles.getStack(Neighbor(Direction).Address).Fusable
canFuse = canFuse And Fusable
If Not canFuse Then Exit Function

On Error GoTo 0
If Neighbor(Direction) = pCell Then
    Fuse = True
    Neighbor(Direction) = Neighbor(Direction) * 2
    pCell.ClearContents
    Style.Apply pCell
    Game.myTiles.delStack Me
    Set pCell = Neighbor(Direction)
    Style.Apply pCell
    pFusable = False
    Game.myTiles.addStack Me
    Game.addScore pCell
End If
End Function

Private Property Get Neighbor(Direction) As Range
Select Case Direction
Case toUp:
Set Neighbor = Intersect(Range("Playground"), pCell.Offset(-1, 0))
Case toDown:
Set Neighbor = Intersect(Range("Playground"), pCell.Offset(1, 0))
Case toRight:
Set Neighbor = Intersect(Range("Playground"), pCell.Offset(0, 1))
Case toLeft:
Set Neighbor = Intersect(Range("Playground"), pCell.Offset(0, -1))
End Select
End Property

Private Sub Class_Terminate()
Set pCell = Nothing
End Sub
