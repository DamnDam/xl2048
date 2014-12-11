VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Grid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''
' Members
'
'
Private Tiles As New Dictionary
Private oldTiles As Dictionary

''
' Public methods
'
'
Public Sub addTile(Cell As tCoordinates, Value As Long)
Dim Tile As New Tile
Tile.Cell = Cell
Tile.Value = Value
Tiles.Add Key(Cell), Tile
End Sub

Public Sub moveTile(Tile As Tile, Cell As tCoordinates)
Tiles.Key(Key(Tile.Cell)) = Key(Cell)
Tile.Cell = Cell
End Sub

Public Function hasTile(Cell As tCoordinates) As Boolean
hasTile = Tiles.Exists(Key(Cell))
End Function

Public Function getTile(Cell As tCoordinates) As Tile
Set getTile = Tiles(Key(Cell))
End Function

Public Sub delTile(Tile As Tile)
Tiles.Remove Key(Tile.Cell)
Set Tile = Nothing
End Sub

Public Function getBoard() As Variant
Dim currBoard(1 To GRID_SIZE, 1 To GRID_SIZE)
Dim Tile As Tile, I As Integer
For I = 0 To Tiles.Count - 1
    Set Tile = Tiles.Items(I)
    currBoard(Tile.Cell.Top, Tile.Cell.Left) = Tile.Value
Next I
getBoard = currBoard
End Function

Public Function getDiff() As String
Dim Tile As Tile, oldTile As Tile, I As Integer

If oldTiles Is Nothing Then
    getDiff = Range(Cells(1, 1), Cells(GRID_SIZE, GRID_SIZE)).Address
Else
    For I = 0 To Tiles.Count - 1
        Set Tile = Tiles.Items(I)
        If oldTiles.Exists(Key(Tile.Cell)) Then
            If oldTiles(Key(Tile.Cell)).Value <> Tile.Value Then
                getDiff = getDiff & "," & Cells(Tile.Cell.Top, Tile.Cell.Left).Address
            End If
        Else
            getDiff = getDiff & "," & Cells(Tile.Cell.Top, Tile.Cell.Left).Address
        End If
    Next I
    For I = 0 To oldTiles.Count - 1
        Set Tile = oldTiles.Items(I)
        If Not Tiles.Exists(Key(Tile.Cell)) Then
            getDiff = getDiff & "," & Cells(Tile.Cell.Top, Tile.Cell.Left).Address
        End If
    Next I
    If Len(getDiff) > 0 Then getDiff = Mid(getDiff, 2)
End If

cleanTiles oldTiles
Set oldTiles = New Dictionary
Dim tileKey As Variant
For I = 0 To Tiles.Count - 1
    tileKey = Tiles.Keys(I)
    Set Tile = Tiles(tileKey)
    Set oldTile = New Tile
    oldTile.Cell = Tile.Cell
    oldTile.Value = Tile.Value
    oldTiles.Add tileKey, oldTile
Next I
End Function

Public Function getRandomEmptyCell() As tCoordinates
Dim RndCell As Integer, Cells As Collection, NoTile As Tile
Set Cells = EmptyCells()
RndCell = Int(Rnd * Cells.Count) + 1
Set NoTile = Cells(RndCell)
getRandomEmptyCell = NoTile.Cell
For Each NoTile In Cells
    Set NoTile = Nothing
Next NoTile
Set Cells = Nothing
End Function

''
' Private methods
'
'
Private Function EmptyCells() As Collection
Dim Cell As tCoordinates, Tile As Tile
Set EmptyCells = New Collection
With Cell
    For .Top = 1 To GRID_SIZE
        For .Left = 1 To GRID_SIZE
            If Not hasTile(Cell) Then
                Set Tile = New Tile
                Tile.Cell = Cell
                EmptyCells.Add Tile
            End If
        Next .Left
    Next .Top
End With
End Function

Private Function Key(Coordinates As tCoordinates) As String
Key = Coordinates.Top & "." & Coordinates.Left
End Function

Private Sub cleanTiles(DelTiles As Dictionary)
Dim I As Integer
If DelTiles Is Nothing Then Exit Sub
For I = 1 To DelTiles.Count
    Set DelTiles.Item(I) = Nothing
Next I
DelTiles.RemoveAll
Set DelTiles = Nothing
End Sub

''
' Events
'
'
Private Sub Class_Terminate()
On Error Resume Next
cleanTiles Tiles
cleanTiles oldTiles
End Sub
