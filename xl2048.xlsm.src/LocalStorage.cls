VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LocalStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IStorageProvider

''
' Private constants
'
'
Private Const emptyGrid = "!....!....!....!...."
Private Const rowSep = "!"
Private Const tileSep = "."

''
' Interface
'
'
Private Sub IStorageProvider_Save(GameState As tGameState)
Dim Grid, strGrid As String, I As Integer, J As Integer

SaveSetting "xl2048", "2048", "Score", CStr(GameState.Score)
SaveSetting "xl2048", "2048", "Best", CStr(GameState.BestScore)

If GameState.GameOver Or GameState.Grid Is Nothing Then
    GameApp.Clear
    SaveSetting "xl2048", "2048", "Grid", ""
    Exit Sub
End If

Grid = GameState.Grid.getBoard
For I = 1 To GRID_SIZE
    strGrid = strGrid & rowSep
    For J = 1 To GRID_SIZE
        strGrid = strGrid & tileSep & Grid(I, J)
    Next J
Next I

SaveSetting "xl2048", "2048", "Grid", strGrid
End Sub

Private Function IStorageProvider_Load() As tGameState
Dim Row, Grid, Cell As tCoordinates

Dim GameState As tGameState

Set GameState.Grid = Nothing
Set GameState.Grid = New Grid

GameState.Score = GetSetting("xl2048", "2048", "Score")
GameState.BestScore = GetSetting("xl2048", "2048", "Best")
Grid = GetSetting("xl2048", "2048", "Grid")

If Grid = "" Or Grid = emptyGrid Then
    Application.OnTime Now, "GameApp.newGame"
    IStorageProvider_Load = GameState
    Exit Function
End If

Grid = Split(Grid, rowSep)
For Cell.Top = 1 To 4
    Row = Split(Grid(Cell.Top), tileSep)
    For Cell.Left = 1 To 4
        If Row(Cell.Left) <> "" Then
            GameState.Grid.addTile Cell, CLng(Row(Cell.Left))
            If CLng(Row(Cell.Left)) >= 2048 Then
                GameState.Continue = True
            End If
        End If
    Next Cell.Left
Next Cell.Top

IStorageProvider_Load = GameState
End Function
