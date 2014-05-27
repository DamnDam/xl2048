VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Game"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Public myTiles As Tiles
Dim mContinue As Boolean

Public Sub newGame()
Dim isSaved As Boolean
isSaved = ThisWorkbook.Saved

Application.ScreenUpdating = False

Clear

PopNew
PopNew

ThisWorkbook.Saved = isSaved
End Sub

Public Sub Continue()
Dim isSaved As Boolean
isSaved = ThisWorkbook.Saved
GameWon = False
mContinue = True
ThisWorkbook.Saved = isSaved
End Sub

Public Sub doMove(Direction As tDirection)
Dim aTile As Tile
Dim firstAdd As String, findCell As Range
Dim lOrder As Integer, lStart As Integer, lEnd As Integer
Dim I As Integer

DoEvents
If GameOver Then Exit Sub
If GameWon Then Exit Sub

Dim isSaved As Boolean
isSaved = ThisWorkbook.Saved

If Direction = toUp Or Direction = toLeft Then
    lOrder = 1
    lStart = 1
    lEnd = 4
Else
    lOrder = -1
    lStart = 4
    lEnd = 1
End If

Set myTiles = New Tiles

For I = lStart To lEnd Step lOrder
    If Direction = toDown Or Direction = toUp Then
        Set findCell = Range("Playground").Rows(I).Find("*", LookIn:=xlValues)
        If findCell Is Nothing Then GoTo nextI
        firstAdd = findCell.Address
        Do
            Set aTile = New Tile
            aTile.setCell findCell
            myTiles.addStack aTile
            Set findCell = Range("Playground").Rows(I).Find("*", After:=findCell, LookIn:=xlValues)
        Loop Until findCell.Address = firstAdd
    Else
        Set findCell = Range("Playground").Columns(I).Find("*", LookIn:=xlValues)
        If findCell Is Nothing Then GoTo nextI
        firstAdd = findCell.Address
        Do
            Set aTile = New Tile
            aTile.setCell findCell
            myTiles.addStack aTile
            Set findCell = Range("Playground").Columns(I).Find("*", After:=findCell, LookIn:=xlValues)
        Loop Until findCell.Address = firstAdd
    End If
nextI:
Next I

Me.Unprotect
Application.ScreenUpdating = False

Dim hasMoved As Boolean, ValidMove As Boolean
Do
    hasMoved = myTiles.Go(Direction)
    ValidMove = ValidMove Or hasMoved
Loop While hasMoved

Me.Protect

If ValidMove Then
    PopNew
Else
    GameOver = checkGameOver()
End If

Set myTiles = Nothing
ThisWorkbook.Saved = isSaved
End Sub

Public Sub Clear()
Me.Unprotect
Range("Playground").Interior.Color = 14277081
Range("Playground").ClearContents
Range("Score") = ""
Me.Protect

mContinue = False
GameOver = False
GameWon = False
End Sub

Private Sub PopNew()
Dim popID As Integer, aleaCell As Range
Dim I As Integer

Me.Unprotect
popID = Int(Rnd * Range("Playground").SpecialCells(xlCellTypeBlanks).Cells.Count) + 1
Set aleaCell = Range("Playground").Cells(1, 1)

For I = 1 To popID
    Set aleaCell = Range("Playground").Find("", After:=aleaCell, LookIn:=xlValues)
Next I

ThisWorkbook.screenActuate 150

If Rnd > 0.9 Then
    aleaCell = 4
Else
    aleaCell = 2
End If
Style.Apply aleaCell
Me.Protect
End Sub

Public Sub addScore(ByVal Points As Long)
Game.Range("Score") = Game.Range("Score") + Points
Game.Range("BestScore") = WorksheetFunction.Max(Game.Range("Score"), Game.Range("BestScore"))
If Points = 2048 Then Application.OnTime Now, "Game.Win"
End Sub

Public Sub Win(Optional xlHidden As Object)
If mContinue Then Exit Sub
Dim isSaved As Boolean
isSaved = ThisWorkbook.Saved
GameWon = True
ThisWorkbook.Saved = isSaved
End Sub

Public Property Get GameOver() As Boolean
GameOver = Me.Shapes.Range(Array("GameOver")).Visible
End Property

Private Property Let GameOver(Value As Boolean)
Dim I As Integer

Me.Unprotect
Application.ScreenUpdating = False
Me.Shapes.Range(Array("GameOver")).Visible = Value

If Value Then
    Me.Shapes.Range(Array("GameOver")).TextFrame2.TextRange.Font.Fill.Transparency = 1
    Me.Shapes.Range(Array("GameOver")).Fill.Transparency = 1
    ThisWorkbook.screenActuate 500
    
    For I = 100 To 0 Step -10
        ThisWorkbook.screenActuate 10
        Me.Shapes.Range(Array("GameOver")).Fill.Transparency = 0.33 + (2 * I / 300)
        Me.Shapes.Range(Array("GameOver")).TextFrame2.TextRange.Font.Fill.Transparency = I / 100
        If Not Range("Playground").Find("", LookIn:=xlValues) Is Nothing Then _
            Exit Property
    Next I
End If

Me.Shapes.Range(Array("Again")).Visible = Value

Me.Protect
End Property

Public Property Get GameWon() As Boolean
GameWon = Me.Shapes.Range(Array("GameWon")).Visible
End Property

Private Property Let GameWon(Value As Boolean)
Dim I As Integer

Me.Unprotect
Application.ScreenUpdating = False
Me.Shapes.Range(Array("GameWon")).Visible = Value

If Value Then
    Me.Shapes.Range(Array("GameWon")).TextFrame2.TextRange.Font.Fill.Transparency = 1
    Me.Shapes.Range(Array("GameWon")).Fill.Transparency = 1
    ThisWorkbook.screenActuate 500
    
    For I = 100 To 0 Step -10
        ThisWorkbook.screenActuate 10
        Me.Shapes.Range(Array("GameWon")).Fill.Transparency = 0.33 + (2 * I / 300)
        Me.Shapes.Range(Array("GameWon")).TextFrame2.TextRange.Font.Fill.Transparency = I / 100
        If Range("Playground").Find(2048, LookIn:=xlValues) Is Nothing Then _
            Exit Property
    Next I
End If

Me.Shapes.Range(Array("Continue")).Visible = Value

Me.Protect
End Property

Private Function checkGameOver() As Boolean
Dim emptyCell As Range

Set emptyCell = Range("Playground").Find("", LookIn:=xlValues)
If Not emptyCell Is Nothing Then Exit Function

checkGameOver = Not myTiles.Movable
End Function
