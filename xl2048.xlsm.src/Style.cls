VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Style"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Public Sub Apply(Range As Range)
Dim Template As Range, Cell As Range
For Each Cell In Range.Cells
    Set Template = Me.UsedRange.Find(Cell.Value, LookIn:=xlValues, LookAt:=xlWhole)
    
    If IsEmpty(Cell) Or Template Is Nothing Then
        Cell.Interior.Color = 14277081
        GoTo nextCell
    End If
    
    Cell.Interior.Color = Template.Interior.Color
    Cell.Font.Color = Template.Font.Color
    Cell.Font.Size = Template.Font.Size
nextCell:
Next Cell
End Sub
