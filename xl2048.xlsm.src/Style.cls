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
Dim TCell As Range
Set TCell = Me.UsedRange.Find(Range.Value, LookIn:=xlValues, LookAt:=xlWhole)

If IsEmpty(Range) Or TCell Is Nothing Then
    Range.Interior.Color = 14277081
    Exit Sub
End If

Range.Interior.Color = TCell.Interior.Color
Range.Font.Color = TCell.Font.Color
Range.Font.Size = TCell.Font.Size
End Sub
