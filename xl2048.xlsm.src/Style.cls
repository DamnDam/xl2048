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

Dim StyleCache As Dictionary

Private Property Get getTemplate(Value As Variant) As Range
Dim Template As Range
If StyleCache Is Nothing Then
    Set StyleCache = New Dictionary
    For Each Template In UsedRange.Cells
        If Not IsEmpty(Template) Then
            StyleCache.Add Template.Value, Template
        End If
    Next Template
End If
If Not StyleCache.Exists(Value) Then
    Set getTemplate = StyleCache("[0]")
Else
    Set getTemplate = StyleCache(Value)
End If
End Property

Public Sub Apply(Range As Range)
Dim Cell As Range, Template As Range

For Each Cell In Range.Cells
    Set Template = getTemplate(Cell.Value)
    
    Cell.Interior.Color = Template.Interior.Color
    Cell.Font.Color = Template.Font.Color
    Cell.Font.Size = Template.Font.Size
nextCell:
Next Cell
End Sub
