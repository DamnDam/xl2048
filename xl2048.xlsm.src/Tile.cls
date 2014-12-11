VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
 
Private mCell As tCoordinates
Public Value As Long

Public Property Get Cell() As tCoordinates
Cell = mCell
End Property

Public Property Let Cell(Cell As tCoordinates)
mCell = Cell
End Property

