VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

''
' Events
'
'
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
GameApp.Save

GameApp.Clear
End Sub

Private Sub Workbook_AfterSave(ByVal Success As Boolean)
GameApp.Load
End Sub

Private Sub Workbook_Open()
Application.ScreenUpdating = False
Board.ScrollArea = "A1:A1"
Board.Activate
Application.WindowState = xlMaximized
Board.Range("A1:H8").Select
ActiveWindow.Zoom = True
Board.Range("XL2048").Select
Me.Saved = True

GameApp.Load
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
GameApp.Save
End Sub
