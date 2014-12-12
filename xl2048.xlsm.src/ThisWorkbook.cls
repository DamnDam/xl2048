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

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
'LocalStorage.Save

Application.ScreenUpdating = False
GameApp.Clear
'Board.Unprotect
'Range("BestScore") = ""
'Board.Protect
End Sub

Private Sub Workbook_AfterSave(ByVal Success As Boolean)
'LocalStorage.Load
End Sub

Private Sub Workbook_Open()
Application.ScreenUpdating = False
'Board.Unprotect
Board.ScrollArea = "A1:A1"
Board.Activate
Application.WindowState = xlMaximized
Board.Range("A1:H8").Select
ActiveWindow.Zoom = True
Board.Range("XL2048").Select
'Board.Protect

'Application.OnTime Now, "'LocalStorage.Load Animate:=True'"
Me.Saved = True
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
'LocalStorage.Save
End Sub
