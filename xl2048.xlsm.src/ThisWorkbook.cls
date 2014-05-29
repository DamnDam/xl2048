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

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Enum tDirection
    toUp = xlUp
    toDown = xlDown
    toLeft = xlToLeft
    toRight = xlToRight
End Enum

Public Sub screenActuate(Optional Delay As Integer)
Game.Activate
Application.EnableEvents = False
Application.ScreenUpdating = True
DoEvents
Application.ScreenUpdating = False
Application.EnableEvents = True
Sleep Delay
Game.Range("XL2048").Select
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
LocalStorage.Save

Application.ScreenUpdating = False
Game.Clear
Game.Unprotect
Range("BestScore") = ""
Game.Protect
End Sub

Private Sub Workbook_AfterSave(ByVal Success As Boolean)
LocalStorage.Load
End Sub

Private Sub Workbook_Activate()
Application.OnKey "{UP}", "'Game.doMove " & toUp & "'"
Application.OnKey "{DOWN}", "'Game.doMove " & toDown & "'"
Application.OnKey "{LEFT}", "'Game.doMove " & toLeft & "'"
Application.OnKey "{RIGHT}", "'Game.doMove " & toRight & "'"
End Sub

Private Sub Workbook_Deactivate()
Application.OnKey "{UP}"
Application.OnKey "{DOWN}"
Application.OnKey "{LEFT}"
Application.OnKey "{RIGHT}"
End Sub

Private Sub Workbook_Open()
Application.ScreenUpdating = False
Game.Unprotect
Game.ScrollArea = "A1:A1"
Game.Activate
Application.WindowState = xlMaximized
Game.Range("A1:H8").Select
ActiveWindow.Zoom = True
Game.Range("XL2048").Select
Game.Protect

Application.OnTime Now, "'LocalStorage.Load Animate:=True'"
Me.Saved = True
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
LocalStorage.Save
End Sub
