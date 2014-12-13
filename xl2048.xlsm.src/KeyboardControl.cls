VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KeyboardControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IControlProvider

''
' Members
'
'
Private mEnable As Boolean
Private WithEvents GameWb As Workbook
Attribute GameWb.VB_VarHelpID = -1

''
' Interface
'
'
Private Manager As IGameManager

Private Sub IControlProvider_Register(GameManager As IGameManager)
Set GameManager.Control = Me
Set Manager = GameManager
End Sub

Private Property Let IControlProvider_Enable(Enable As Boolean)
If Enable Then
    Application.OnKey "{UP}", "'GameApp.KBdoMove " & toUp & "'"
    Application.OnKey "{DOWN}", "'GameApp.KBdoMove " & toDown & "'"
    Application.OnKey "{LEFT}", "'GameApp.KBdoMove " & toLeft & "'"
    Application.OnKey "{RIGHT}", "'GameApp.KBdoMove " & toRight & "'"
Else
    Application.OnKey "{UP}"
    Application.OnKey "{DOWN}"
    Application.OnKey "{LEFT}"
    Application.OnKey "{RIGHT}"
End If
mEnable = Enable
End Property

Private Property Get IControlProvider_Enable() As Boolean
IControlProvider_Enable = mEnable
End Property

''
' Event Callbacks
'
'
Public Sub Callback_DoMove(Direction As tDirection)
If Manager Is Nothing Then Exit Sub
Manager.doMove Direction
End Sub

''
' Events
'
'
Private Sub Class_Terminate()
IControlProvider_Enable = False
End Sub

Private Sub Class_Initialize()
Set GameWb = ThisWorkbook
End Sub

Private Sub GameWb_Activate()
IControlProvider_Enable = True
End Sub

Private Sub GameWb_Deactivate()
IControlProvider_Enable = False
End Sub
