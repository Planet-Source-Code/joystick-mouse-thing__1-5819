VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Mouse"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

Private Type PointAPI
    X As Long
    Y As Long
End Type

Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

Private Declare Function GetCursorPos Lib "user32" _
(lpPoint As PointAPI) As Long

Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long

Dim MousePos As PointAPI

Const MOUSEEVENTF_LEFTDOWN = &H2
Const MOUSEEVENTF_LEFTUP = &H4
Const MOUSEEVENTF_MIDDLEDOWN = &H20
Const MOUSEEVENTF_MIDDLEUP = &H40
Const MOUSEEVENTF_RIGHTDOWN = &H8
Const MOUSEEVENTF_RIGHTUP = &H10

Public Sub MDown()


    mouse_event MOUSEEVENTF_MIDDLEDOWN, 0, 0, 0, 0
End Sub



Public Sub MUp()


    mouse_event MOUSEEVENTF_MIDDLEUP, 0, 0, 0, 0
End Sub
Public Sub RDown()


    mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0
End Sub



Public Sub RUp()


    mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
End Sub
Public Sub LDown()
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
End Sub

Public Sub LUp()
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
End Sub
Property Let X(NewX)
SetCursorPos NewX, Y
End Property

Property Get X()
GetCursorPos MousePos
X = MousePos.X
End Property
Property Get Y()
GetCursorPos MousePos
Y = MousePos.Y
End Property


Property Let Y(NewY)
SetCursorPos X, NewY
End Property

