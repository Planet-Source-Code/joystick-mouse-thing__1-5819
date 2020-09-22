VERSION 4.00
Begin VB.Form JoyVXD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   1275
   ClientLeft      =   285
   ClientTop       =   2235
   ClientWidth     =   1935
   ClipControls    =   0   'False
   Height          =   2055
   Icon            =   "JoyMouse.frx":0000
   Left            =   225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   1935
   ShowInTaskbar   =   0   'False
   Top             =   1515
   Visible         =   0   'False
   Width           =   2055
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   600
      Top             =   600
   End
   Begin VB.Menu PopUp 
      Caption         =   " "
      Begin VB.Menu mnuDef 
         Caption         =   "Define Buttons"
      End
      Begin VB.Menu mnuSpd 
         Caption         =   "Change Speed"
      End
      Begin VB.Menu mnuPause 
         Caption         =   "Pause"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "JoyVXD"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit
Dim Iconobject As Object
Dim Mouse As New Mouse
Dim ji As JOYINFOEX
Dim caps As JOYCAPS
Dim mask As Long
Sub DoWhatever(code As Long, Special As Byte)
Select Case code
   Case 0
      Select Case Special
         Case 0
            Mouse.LUp
         Case 1
            Mouse.LDown
      End Select
   Case 1
      Select Case Special
         Case 0
            Mouse.RUp
         Case 1
            Mouse.RDown
      End Select
End Select
End Sub

Private Sub Form_Load()
Dim rc As Long 'scratch
Set Iconobject = JoyVXD.Icon
AddIcon JoyVXD, Iconobject.Handle, Iconobject, "Animated TrayIcon"
rc = joyGetDevCaps(JOYSTICKID1, caps, Len(caps))
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static Message As Long
Message = X / Screen.TwipsPerPixelX
Select Case Message
Case WM_LBUTTONUP
Me.PopupMenu Popup
Case WM_RBUTTONUP
Me.PopupMenu Popup
End Select

End Sub



Private Sub Form_Unload(Cancel As Integer)
    delIcon Iconobject.Handle
    delIcon JoyVXD.Icon.Handle
End Sub


Private Sub mnuExit_Click()
Unload JoyVXD
'End
End Sub

Private Sub Timer1_Timer()

Static StateArray(255, 2) As Byte

Dim rc As Long 'scratch
Dim i As Long
Dim xpos As Long
Dim ypos As Long
Dim AlreadyDone As Boolean

ji.dwSize = Len(ji)
ji.dwFlags = JOY_RETURNALL

rc = joyGetPosEx(JOYSTICKID1, ji)

xpos = ji.dwXpos
ypos = ji.dwYpos

If xpos - 150 < 0 Then Mouse.X = Mouse.X - 15
If xpos + 150 > 64000 Then Mouse.X = Mouse.X + 15
If ypos - 150 < 0 Then Mouse.Y = Mouse.Y - 15
If ypos + 150 > 64000 Then Mouse.Y = Mouse.Y + 15

mask = 1

For i = 0 To (caps.wNumButtons - 1)
   
   If ji.dwButtons And mask Then StateArray(i, 2) = 1
                     
   mask = mask * 2
   
Next

For i = 0 To (caps.wNumButtons - 1)
   If StateArray(i, 2) <> StateArray(i, 1) Then
      DoWhatever i, StateArray(i, 2)
   End If
   StateArray(i, 1) = StateArray(i, 2)
   StateArray(i, 2) = 0
Next

End Sub


