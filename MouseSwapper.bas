Attribute VB_Name = "mdlMouseSwapper"
Option Explicit
Public Declare Function SwapMouseButton Lib "user32" (ByVal bSwap As Long) As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
'frmMouseSwapper is only to use the icon for the compiled file

Public Sub Main()
Dim lngTmp As Long
Dim intTmp As Integer

'look for the Ctrl key pressed to set mouse buttons back to normal
intTmp = GetKeyState(vbKeyControl)

'If the &H1 bit of the return value is set, the key is toggled.
'If the &H8000 bit of the return value is set, the key is currently pressed down.
If intTmp And &H8000 Then
    lngTmp = SwapMouseButton(0)
Else
    lngTmp = SwapMouseButton(1)
End If

End Sub
