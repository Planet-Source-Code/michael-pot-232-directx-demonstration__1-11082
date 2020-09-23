Attribute VB_Name = "Module1"
Public DirectX As New DirectX7
Public DD As DirectDraw7


Function ExModeActive() As Boolean
     Dim TestCoopRes As Long 'holds the return value of the test.

     TestCoopRes = DD.TestCooperativeLevel 'Tells DDraw to do the test

     If (TestCoopRes = DD_OK) Then
         ExModeActive = True 'everything is fine
     Else
         ExModeActive = False 'this computer doesn't support this mode
     End If
End Function


