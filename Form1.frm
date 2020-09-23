VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Mikes DirectX Particles (michaelpote@worldonline.co.za)"
   ClientHeight    =   5700
   ClientLeft      =   1920
   ClientTop       =   1560
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   380
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   439
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Pddsd As DDSURFACEDESC2, ScrDdsd As DDSURFACEDESC2
Dim Primary As DirectDrawSurface7
Dim Backbuffer As DirectDrawSurface7
Dim Surf As DirectDrawSurface7, SDDSD As DDSURFACEDESC2
Dim Ending As Boolean, bRestore As Boolean
Dim MX As Long, MY As Long
Dim PX(0 To 10000) As Long
Dim PY(0 To 10000) As Long
Dim PV(0 To 10000) As Long
Dim PSV(0 To 10000) As Long
Dim PL(0 To 10000) As Long
Dim PS(0 To 10000) As Long
Dim BX(0 To 13) As Single
Dim BY(0 To 13) As Single
Dim BWid(0 To 13) As Single
Dim BHgt(0 To 13) As Single
Dim I As Long, NumParticles As Long
Dim LFr As Long, Tim As Long, FPS As Long

Private Sub Form_DblClick()
EndIt
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 107 Then
NumParticles = NumParticles + 10
ElseIf KeyCode = 109 Then
NumParticles = NumParticles - 10
End If
End Sub

Private Sub Form_Load()
Init
InitSurfaces
FillBs

bRestore = False
Do Until ExModeActive
DoEvents
bRestore = True
Loop
DoEvents
If bRestore Then
bRestore = False
DD.RestoreAllSurfaces
InitSurfaces
End If

Dim ScreenRect As RECT
With ScreenRect
.Bottom = 600
.Right = 800
.Left = 0
.Top = 0
End With
Ending = False

Do While Ending = False
LFr = LFr + 1
If Tim < Int(Timer) Then
Tim = Int(Timer)
FPS = LFr
LFr = 0
End If
DoEvents
bRestore = False
Do Until ExModeActive
DoEvents
bRestore = True
Loop
DoEvents
If bRestore Then
bRestore = False
DD.RestoreAllSurfaces
InitSurfaces
End If

Backbuffer.BltColorFill ScreenRect, 0
Backbuffer.DrawText 0, 0, "Frames per second:" & FPS & "            Number of particles: " & NumParticles & " (+ to add, - to subtract particles)", True
For I = 0 To NumParticles
If PL(I) <= 0 Then
Randomize Timer
PS(I) = Int(Rnd * 13)      'Int(Rnd * 5) + 8 - Put this in for only smaller balls
PL(I) = Rnd * 400 + 100 'Change the 400 and the 100 for a Longer\Shorter life
PX(I) = MX + (Rnd * 4) - 2
PY(I) = MY + (Rnd * 4) - 2
PV(I) = Rnd * 12 - 6
PSV(I) = Rnd * 12 - 6
End If

PL(I) = PL(I) - 1
PV(I) = PV(I) - 1 'Put a 0 here for Zero Gravity

If PY(I) > 600 Then PY(I) = 600: PV(I) = -(PV(I) * (Rnd * 0.5) + 0.5)
If PX(I) <= 0 Or PX(I) >= 800 Then PSV(I) = -PSV(I)

PX(I) = PX(I) - (PSV(I) / 2)
PY(I) = PY(I) - (PV(I) / 2)
BltFast CLng(PX(I)), CLng(PY(I)), BX(PS(I)), BY(PS(I)), BWid(PS(I)), BHgt(PS(I)), Surf, True
Next

Primary.Flip Nothing, DDFLIP_WAIT
Loop
EndIt
End Sub

Public Sub Init()
Dim Ret As Boolean
On Local Error GoTo ErrorOut
Set DD = DirectX.DirectDrawCreate("")

Me.Show
DD.SetCooperativeLevel Me.hWnd, DDSCL_FULLSCREEN Or DDSCL_ALLOWMODEX Or DDSCL_EXCLUSIVE
DD.SetDisplayMode 800, 600, 16, 0, DDSDM_DEFAULT

     Pddsd.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
     Pddsd.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX
     Pddsd.lBackBufferCount = 1
     DoEvents
     Set Primary = DD.CreateSurface(Pddsd)
     Dim Caps As DDSCAPS2
     Caps.lCaps = DDSCAPS_BACKBUFFER
     Set Backbuffer = Primary.GetAttachedSurface(Caps)
     Backbuffer.GetSurfaceDesc ScrDdsd
     Backbuffer.SetFontTransparency True
     Backbuffer.SetForeColor vbWhite

InitFlag = True
Exit Sub
ErrorOut:
MsgBox "Sorry but an error occured while trying to setup directX." & Chr(13) & "Make sure you have the latest version of DirectX.", vbCritical
EndIt
End Sub

Public Sub EndIt()
Ending = True
DD.RestoreDisplayMode
DD.SetCooperativeLevel Me.hWnd, DDSCL_NORMAL
If Err.Description <> "" Then MsgBox Err.Description, vbCritical
DoEvents
Unload Me
End
End Sub

Public Sub BltFast(dX As Long, dY As Long, SrcX As Single, SrcY As Single, SrcWid As Single, SrcHgt As Single, ByRef Surf As DirectDrawSurface7, ColourKey As Boolean)
On Local Error GoTo Errot
If dX > 800 Or dY > 600 Then Exit Sub
Dim SrcRect As RECT, Retval

With SrcRect

If SrcY + dY <= 0 Then
.Top = -dY + SrcY
dY = 0
Else
.Top = SrcY
End If

If SrcX + dX <= 0 Then
.Left = -(X) + SrcX
dX = 0
Else
.Left = SrcX
End If

If dY + SrcHgt > 600 Then
.Bottom = SrcY + SrcHgt - ((dY + SrcHgt) - 600)
Else
.Bottom = SrcY + SrcHgt
End If

If dX + SrcWid > 800 Then
.Right = SrcX + SrcWid - ((dX + SrcWid) - 800)
Else
.Right = SrcX + SrcWid
End If

End With
If ColourKey Then
Retval = Backbuffer.BltFast(dX, dY, Surf, SrcRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
Else
Retval = Backbuffer.BltFast(dX, dY, Surf, SrcRect, DDBLTFAST_WAIT)
End If
Exit Sub
Errot:
EndIt
End Sub

Private Sub CreateSurface(ByRef Surf As DirectDrawSurface7, ByRef Ddsd As DDSURFACEDESC2, Filename As String, Wid As Integer, Hgt As Integer, ColourKey As Boolean)
'On Local Error GoTo SIE
Set Surf = Nothing
Ddsd.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
Ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
Ddsd.lWidth = Wid
Ddsd.lHeight = Hgt
Set Surf = DD.CreateSurfaceFromFile(Filename, Ddsd)
If ColourKey = True Then
Dim key As DDCOLORKEY
key.low = 0
key.high = 0
Surf.SetColorKey DDCKEY_SRCBLT, key
End If
Exit Sub
SIE:
EndIt
End Sub

Sub InitSurfaces()
'This is a little routine I wrote to load a bitmap into a directx surface
CreateSurface Surf, SDDSD, App.Path & "\Stuff.bmp", 300, 300, True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MX = X
MY = Y
End Sub


Sub FillBs()
BX(0) = 2
BY(0) = 3
BWid(0) = 14
BHgt(0) = 14
BX(1) = 17
BY(1) = 2
BWid(1) = 14
BHgt(1) = 14
BX(2) = 34
BY(2) = 1
BWid(2) = 15
BHgt(2) = 15
BX(3) = 51
BY(3) = 1
BWid(3) = 15
BHgt(3) = 15
BX(4) = 6
BY(4) = 21
BWid(4) = 5
BHgt(4) = 5
BX(5) = 21
BY(5) = 21
BWid(5) = 5
BHgt(5) = 5
BX(6) = 39
BY(6) = 22
BWid(6) = 5
BHgt(6) = 5
BX(7) = 54
BY(7) = 20
BWid(7) = 8
BHgt(7) = 8
BX(8) = 6
BY(8) = 34
BWid(8) = 5
BHgt(8) = 5
BX(9) = 16
BY(9) = 33
BWid(9) = 5
BHgt(9) = 5
BX(10) = 27
BY(10) = 34
BWid(10) = 5
BHgt(10) = 5
BX(11) = 36
BY(11) = 33
BWid(11) = 5
BHgt(11) = 5
BX(12) = 46
BY(12) = 33
BWid(12) = 5
BHgt(12) = 5
BX(13) = 56
BY(13) = 33
BWid(13) = 5
BHgt(13) = 5
End Sub
