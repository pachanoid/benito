Attribute VB_Name = "modJoystick"
Option Explicit
'***  Color constants  ***
Public Const R01 = &HFF 'Red
Public Const G01 = &H808080    'Gray
Public Const G02 = &HC0C0C0    'Ligth Gray
Public Const BLK = &H0  'Black

'***  Math constants  ***
Public Const PI = 3.14159265358979

'***  File constants  ***
Public Const FSIZE = 32

'**********************************************************
'NOTE:  I did disabled some of the public constants with ',
'because this code isn't used in this program.
'**********************************************************

'***  joyGetPosEx  Constants  ***
'Public Const JOYSTICKID1 = 0
'Public Const JOYSTICKID2 = 1
Public Const JOY_POVCENTERED = &HFFFF
Public Const JOY_POVFORWARD = &H0
Public Const JOY_POVFRDRHT = &H1194
Public Const JOY_POVRIGHT = &H2328
Public Const JOY_POVBRDRHT = &H34BC
Public Const JOY_POVBACKWARD = &H4650
Public Const JOY_POVBRDLFT = &H57E4
Public Const JOY_POVLEFT = &H6978
Public Const JOY_POVFRDLFT = &H7B0C
'Public Const JOY_RETURNX = &H1&
'Public Const JOY_RETURNY = &H2&
'Public Const JOY_RETURNZ = &H4&
'Public Const JOY_RETURNR = &H8&
'Public Const JOY_RETURNU = &H10
'Public Const JOY_RETURNV = &H20
'Public Const JOY_RETURNPOV = &H40&
'Public Const JOY_RETURNBUTTONS = &H80&
'Public Const JOY_RETURNRAWDATA = &H100&
'Public Const JOY_RETURNPOVCTS = &H200&
'Public Const JOY_RETURNCENTERED = &H400&
'Public Const JOY_USEDEADZONE = &H800&
Public Const JOY_RETURNALL = &HFF
'Public Const JOY_CAL_READALWAYS = &H10000
'Public Const JOY_CAL_READXONLY = &H100000
'Public Const JOY_CAL_READ3 = &H40000
'Public Const JOY_CAL_READ4 = &H80000
'Public Const JOY_CAL_READXYONLY = &H20000
'Public Const JOY_CAL_READYONLY = &H200000
'Public Const JOY_CAL_READ5 = &H400000
'Public Const JOY_CAL_READ6 = &H800000
'Public Const JOY_CAL_READZONLY = &H1000000
'Public Const JOY_CAL_READRONLY = &H2000000
'Public Const JOY_CAL_READUONLY = &H4000000
'Public Const JOY_CAL_READVONLY = &H8000000

'***  sndPlaySound constants  ***
Public Const SND_ASYNC = &H1
'Public Const SND_LOOP = &H8
'Public Const SND_MEMORY = &H4
Public Const SND_NODEFAULT = &H2
'Public Const SND_NOSTOP = &H10
Public Const SND_SYNC = &H0

'***  Calibrate constants  ***
Public Const SenX = 1023              '  Sensitivity range to POV.
Public Const XY4D = 16383             '  Sensitivity Down limit to X+Y Axis.
Public Const XY4U = 49152             '  Sensitivity Up limit to X+Y Axis.

'***  Joystick registers  ***
Type JOYINFOEX
        dwSize As Long                '  size of structure
        dwFlags As Long               '  flags to indicate what to return
        dwXpos As Long                '  x position
        dwYpos As Long                '  y position
        dwZpos As Long                '  z position
        dwRpos As Long                '  rudder/4th axis position
        dwUpos As Long                '  5th axis position
        dwVpos As Long                '  6th axis position
        dwButtons As Long             '  button states
        dwButtonNumber As Long        '  current button number pressed
        dwPOV As Long                 '  point of view state
        dwReserved1 As Long           '  reserved for communication between winmm driver
        dwReserved2 As Long           '  reserved for future expansion
End Type

Type JOYCAPS
    wMid As Integer
    wPid As Integer
    szPname As String * 32
    wXmin As Long
    wXmax As Long
    wYmin As Long
    wYmax As Long
    wZmin As Long
    wZmax As Long
    wNumButtons As Long
    wPeriodMin As Long
    wPeriodMax As Long
    wRmin As Long
    wRmax As Long
    wUmin As Long
    wUmax As Long
    wVmin As Long
    wVmax As Long
    wCaps As Long
    wMaxAxes As Long
    wNumAxes As Long
    wMaxButtons As Long
    szRegKey As String * 32
    szOEMVxD As String * 260
End Type

'***  Joystick API functions ***
Declare Function joyGetPosEx Lib "winmm.dll" (ByVal uJoyID As Long, _
                 ByRef pji As JOYINFOEX) As Long
Declare Function joyGetDevCaps Lib "winmm.dll" Alias "joyGetDevCapsA" _
                (ByVal ID As Long, ByRef lpCaps As JOYCAPS, ByVal uSize As Long) As Long
Declare Function joyGetNumDevs Lib "winmm.dll" () As Long

Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" _
                (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
                
'***  Public Vars  ***
Public InfoJoyEX As JOYINFOEX 'Info-Joystick Register
Public CapX As JOYCAPS 'Caps-Joystick Register
Public xZbk As Long 'BackUp key of Z position
Public xPOV As Long 'BackUp key of POV position
Public IDJoy As Long 'ID Joystick Value
Public InSn As Boolean 'Play Sound key of Z change position
Public BnSn(0 To 31) As Boolean 'Play Sound key of Button pressed (binary string)
Public bPOV As Boolean 'Play sound key of POV status
Public bXYs(0 To 3, 0 To 3) As Boolean 'Play sound key of XY status
'**************************************************
               

Public Sub Sound(ByVal SName As String, ByVal F1 As Long)
'Play sound procedure
Dim ID As Long

ID = sndPlaySound(SName, F1)
     
'frmjoystick.ESound1.

End Sub

Function DecBin(ByVal decim As Long, ByVal tam As Integer) As String
'***  Convert integer number to binary string  ***
Dim co As Integer, binar As String
binar = ""
Do Until decim = 0
    If (decim Mod 2) = 1 Then
        binar = "1" + binar
    Else
        binar = "0" + binar
    End If
    decim = decim \ 2
Loop
For co = (Len(binar) + 1) To tam
    binar = "0" + binar
Next co
DecBin = binar

End Function

Public Function JoyEst(ByVal IDx As Long) As String
'Indentify the Joystick status.
Dim xRes As String
Select Case IDx
    Case 0
        xRes = "Connetcted"
    Case 167
        xRes = "Unplugged"
    Case Else
        xRes = "Unknown"
End Select
JoyEst = xRes

End Function

Public Function MPathX(ByVal sPhX As String) As String
'Add inverted slash if this not exist
Dim sPhY As String
sPhY = sPhX
If Mid(sPhY, Len(sPhY), 1) <> "\" Then
    sPhY = sPhY & "\"
End If
MPathX = sPhY

End Function

