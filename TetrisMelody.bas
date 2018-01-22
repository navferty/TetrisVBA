Attribute VB_Name = "TetrisMelody"
'@Folder GameModules

Option Explicit

Public Declare PtrSafe Function Beep Lib "kernel32" _
(ByVal dwFreq As Long, _
ByVal dwDuration As Long) As Long

Public Sub StartMelody()
    Beep 658, 125
    Beep 1320, 500
    Beep 990, 250
    Beep 1056, 250
    Beep 1188, 250
    Beep 1320, 125
    Beep 1188, 125
    Beep 1056, 250
    Beep 990, 250
    Beep 880, 500
    Beep 880, 250
    Beep 1056, 250
    Beep 1320, 500
    Beep 1188, 250
    Beep 1056, 250
    Beep 990, 750
    Beep 1056, 250
    Beep 1188, 500
    Beep 1320, 500
    Beep 1056, 500
    Beep 880, 500
    Beep 880, 500
End Sub

Public Sub FullMelody()

    StartMelody

    Sleep 250
    Beep 1188, 500
    Beep 1408, 250
    Beep 1760, 500
    Beep 1584, 250
    Beep 1408, 250
    Beep 1320, 750
    Beep 1056, 250
    Beep 1320, 500
    Beep 1188, 250
    Beep 1056, 250
    Beep 990, 500
    Beep 990, 250
    Beep 1056, 250
    Beep 1188, 500
    Beep 1320, 500
    Beep 1056, 500
    Beep 880, 500
    Beep 880, 500

    Sleep 500
    Beep 1320, 500
    Beep 990, 250
    Beep 1056, 250
    Beep 1188, 250
    Beep 1320, 125
    Beep 1188, 125
    Beep 1056, 250
    Beep 990, 250
    Beep 880, 500
    Beep 880, 250
    Beep 1056, 250
    Beep 1320, 500
    Beep 1188, 250
    Beep 1056, 250
    Beep 990, 750
    Beep 1056, 250
    Beep 1188, 500
    Beep 1320, 500
    Beep 1056, 500
    Beep 880, 500
    Beep 880, 500

    Sleep 250
    Beep 1188, 500
    Beep 1408, 250
    Beep 1760, 500
    Beep 1584, 250
    Beep 1408, 250
    Beep 1320, 750
    Beep 1056, 250
    Beep 1320, 500
    Beep 1188, 250
    Beep 1056, 250
    Beep 990, 500
    Beep 990, 250
    Beep 1056, 250
    Beep 1188, 500
    Beep 1320, 500
    Beep 1056, 500
    Beep 880, 500
    Beep 880, 500

    Sleep 500
    Beep 660, 1000
    Beep 528, 1000
    Beep 594, 1000
    Beep 495, 1000
    Beep 528, 1000
    Beep 440, 1000
    Beep 419, 1000
    Beep 495, 1000
    Beep 660, 1000
    Beep 528, 1000
    Beep 594, 1000
    Beep 495, 1000
    Beep 528, 500
    Beep 660, 500
    Beep 880, 1000
    Beep 838, 2000
    Beep 660, 1000
    Beep 528, 1000
    Beep 594, 1000
    Beep 495, 1000
    Beep 528, 1000
    Beep 440, 1000
    Beep 419, 1000
    Beep 495, 1000
    Beep 660, 1000
    Beep 528, 1000
    Beep 594, 1000
    Beep 495, 1000
    Beep 528, 500
    Beep 660, 500
    Beep 880, 1000
    Beep 838, 2000
End Sub


