Attribute VB_Name = "TetrisGame"
'@Folder GameModules

Option Explicit

Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Declare PtrSafe Function GetKeyboardState Lib "User32.DLL" (kbArray As KeyboardBytes) As Long

Type KeyboardBytes
    kbb(0 To 255) As Byte
End Type

Public foreignApp As New Application

Public Const FieldTop As Integer = 3
Public Const FieldBottom As Integer = 22
Public Const FieldLeft As Integer = 5
Public Const FieldRight As Integer = 14


Public Sub Test()
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim d As MoveFigureDirection
    
    Dim ff As New FigureFactory
    Dim f As Figure
    Dim nextF As Figure
    Dim bh As New BottomHeap
    
    Dim startCol As Integer
    Dim nextFigureType As FigureType
    
    Dim isExtended As Boolean
    Dim typeMultiplier As Variant
    
    isExtended = ActiveSheet.Range("Q18").Value
    
    typeMultiplier = IIf(isExtended, CDec(7.7), CDec(6))
    
    DisableArrowKeys
    
    Sleep 100
    
    Do While True
        
        Randomize
        startCol = FieldLeft + Round(Rnd() * (FieldRight - FieldLeft - 3), 0)
        Randomize
        
        Set f = ff.CreateFigure(nextFigureType, ActiveSheet.Cells(FieldTop + 1, startCol))
        
        Randomize
        nextFigureType = Round(Rnd() * typeMultiplier, 0)
        Randomize
        
        
        
        Set nextF = ff.CreateFigure(nextFigureType, ActiveSheet.Range("S3"))
        
        nextF.Draw
        
        If bh.CheckIfIntersect(f) Then
            MsgBox "Ooops"
            StartMelody
            bh.Undraw
            nextF.Undraw
            EnableArrowKeys
            End
        End If
        
        f.Draw
        Sleep 200
        DoEvents
        
        Do While Not bh.CheckFigureIsDropped(f)
            
            f.MoveDown
            
            For j = 1 To 5
                For k = 1 To 50
                    If d = None Then
                        d = GetDirectionFromKeyboard
                    End If
                    Sleep 1
                Next k
                
                Select Case d
                Case MoveFigureDirection.ToDown
                    Do While Not bh.CheckFigureIsDropped(f)
                        f.MoveDown
                        Sleep 10
                        DoEvents
                    Loop
                    d = None
                Case MoveFigureDirection.ToLeft
                    If Not bh.CheckIfIntersect(f.GetTrialFigure(ToLeft)) Then f.MoveLeft
                    d = None
                Case MoveFigureDirection.ToRight
                    If Not bh.CheckIfIntersect(f.GetTrialFigure(ToRight)) Then f.MoveRight
                    d = None
                Case MoveFigureDirection.ToTurn
                    f.TurnClockwise
                    d = None
                End Select
                
                Sleep 50
                DoEvents
            Next j
            
        Loop
        
        bh.AddFigure f
        
        bh.ClearFilledRows
        
        nextF.Undraw
        
    Loop
    
    bh.Undraw
    foreignApp.Quit
End Sub

Private Function GetDirectionFromKeyboard() As MoveFigureDirection
    Dim kbArray As KeyboardBytes
    GetKeyboardState kbArray
    If kbArray.kbb(37) = 128 Then
        GetDirectionFromKeyboard = ToLeft
    ElseIf kbArray.kbb(38) = 128 Then
        GetDirectionFromKeyboard = ToTurn
    ElseIf kbArray.kbb(39) = 128 Then
        GetDirectionFromKeyboard = ToRight
    ElseIf kbArray.kbb(40) = 128 Then
        GetDirectionFromKeyboard = ToDown
    Else
        GetDirectionFromKeyboard = None
    End If
End Function

Private Sub EnableArrowKeys()
    With Application
        .OnKey "{UP}"
        .OnKey "{DOWN}"
        .OnKey "{LEFT}"
        .OnKey "{RIGHT}"
    End With
End Sub

Private Sub DisableArrowKeys()
    With Application
        .OnKey "{UP}", ""
        .OnKey "{DOWN}", ""
        .OnKey "{LEFT}", ""
        .OnKey "{RIGHT}", ""
    End With
End Sub

