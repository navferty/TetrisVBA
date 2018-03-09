VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Figure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'@Folder GameClasses.Figures

Private pFigureRange As Range
Private pMyType As FigureType

Public Enum MoveFigureDirection
    None
    ToLeft
    ToRight
    ToDown
    ToTurn
End Enum

Public Property Get FigureRange() As Range
    Set FigureRange = pFigureRange
End Property

Public Property Set FigureRange(newRange As Range)
    Set pFigureRange = newRange
End Property

Public Property Get MyType() As FigureType
    MyType = pMyType
End Property

Public Property Let MyType(RHS As FigureType)
    pMyType = RHS
End Property

Public Sub AddToRange(r As Range)
    If pFigureRange Is Nothing Then
        Set pFigureRange = r
    Else
        Set pFigureRange = Union(pFigureRange, r)
    End If
End Sub

Public Sub Draw()
    Dim c As Range
    For Each c In pFigureRange.Cells
        c.Interior.Color = vbGreen
    Next
End Sub

Public Sub Undraw()
    Dim c As Range
    For Each c In pFigureRange.Cells
        c.Interior.Color = vbWhite
    Next
End Sub

Public Sub MoveDown()
    If CheckFigureOnField(ToDown) Then
        Me.Undraw
        Set pFigureRange = GetOffsetRange(pFigureRange, ToDown)
        Me.Draw
    End If
End Sub

Public Sub MoveLeft()
    If CheckFigureOnField(ToLeft) Then
        Me.Undraw
        Set pFigureRange = GetOffsetRange(pFigureRange, ToLeft)
        Me.Draw
    End If
End Sub

Public Sub MoveRight()
    If CheckFigureOnField(ToRight) Then
        Me.Undraw
        Set pFigureRange = GetOffsetRange(pFigureRange, ToRight)
        Me.Draw
    End If
End Sub

Private Function CheckFigureOnField(d As MoveFigureDirection) As Boolean
    Dim c As Range
    Dim r As Range
    Dim f As Range
    
    Set f = ActiveSheet.Range(Cells(FieldTop, FieldLeft), Cells(FieldBottom, FieldRight))
    Set r = GetOffsetRange(pFigureRange, d)
    
    For Each c In r.Cells
        If Intersect(f, c) Is Nothing Then
            CheckFigureOnField = False
            Exit Function
        End If
    Next
    CheckFigureOnField = True
End Function

Public Sub TurnClockwise()
    'MsgBox "Not implemented"
    Dim baseCell As Range
    Dim c As Range
    Dim newC As Range
    
    Me.Undraw
    
    Dim newFigureRange As Range
    For Each c In pFigureRange.Cells
        If baseCell Is Nothing Then
            Set baseCell = c
            Set newFigureRange = c
        Else
            Set newC = baseCell.Offset(c.Column - baseCell.Column, baseCell.Row - c.Row)
            Set newFigureRange = Union(newFigureRange, newC)
        End If
    Next
    
    'not forget about borders!
    
    Set pFigureRange = newFigureRange
    
    Do While Not CheckFigureOnField(None)
        Me.Undraw
        Set pFigureRange = GetOffsetRange(pFigureRange, GetDirectionOfFigureOverflow)
        Me.Draw
    Loop
    
    Me.Draw
    'get offset for each cell (from base cell), switch row/col of offset with sign
    ' then, if out of field, move id proper direction till is on field
End Sub

Public Sub TurnCounterclockwise()
    MsgBox "Not implemented"
    'not forget about borders!
End Sub

Public Function GetTrialFigure(d As MoveFigureDirection) As Figure
    Dim newF As New Figure
    Set newF.FigureRange = GetOffsetRange(pFigureRange, d)
    Set GetTrialFigure = newF
End Function

Private Function GetDirectionOfFigureOverflow() As MoveFigureDirection
    Dim c As Range
    For Each c In pFigureRange.Cells
        If c.Column < FieldLeft Then
            GetDirectionOfFigureOverflow = ToRight
            Exit Function
        ElseIf c.Column > FieldRight Then
            GetDirectionOfFigureOverflow = ToLeft
            Exit Function
        End If
    Next
End Function

Private Function GetOffsetRange(initialRange As Range, d As MoveFigureDirection) As Range
    Dim c As Range
    Dim r As Range
    
    Dim rOffset As Integer
    Dim cOffset As Integer
    
    Select Case d
    Case MoveFigureDirection.ToDown
        rOffset = 1
    Case MoveFigureDirection.ToLeft
        cOffset = -1
    Case MoveFigureDirection.ToRight
        cOffset = 1
    End Select
    
    For Each c In initialRange.Cells
        If r Is Nothing Then
            Set r = c.Offset(rOffset, cOffset)
        Else
            Set r = Union(r, c.Offset(rOffset, cOffset))
        End If
    Next
    Set GetOffsetRange = r
End Function

