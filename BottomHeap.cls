VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BottomHeap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'@Folder GameClasses

Private pHeapRange As Range

Public Property Get HeapRange() As Range
    Set HeapRange = pHeapRange
End Property

Public Property Set HeapRange(ByVal vNewValue As Range)
    Set pHeapRange = vNewValue
End Property

Public Sub Clear()
    
End Sub

Public Sub ClearFilledRows()
    Dim i As Integer
    Dim j As Integer
    Dim newHeapRange As Range
    
    Dim f As Range
    Set f = ActiveSheet.Range(Cells(FieldTop, FieldLeft), Cells(FieldBottom, FieldRight))
    
    Me.Undraw
    
    For i = FieldBottom To FieldTop Step -1
        If CheckRowFilled(i) Then
            Set newHeapRange = Nothing
            For j = FieldBottom To i + 1 Step -1
                AddRowToRange newHeapRange, j, 0
            Next j
            For j = i - 1 To FieldTop Step -1
                AddRowToRange newHeapRange, j, 1
            Next j
            Set pHeapRange = newHeapRange
            i = i + 1
        End If
    Next i
    
    Me.Draw
End Sub

Private Sub AddRowToRange(r As Range, rowIndex As Integer, rowOffset As Integer)
    Dim c As Range
    For Each c In pHeapRange.Cells
        If c.Row = rowIndex Then
            If r Is Nothing Then
                Set r = c.Offset(rowOffset, 0)
            Else
                Set r = Union(r, c.Offset(rowOffset, 0))
            End If
        End If
    Next
End Sub

Private Function CheckRowFilled(rowIndex As Integer) As Boolean
    Dim i As Integer
    
    If pHeapRange Is Nothing Then Exit Function
    
    For i = FieldLeft To FieldRight
        If Intersect(pHeapRange, ActiveSheet.Cells(rowIndex, i)) Is Nothing Then
            CheckRowFilled = False
            Exit Function
        End If
    Next i
    CheckRowFilled = True
End Function

Public Function CheckFigureIsDropped(f As Figure) As Boolean
    Dim c As Range
    For Each c In f.FigureRange.Cells
        If Not pHeapRange Is Nothing Then
            If Not Intersect(c.Offset(1, 0), pHeapRange) Is Nothing Then
                CheckFigureIsDropped = True
                Exit Function
            End If
        End If
        If c.Row >= FieldBottom Then
            CheckFigureIsDropped = True
            Exit Function
        End If
    Next
End Function

Public Sub AddFigure(f As Figure)
    If pHeapRange Is Nothing Then
        Set pHeapRange = f.FigureRange
    Else
        Set pHeapRange = Union(pHeapRange, f.FigureRange)
    End If
End Sub

Public Sub Draw()
    Dim c As Range
    If pHeapRange Is Nothing Then Exit Sub
    
    For Each c In pHeapRange.Cells
        c.Interior.Color = vbGreen
    Next
End Sub

Public Sub Undraw()
    Dim c As Range
    If pHeapRange Is Nothing Then Exit Sub
    
    For Each c In pHeapRange.Cells
        c.Interior.Color = vbWhite
    Next
End Sub

Public Function CheckIfIntersect(f As Figure) As Boolean
    Dim c As Range
    For Each c In f.FigureRange.Cells
        If Not pHeapRange Is Nothing Then
            If Not Intersect(pHeapRange, c) Is Nothing Then
                CheckIfIntersect = True
                Exit Function
            End If
        End If
    Next
End Function

