VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FigureFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'@Folder GameClasses.Figures

Public Enum FigureType
    Type_T
    Type_R
    Type_L
    Type_O
    Type_Z
    Type_S
    Type_I
    Type_Taxi
    Type_XXX
End Enum

Public Function CreateFigure(figType As FigureType, topLeftCornerCell As Range) As Figure
    Dim f As New Figure
    f.MyType = figType
    f.AddToRange topLeftCornerCell
    
    Select Case figType
    Case FigureType.Type_L
        f.AddToRange topLeftCornerCell.Offset(1, 0)
        f.AddToRange topLeftCornerCell.Offset(2, 0)
        f.AddToRange topLeftCornerCell.Offset(2, 1)
    Case FigureType.Type_O
        f.AddToRange topLeftCornerCell.Offset(1, 0)
        f.AddToRange topLeftCornerCell.Offset(0, 1)
        f.AddToRange topLeftCornerCell.Offset(1, 1)
    Case FigureType.Type_R
        f.AddToRange topLeftCornerCell.Offset(0, 1)
        f.AddToRange topLeftCornerCell.Offset(1, 0)
        f.AddToRange topLeftCornerCell.Offset(2, 0)
    Case FigureType.Type_S
        f.AddToRange topLeftCornerCell.Offset(1, 0)
        f.AddToRange topLeftCornerCell.Offset(1, 1)
        f.AddToRange topLeftCornerCell.Offset(2, 1)
    Case FigureType.Type_T
        f.AddToRange topLeftCornerCell.Offset(0, 1)
        f.AddToRange topLeftCornerCell.Offset(0, 2)
        f.AddToRange topLeftCornerCell.Offset(1, 1)
    Case FigureType.Type_Z
        f.AddToRange topLeftCornerCell.Offset(0, 1)
        f.AddToRange topLeftCornerCell.Offset(1, 1)
        f.AddToRange topLeftCornerCell.Offset(1, 2)
    Case FigureType.Type_I
        f.AddToRange topLeftCornerCell.Offset(1, 0)
        f.AddToRange topLeftCornerCell.Offset(2, 0)
        f.AddToRange topLeftCornerCell.Offset(3, 0)
    Case FigureType.Type_Taxi
        f.AddToRange topLeftCornerCell.Offset(1, 1)
        f.AddToRange topLeftCornerCell.Offset(2, 0)
        f.AddToRange topLeftCornerCell.Offset(3, 1)
    Case FigureType.Type_XXX
        f.AddToRange topLeftCornerCell.Offset(0, 1)
        f.AddToRange topLeftCornerCell.Offset(0, 2)
        f.AddToRange topLeftCornerCell.Offset(0, 3)
        f.AddToRange topLeftCornerCell.Offset(0, 4)
        f.AddToRange topLeftCornerCell.Offset(1, 2)
        f.AddToRange topLeftCornerCell.Offset(-1, 2)
    End Select
    Set CreateFigure = f
End Function

