VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DynamicControlPosition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Represents the relative position and margin of a dynamic adorner component."
'@Folder rd_MVVM.Infrastructure.Validation.ErrorFormatting
'@ModuleDescription "Represents the relative position and margin of a dynamic adorner component."
'@PredeclaredId
'@Exposed
Option Explicit
Implements IDynamicPosition

Private Type TState
    Position As rd_MVVM.RelativePosition
    LeftMargin As Double
    TopMargin As Double
    RightMargin As Double
    BottomMargin As Double
End Type

Private This As TState

Public Function Create(ByVal Position As RelativePosition, ByVal LeftMargin As Double, Optional ByVal TopMargin As Variant, Optional ByVal RightMargin As Variant, Optional ByVal BottomMargin As Variant) As IDynamicPosition
    GuardClauses.GuardNonDefaultInstance Me, rd_MVVM.DynamicControlPosition
    
    Dim result As rd_MVVM.DynamicControlPosition
    Set result = New rd_MVVM.DynamicControlPosition
    result.Position = Position
    
    'if only left margin is specified, the margin applies to all sides.
    'if only left and top margins are specified, left+right match, and top+bottom match.
    'thus: Margin.Create(6) makes a margin of 6 units on all sides.
    '      Margin.Create(6, 3) makes a margin of 6 units left & right, and 3 units top & bottom.
    '      Margin.Create(6, 2, 3) makes a margin of 6 units left, 3 units right, and 2 units top & bottom.
    
    result.Margin(LeftSide) = LeftMargin
    
    If IsMissing(RightMargin) Then
        'if unspecified, assume same as left
        result.Margin(RightSide) = result.Margin(LeftSide)
    Else
        result.Margin(RightSide) = RightMargin
    End If
    
    If IsMissing(TopMargin) Then
        'if unspecified, assume same as left
        result.Margin(TopSide) = result.Margin(LeftSide)
    Else
        result.Margin(TopSide) = TopMargin
    End If
    
    If IsMissing(BottomMargin) Then
        'if unspecified, assume same as top
        result.Margin(BottomSide) = result.Margin(TopSide)
    Else
        result.Margin(BottomSide) = BottomMargin
    End If
    
    Set Create = result
End Function

Public Property Get Margin(ByVal Side As MarginSide) As Double
    Select Case Side
        Case MarginSide.BottomSide
            Margin = This.BottomMargin
        Case MarginSide.LeftSide
            Margin = This.LeftMargin
        Case MarginSide.RightSide
            Margin = This.RightMargin
        Case MarginSide.TopSide
            Margin = This.TopMargin
    End Select
End Property

Friend Property Let Margin(ByVal Side As MarginSide, ByVal RHS As Double)
    Select Case Side
        Case MarginSide.BottomSide
            This.BottomMargin = RHS
        Case MarginSide.LeftSide
            This.LeftMargin = RHS
        Case MarginSide.RightSide
            This.RightMargin = RHS
        Case MarginSide.TopSide
            This.TopMargin = RHS
    End Select
End Property

Public Property Get Position() As rd_MVVM.RelativePosition
    Position = This.Position
End Property

Friend Property Let Position(ByVal RHS As rd_MVVM.RelativePosition)
    This.Position = RHS
End Property

Private Property Get IDynamicPosition_Margin(ByVal Side As MarginSide) As Double
    IDynamicPosition_Margin = Me.Margin(Side)
End Property

Private Property Get IDynamicPosition_Position() As RelativePosition
    IDynamicPosition_Position = This.Position
End Property
