VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DynamicControlPosition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Represents the relative position and margin of a dynamic adorner component."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Represents the relative position and margin of a dynamic adorner component. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit
Implements IDynamicPosition

Private Type TState
    Position As cc_isr_MVVM.RelativePosition
    LeftMargin As Double
    TopMargin As Double
    RightMargin As Double
    BottomMargin As Double
End Type

Private This As TState

''' <summary>   Initializes this <see cref="DynamicControlPosition"/> object. </summary>
Public Sub Initialize(ByVal a_position As RelativePosition, _
        ByVal a_leftMargin As Double, _
        Optional ByVal a_topMargin As Variant, _
        Optional ByVal a_rightMargin As Variant, _
        Optional ByVal a_bottomMargin As Variant)
    
    Me.Position = a_position
    
    'if only left margin is specified, the margin applies to all sides.
    'if only left and top margins are specified, left+right match, and top+bottom match.
    'thus: Margin.Initialize(6) makes a margin of 6 units on all sides.
    '      Margin.Initialize(6, 3) makes a margin of 6 units left & right, and 3 units top & bottom.
    '      Margin.Initialize(6, 2, 3) makes a margin of 6 units left, 3 units right, and 2 units top & bottom.
    
    Me.Margin(LeftSide) = a_leftMargin
    
    If VBA.IsMissing(a_rightMargin) Then
        'if unspecified, assume same as left
        Me.Margin(RightSide) = Me.Margin(LeftSide)
    Else
        Me.Margin(RightSide) = a_rightMargin
    End If
    
    If VBA.IsMissing(a_topMargin) Then
        'if unspecified, assume same as left
        Me.Margin(TopSide) = Me.Margin(LeftSide)
    Else
        Me.Margin(TopSide) = a_topMargin
    End If
    
    If VBA.IsMissing(a_bottomMargin) Then
        'if unspecified, assume same as top
        Me.Margin(BottomSide) = Me.Margin(TopSide)
    Else
        Me.Margin(BottomSide) = a_bottomMargin
    End If
    
End Sub

Public Property Get Margin(ByVal a_side As MarginSide) As Double
    Select Case a_side
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

Public Property Let Margin(ByVal a_side As MarginSide, ByVal a_value As Double)
    Select Case a_side
        Case MarginSide.BottomSide
            This.BottomMargin = a_value
        Case MarginSide.LeftSide
            This.LeftMargin = a_value
        Case MarginSide.RightSide
            This.RightMargin = a_value
        Case MarginSide.TopSide
            This.TopMargin = a_value
    End Select
End Property

Public Property Get Position() As cc_isr_MVVM.RelativePosition
    Position = This.Position
End Property

Public Property Let Position(ByVal a_value As cc_isr_MVVM.RelativePosition)
    This.Position = a_value
End Property

Private Property Get IDynamicPosition_Margin(ByVal a_side As MarginSide) As Double
    IDynamicPosition_Margin = Me.Margin(a_side)
End Property

Private Property Get IDynamicPosition_Position() As RelativePosition
    IDynamicPosition_Position = This.Position
End Property
