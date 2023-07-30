VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDynamicPosition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Represents the relative position and margin of a dynamic UI component."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Represents the relative position and margin of a dynamic UI component. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit

Public Enum RelativePosition
    Default = 0
    AboveLeft
    AboveRight
    BelowLeft
    BelowRight
    InsideLeft
    InsideRight
End Enum

Public Enum MarginSide
    LeftSide = 1
    TopSide = 2
    RightSide = 4
    BottomSide = 8
End Enum

''' <summary>   Gets the position of the adorner relative to its target. </summary>
Public Property Get Position() As RelativePosition
Attribute Position.VB_Description = "Gets the position of the adorner relative to its target."
End Property

''' <summary>   Gets the margin value for the specified side. </summary>
Public Property Get Margin(ByVal a_side As MarginSide) As Double
Attribute Margin.VB_Description = "Gets the margin value for the specified side."
End Property
