VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SideMargins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit

Private Type TSideMargins
    BottomMargin As Double
    LeftMargin As Double
    RightMargin As Double
    TopMargin As Double
End Type

Private This As TSideMargins

Public Property Get BottomMargin() As Double
    BottomMargin = This.BottomMargin
End Property

Public Property Let BottomMargin(ByVal a_value As Double)
    This.BottomMargin = a_value
End Property

Public Property Get LeftMargin() As Double
    LeftMargin = This.LeftMargin
End Property

Public Property Let LeftMargin(ByVal a_value As Double)
    This.LeftMargin = a_value
End Property

Public Property Get RightMargin() As Double
    RightMargin = This.RightMargin
End Property

Public Property Let RightMargin(ByVal a_value As Double)
    This.RightMargin = a_value
End Property

Public Property Get TopMargin() As Double
    TopMargin = This.TopMargin
End Property

Public Property Let TopMargin(ByVal a_value As Double)
    This.TopMargin = a_value
End Property
