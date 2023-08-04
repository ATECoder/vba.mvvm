VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SideMargins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder rd_MVVM.Infrastructure.View.Dynamic.Layout
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

Public Property Let BottomMargin(ByVal RHS As Double)
    This.BottomMargin = RHS
End Property

Public Property Get LeftMargin() As Double
    LeftMargin = This.LeftMargin
End Property

Public Property Let LeftMargin(ByVal RHS As Double)
    This.LeftMargin = RHS
End Property

Public Property Get RightMargin() As Double
    RightMargin = This.RightMargin
End Property

Public Property Let RightMargin(ByVal RHS As Double)
    This.RightMargin = RHS
End Property

Public Property Get TopMargin() As Double
    TopMargin = This.TopMargin
End Property

Public Property Let TopMargin(ByVal RHS As Double)
    This.TopMargin = RHS
End Property
