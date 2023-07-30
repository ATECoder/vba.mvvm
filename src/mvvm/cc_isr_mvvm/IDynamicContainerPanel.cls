VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDynamicContainerPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Describes a panel layout component."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Describes a panel layout component. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit

Public Enum LayoutDirection
    TopToBottom
    LeftToRight
    RightToLeft
End Enum

''' <summary>   Gets a value that controls how child controls are sized. </summary>
Public Property Get Direction() As LayoutDirection
Attribute Direction.VB_Description = "Gets a value that controls how child controls are sized."
End Property

''' <summary>   Gets the width of the panel. </summary>
Public Property Get Width() As Single
Attribute Width.VB_Description = "Gets the width of the panel."
End Property

''' <summary>   Gets the height of the panel. </summary>
Public Property Get Height() As Single
Attribute Height.VB_Description = "Gets the height of the panel."
End Property
