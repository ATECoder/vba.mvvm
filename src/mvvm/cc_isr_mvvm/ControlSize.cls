VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ControlSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit

Private Type TControlSize
    Height As Single
    Width As Single
End Type

Private This As TControlSize

Public Property Get Height() As Single
    Height = This.Height
End Property

Public Property Let Height(ByVal a_value As Single)
    This.Height = a_value
End Property

Public Property Get Width() As Single
    Width = This.Width
End Property

Public Property Let Width(ByVal a_value As Single)
    This.Width = a_value
End Property

