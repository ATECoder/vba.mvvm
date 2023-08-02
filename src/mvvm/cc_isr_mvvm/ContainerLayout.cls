VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ContainerLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Encapsulates basic layout logic for dynamic MSForms controls."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Encapsulates basic layout logic for dynamic MSForms controls. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit
Implements IContainerLayout

Private Type TState
    Parent As MSForms.Controls
    Direction As LayoutDirection
    NextControlTop As Long
    NextControlLeft As Long
    Margin As Long
End Type

Private Const DefaultMargin As Long = 3

Private This As TState

''' <summary>   Initializes and returns an <see cref="IContainerLayout"/> representation of
'''             this <see cref="ContainerLayout"/> object. </summary>
''' <returns>   [<see cref="IContainerLayout"/>]. </returns>
Public Function Initialize(ByVal a_parent As MSForms.Controls, Optional ByVal a_direction As LayoutDirection, _
        Optional ByVal a_top As Long = 0, _
        Optional ByVal a_left As Long = 0, _
        Optional ByVal a_margin As Long = DefaultMargin) As IContainerLayout
    
    Set Me.Parent = a_parent
    Me.NextControlTop = a_top + a_margin
    Me.NextControlLeft = a_left + a_margin
    Me.Direction = a_direction
    Me.Margin = a_margin
    
    Set Initialize = Me

End Function

Public Property Get Parent() As MSForms.Controls
    Set Parent = This.Parent
End Property

Public Property Set Parent(ByVal a_value As MSForms.Controls)
    Set This.Parent = a_value
End Property

Public Property Get NextControlTop() As Long
    NextControlTop = This.NextControlTop
End Property

Public Property Let NextControlTop(ByVal a_value As Long)
    This.NextControlTop = a_value
End Property

Public Property Get NextControlLeft() As Long
    NextControlLeft = This.NextControlLeft
End Property

Public Property Let NextControlLeft(ByVal a_value As Long)
    This.NextControlLeft = a_value
End Property

Public Property Get Direction() As LayoutDirection
    Direction = This.Direction
End Property

Public Property Let Direction(ByVal a_value As LayoutDirection)
    This.Direction = a_value
End Property

Public Property Get Margin() As Long
    Margin = This.Margin
End Property

Public Property Let Margin(ByVal a_value As Long)
    This.Margin = a_value
End Property

Private Sub OffsetControlSize(ByVal a_control As MSForms.Control, ByVal a_margin As Long)
    Select Case This.Direction
    
        Case LayoutDirection.TopToBottom
            This.NextControlTop = a_control.Top + a_control.Height + a_margin
        
        Case LayoutDirection.LeftToRight
            This.NextControlLeft = a_control.Left + a_control.Width + a_margin
            
        Case Else
            Debug.Assert False ' not implemented...
            
    End Select
End Sub

Private Function IContainerLayout_Add(ByVal ProgID As String, Optional ByVal a_height As Variant, _
        Optional ByVal a_width As Variant, Optional ByVal a_applyMargin As Boolean = True) As Object
    Dim p_result As Control
    Set p_result = This.Parent.Add(ProgID)
    
    p_result.Top = This.NextControlTop
    p_result.Left = This.NextControlLeft
    
    If Not VBA.IsMissing(a_height) Then
        p_result.Height = a_height
    End If
    
    If Not VBA.IsMissing(a_width) Then
        p_result.Width = a_width
    End If
    
    OffsetControlSize p_result, IIf(a_applyMargin, This.Margin, 0)
    Set IContainerLayout_Add = p_result
    
End Function

Private Sub IContainerLayout_SkipPosition(Optional ByVal a_margin As Long)
    Select Case This.Direction
        
        Case LayoutDirection.TopToBottom
            This.NextControlTop = This.NextControlTop + IIf(a_margin > 0, a_margin, This.Margin * 2)
            
        Case LayoutDirection.LeftToRight
            This.NextControlLeft = This.NextControlLeft + IIf(a_margin > 0, a_margin, This.Margin * 2)
            
    End Select
End Sub
