VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Layout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An object that encapsulates the UserForm Resizer"
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   An object that encapsulates the UserForm Layout"
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit

Implements ILayout

Private Type TState
    TotalMonitors As Long
    HorizontalResInPixel As Long
    VerticalResInPixel As Long
    
    Object As Object
    Width As Single
    Height As Single
    LayoutBindings As Collection
    
End Type

Private This As TState

''' <summary>   Initializes this <see cref="Layout"/> object. </summary>
Public Sub Initialize(ByVal a_object As Object, Optional ByVal a_newWidthPercent As Long = 50, _
        Optional ByVal a_newHeightPercent As Long = 50)
    
    cc_isr_Core.GuardClauses.GuardNullReference a_object, VBA.Information.TypeName(Me)
    
    This.HorizontalResInPixel = WinApi.HorizontalResInPixel
    This.VerticalResInPixel = WinApi.VerticalResInPixel
    
    Set Me.Object = a_object
    Me.Width = a_newWidthPercent
    Me.Height = a_newHeightPercent
    Set Me.LayoutBindings = New Collection
    
End Sub

Public Property Get Object() As Object
    Set Object = This.Object
End Property

Public Property Set Object(ByVal a_value As Object)
    Set This.Object = a_value
End Property

Public Property Get Width() As Single
    Width = This.Object.Width
End Property

Public Property Let Width(ByVal a_value As Single)

    cc_isr_Core.GuardClauses.GuardExpression Sgn(a_value - 0) + Sgn(a_value - 100), _
            "cc_isr_MVVM.Layout", "Value not between 0-100", _
            cc_isr_Core.UserDefinedErrors.InvalidArgumentError
    
    This.Width = WinApi.HorizontalResInPixel * (a_value / 100)
    
End Property

Public Property Get Height() As Single
    Height = This.Object.Height
End Property

Public Property Let Height(ByVal a_value As Single)

    cc_isr_Core.GuardClauses.GuardExpression Sgn(a_value - 0) + Sgn(a_value - 100), _
            "cc_isr_MVVM.Layout", "Value not between 0-100", _
            cc_isr_Core.UserDefinedErrors.InvalidArgumentError
            
    This.Height = WinApi.VerticalResInPixel * (a_value / 100)

End Property

Public Property Get LayoutBindings() As Collection

    Set LayoutBindings = This.LayoutBindings

End Property

Public Property Set LayoutBindings(ByVal a_value As Collection)
    
    cc_isr_Core.GuardClauses.GuardNullReference a_value, VBA.Information.TypeName(Me)
    
    Set This.LayoutBindings = a_value

End Property

Public Sub BindControlLayout(ByVal a_parent As Object, ByVal a_child As MSForms.Control, ByVal a_anchorEdges As AnchorEdges)
    
    cc_isr_Core.GuardClauses.GuardNullReference a_parent
    cc_isr_Core.GuardClauses.GuardNullReference a_child
    
    Dim Layout As ControlLayout
    Set Layout = New ControlLayout
    Layout.Bind a_parent, a_child, a_anchorEdges

    This.LayoutBindings.Add Layout
End Sub

Public Sub ResizeLayout()
    
    If This.Width < Object.Width Then Object.Width = BindingDefaults.LayoutDefaultWidth Else Object.Width = This.Width
    If This.Height < Object.Height Then Object.Height = BindingDefaults.LayoutDefaultHeight Else Object.Height = This.Height
    
    On Error GoTo CleanFail
    Dim Layout As ControlLayout
    For Each Layout In This.LayoutBindings
        Layout.Resize Object
    Next

CleanExit:
    Exit Sub

CleanFail:
    MsgBox VBA.Err.Description, Title:=VBA.Err.Number
    Resume CleanExit
    Resume
    
End Sub

Private Sub ILayout_BindControlLayout(ByVal a_parent As Object, ByVal a_child As MSForms.Control, ByVal a_anchorEdges As AnchorEdges)
    BindControlLayout a_parent, a_child, a_anchorEdges
End Sub

Private Sub ILayout_ResizeLayout()
    ResizeLayout
End Sub
