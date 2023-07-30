VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ControlLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Encapsulates basic layout logic for dynamic MSForms controls."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Encapsulates basic layout logic for dynamic MSForms controls. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit

Public Enum AnchorEdges
    LeftAnchor = 1
    TopAnchor = 2
    RightAnchor = 4
    BottomAnchor = 8
    AnchorAll = LeftAnchor + TopAnchor + RightAnchor + BottomAnchor
End Enum

Private Type TControlLayout
    Anchors As AnchorEdges
    Margins As SideMargins
    ContainerSize As ControlSize
    ContainerItemSize As ControlSize
    BoundControl As MSForms.Control
End Type

Private Const DefaultMargin As Long = 3

Private This As TControlLayout

Private Sub Layout(ByVal Object As MSForms.Control)

    If (Me.Anchors And AnchorEdges.TopAnchor) = AnchorEdges.TopAnchor Then

        If (Me.Anchors And AnchorEdges.BottomAnchor) = AnchorEdges.BottomAnchor Then
            Object.Height = This.ContainerSize.Height - Object.Top - This.Margins.BottomMargin
        End If

    ElseIf (Me.Anchors And AnchorEdges.BottomAnchor) = AnchorEdges.BottomAnchor Then
        Object.Top = This.ContainerSize.Height - Object.Height - This.Margins.BottomMargin
    End If


    If (Me.Anchors And AnchorEdges.LeftAnchor) = AnchorEdges.LeftAnchor Then

        If (Me.Anchors And AnchorEdges.RightAnchor) = AnchorEdges.RightAnchor Then
            Object.Width = This.ContainerSize.Width - Object.Left - This.Margins.RightMargin
        End If

    ElseIf (Me.Anchors And AnchorEdges.RightAnchor) = AnchorEdges.RightAnchor Then
        Object.Left = This.ContainerSize.Width - Object.Width - This.Margins.RightMargin
        
    End If

End Sub

Public Sub Bind(ByVal a_container As Object, ByVal a_containerItem As MSForms.Control, ByVal a_anchorEdges As AnchorEdges)
    cc_isr_Core.GuardClauses.GuardNullReference a_container
    cc_isr_Core.GuardClauses.GuardNullReference a_containerItem
    
    Set This.ContainerSize = New ControlSize
    Set This.ContainerItemSize = New ControlSize
    Set This.Margins = New SideMargins
    This.Anchors = a_anchorEdges

    This.ContainerSize.Height = a_container.InsideHeight
    This.ContainerSize.Width = a_container.InsideWidth

    This.ContainerItemSize.Height = a_containerItem.Height
    This.ContainerItemSize.Width = a_containerItem.Width

    If (a_anchorEdges And AnchorEdges.BottomAnchor) = AnchorEdges.BottomAnchor Then
        This.Margins.BottomMargin = This.ContainerSize.Height - a_containerItem.Top - a_containerItem.Height
    End If

    If (a_anchorEdges And AnchorEdges.LeftAnchor) = AnchorEdges.LeftAnchor Then
        This.Margins.LeftMargin = a_containerItem.Left
    End If

    If (a_anchorEdges And AnchorEdges.RightAnchor) = AnchorEdges.RightAnchor Then
        This.Margins.RightMargin = This.ContainerSize.Width - a_containerItem.Left - a_containerItem.Width
    End If

    If (a_anchorEdges And AnchorEdges.TopAnchor) = AnchorEdges.TopAnchor Then
        This.Margins.TopMargin = a_containerItem.Top
    End If
    Set BoundControl = a_containerItem
    
End Sub

Public Sub Resize(ByVal a_object As Object)
    This.ContainerSize.Height = a_object.InsideHeight
    This.ContainerSize.Width = a_object.InsideWidth
    Layout BoundControl
End Sub

Private Sub Class_Terminate()
    Set This.ContainerSize = Nothing
    Set This.ContainerItemSize = Nothing
End Sub

Public Property Get Anchors() As AnchorEdges
    Anchors = This.Anchors
End Property

Public Property Let Anchors(ByVal a_value As AnchorEdges)
    This.Anchors = a_value
End Property

Public Property Get BoundControl() As MSForms.Control
    Set BoundControl = This.BoundControl
End Property

Public Property Set BoundControl(ByVal a_value As MSForms.Control)
    Set This.BoundControl = a_value
End Property

