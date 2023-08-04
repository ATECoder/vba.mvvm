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
'@Folder MVVM.Infrastructure.View.Dynamic.Layout
'@ModuleDescription "Encapsulates basic layout logic for dynamic MSForms controls."
'@Exposed
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

    If (Anchors And TopAnchor) = TopAnchor Then

        If (Anchors And BottomAnchor) = BottomAnchor Then
            Object.Height = This.ContainerSize.Height - Object.Top - This.Margins.BottomMargin
        End If

    ElseIf (Anchors And BottomAnchor) = BottomAnchor Then
        Object.Top = This.ContainerSize.Height - Object.Height - This.Margins.BottomMargin
    End If


    If (Anchors And LeftAnchor) = LeftAnchor Then

        If (Anchors And RightAnchor) = RightAnchor Then
            Object.Width = This.ContainerSize.Width - Object.Left - This.Margins.RightMargin
        End If

    ElseIf (Anchors And RightAnchor) = RightAnchor Then
        Object.Left = This.ContainerSize.Width - Object.Width - This.Margins.RightMargin
        
    End If

End Sub

Public Sub Bind(ByVal Container As Object, ByVal ContainerItem As MSForms.Control, ByVal Anchor As AnchorEdges)
    GuardClauses.GuardNullReference Container
    GuardClauses.GuardNullReference ContainerItem
    
    Set This.ContainerSize = New ControlSize
    Set This.ContainerItemSize = New ControlSize
    Set This.Margins = New SideMargins
    This.Anchors = Anchor

    This.ContainerSize.Height = Container.InsideHeight
    This.ContainerSize.Width = Container.InsideWidth

    This.ContainerItemSize.Height = ContainerItem.Height
    This.ContainerItemSize.Width = ContainerItem.Width

    If (Anchor And BottomAnchor) = BottomAnchor Then
        This.Margins.BottomMargin = This.ContainerSize.Height - ContainerItem.Top - ContainerItem.Height
    End If

    If (Anchor And LeftAnchor) = LeftAnchor Then
        This.Margins.LeftMargin = ContainerItem.Left
    End If

    If (Anchor And RightAnchor) = RightAnchor Then
        This.Margins.RightMargin = This.ContainerSize.Width - ContainerItem.Left - ContainerItem.Width
    End If

    If (Anchor And TopAnchor) = TopAnchor Then
        This.Margins.TopMargin = ContainerItem.Top
    End If
    Set BoundControl = ContainerItem
    
End Sub

Public Sub Resize(ByVal Object As Object)
    This.ContainerSize.Height = Object.InsideHeight
    This.ContainerSize.Width = Object.InsideWidth
    Layout BoundControl
End Sub

Private Sub Class_Terminate()
    Set This.ContainerSize = Nothing
    Set This.ContainerItemSize = Nothing
End Sub

Public Property Get Anchors() As AnchorEdges
    Anchors = This.Anchors
End Property

Public Property Let Anchors(ByVal RHS As AnchorEdges)
    This.Anchors = RHS
End Property

Public Property Get BoundControl() As MSForms.Control
    Set BoundControl = This.BoundControl
End Property

Public Property Set BoundControl(ByVal RHS As MSForms.Control)
    Set This.BoundControl = RHS
End Property

