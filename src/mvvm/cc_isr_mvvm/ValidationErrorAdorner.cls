VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValidationErrorAdorner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An object that dynamically decorates a target MSForms.Control object."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   An object that dynamically decorates a target MSForms.Control object. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit
Implements IDynamicAdorner
Implements cc_isr_Core.IDisposable

Private Type TState
    Target As MSForms.Control
    TargetFormatter As cc_isr_MVVM.IValidationErrorFormatter
    
    AdornerLabel As MSForms.Label
    LabelFormatter As cc_isr_MVVM.IValidationErrorFormatter
    LabelControl As MSForms.Control
    
    AdornerIcon As MSForms.Image
    IconFormatter As cc_isr_MVVM.IValidationErrorFormatter
    IconControl As MSForms.Control
    
    LabelPosition As IDynamicPosition
    IconPosition As IDynamicPosition
    
    LabelPositionSet As Boolean
    IconPositionSet As Boolean
    Visible As Boolean
    
End Type

Private This As TState

''' <summary>   Initializes and returns an <see cref="IDynamicAdorner"/> representation of
'''             this <see cref="ValidationErrorAdornaer"/> object. </summary>
Public Function Initialize(ByVal a_target As MSForms.Control, _
        Optional ByVal a_targetFormatter As cc_isr_MVVM.IValidationErrorFormatter = Nothing, _
        Optional ByVal a_adornerLabel As Variant, _
        Optional ByVal a_labelFormatter As cc_isr_MVVM.IValidationErrorFormatter = Nothing, _
        Optional ByVal a_labelPosition As cc_isr_MVVM.IDynamicPosition = Nothing, _
        Optional ByVal a_adornerIcon As Variant, _
        Optional ByVal a_iconFormatter As cc_isr_MVVM.IValidationErrorFormatter = Nothing, _
        Optional ByVal a_iconPosition As cc_isr_MVVM.IDynamicPosition = Nothing) As IDynamicAdorner
    
    cc_isr_Core.GuardClauses.GuardNullReference a_target, VBA.Information.TypeName(Me)
    
    Me.InitTarget a_target, a_targetFormatter
    If VBA.IsMissing(a_adornerLabel) Then
        Set a_adornerLabel = Nothing
        Me.InitLabel a_adornerLabel, a_labelFormatter, a_labelPosition
    ElseIf IsObject(a_adornerLabel) Then
        If Not a_adornerLabel Is Nothing Then
            Me.InitLabel a_adornerLabel, a_labelFormatter, a_labelPosition
        End If
    End If
    
    If VBA.IsMissing(a_adornerIcon) Then
        Set a_adornerIcon = Nothing
        Me.InitIcon a_adornerIcon, a_iconFormatter, a_iconPosition
    ElseIf IsObject(a_adornerIcon) Then
        If Not a_adornerIcon Is Nothing Then
            Set a_adornerIcon = a_adornerIcon
            Me.InitIcon a_adornerIcon, a_iconFormatter, a_iconPosition
        End If
    End If
    
    Set Initialize = Me
    
End Function

Friend Sub InitTarget(ByVal a_target As Object, ByVal a_formatter As IValidationErrorFormatter)
    
    Set Me.Target = a_target
    If a_formatter Is Nothing Then
        Set Me.TargetFormatter = Me.DefaultTargetFormatter
    Else
        Set Me.TargetFormatter = a_formatter
    End If

End Sub

Friend Sub InitLabel(ByRef a_adornerLabel As Variant, ByRef Formatter As IValidationErrorFormatter, ByRef a_position As IDynamicPosition)
    
    If a_position Is Nothing Then Set a_position = DefaultLabelPosition
    
    If a_adornerLabel Is Nothing Then
        Dim Parent As MSForms.Controls
        Set Parent = Me.Target.Parent.Controls
        
        Set a_adornerLabel = Parent.Add(BindingDefaults.LabelProgId, Visible:=False)
        
        a_adornerLabel.BackStyle = fmBackStyleTransparent
        a_adornerLabel.WordWrap = False
        a_adornerLabel.AutoSize = True
        a_adornerLabel.Font.Size = Me.Target.Parent.Font.Size - 2
                        
    End If
    
    PositionAdornerControlX Me.Target, a_adornerLabel, a_position
    PositionAdornerControlY Me.Target, a_adornerLabel, a_position
    
    Set Me.AdornerLabel = a_adornerLabel
    Set Me.AdornerLabelPosition = a_position
    If Formatter Is Nothing Then
        Set Me.AdornerLabelFormatter = DefaultLabelFormatter
    Else
        Set Me.AdornerLabelFormatter = Formatter
    End If
End Sub

Friend Sub InitIcon(ByRef a_adornerIcon As Variant, ByRef Formatter As IValidationErrorFormatter, ByRef a_position As IDynamicPosition)
    If a_position Is Nothing Then Set a_position = DefaultIconPosition
    
    If a_adornerIcon Is Nothing Then
        Dim Parent As MSForms.Controls
        Set Parent = Me.Target.Parent.Controls
        
        Set a_adornerIcon = Parent.Add(BindingDefaults.ImageProgId, Visible:=False)
        Set a_adornerIcon.Picture = Resources.ValidationErrorIcon.Picture
        
        a_adornerIcon.BackStyle = fmBackStyleOpaque
        a_adornerIcon.BorderStyle = fmBorderStyleSingle
        a_adornerIcon.PictureSizeMode = fmPictureSizeModeClip
        a_adornerIcon.PictureAlignment = fmPictureAlignmentCenter
        a_adornerIcon.PictureTiling = False
        a_adornerIcon.AutoSize = True
        
    End If
    
    PositionAdornerControlX Me.Target, a_adornerIcon, a_position
    PositionAdornerControlY Me.Target, a_adornerIcon, a_position
    
    Set Me.AdornerIcon = a_adornerIcon
    Set Me.AdornerIconPosition = a_position
    If Formatter Is Nothing And Not a_adornerIcon Is Nothing Then
        Set Me.AdornerIconFormatter = DefaultIconFormatter
    Else
        Set Me.AdornerIconFormatter = Formatter
    End If
End Sub

Friend Property Get DefaultTargetFormatter() As cc_isr_MVVM.IValidationErrorFormatter

    Set DefaultTargetFormatter = Factory.NewValidationErrorFormatter() _
        .WithErrorBackgroundColor _
        .WithErrorBoldFont _
        .WithErrorBorderColor _
        .WithErrorForeColor
        
End Property

Friend Property Get DefaultLabelFormatter() As cc_isr_MVVM.IValidationErrorFormatter

    Set DefaultLabelFormatter = Factory.NewValidationErrorFormatter() _
        .WithErrorBoldFont _
        .WithErrorForeColor _
        .WithTargetOnlyVisibleOnError
        
End Property

Friend Property Get DefaultLabelPosition() As cc_isr_MVVM.IDynamicPosition

    Set DefaultLabelPosition = Factory.CreateDynamicControlPosition(RelativePosition.BelowRight, 3, 0)
    
End Property

Friend Property Get DefaultIconFormatter() As cc_isr_MVVM.IValidationErrorFormatter

    Set DefaultIconFormatter = Factory.NewValidationErrorFormatter() _
        .WithErrorBorderColor _
        .WithTargetOnlyVisibleOnError
        
End Property

Friend Property Get DefaultIconPosition() As cc_isr_MVVM.IDynamicPosition

    Set DefaultIconPosition = Factory.CreateDynamicControlPosition(RelativePosition.InsideRight, 3)

End Property

Public Property Get Target() As Object
    Set Target = This.Target
End Property

Public Property Set Target(ByVal a_value As Object)
    
    cc_isr_Core.GuardClauses.GuardDoubleInitialization This.Target, VBA.Information.TypeName(Me)
    Set This.Target = a_value

End Property

Public Property Get TargetFormatter() As IValidationErrorFormatter
    Set TargetFormatter = This.TargetFormatter
End Property

Public Property Set TargetFormatter(ByVal a_value As IValidationErrorFormatter)
    Set This.TargetFormatter = a_value
End Property

Public Property Get Visible() As Boolean
    Visible = This.Visible
End Property

Private Sub PositionAdornerControlX(ByVal a_target As MSForms.Control, ByVal Adorner As MSForms.Control, ByVal a_position As IDynamicPosition)
    Select Case a_position.Position
    
        Case RelativePosition.AboveRight, _
             RelativePosition.BelowRight, _
             RelativePosition.InsideRight, _
             RelativePosition.Default
            
            TrySetTextAlign Adorner, fmTextAlignRight
            Adorner.Left = a_target.Left + a_target.Width - Adorner.Width - a_position.Margin(RightSide)
            
        Case RelativePosition.AboveLeft, _
             RelativePosition.BelowLeft, _
             RelativePosition.InsideLeft
            
            TrySetTextAlign Adorner, fmTextAlignLeft
            Adorner.Left = a_target.Left + a_position.Margin(LeftSide)
            
    End Select
End Sub

Private Sub TrySetTextAlign(ByVal a_target As Object, ByVal a_value As MSForms.fmTextAlign)
    'late-bound assignment will fail if Target doesn't have a TextAlign property.
    On Error Resume Next
    'if execution stops here, go Tools ~> Options ~> General ~> Error Trapping,
    'and make sure "Break on unhandled errors" is selected.
    'only use "Break on all errors" for debugging & diagnostics.
    a_target.TextAlign = a_value
    On Error GoTo 0
End Sub

Private Sub PositionAdornerControlY(ByVal a_target As MSForms.Control, ByVal Adorner As MSForms.Control, ByVal a_position As IDynamicPosition)
    Select Case a_position.Position
    
        Case RelativePosition.AboveLeft, _
             RelativePosition.AboveRight
            
            Adorner.Top = a_target.Top - Adorner.Height - a_position.Margin(BottomSide)
            
        Case RelativePosition.InsideLeft, _
             RelativePosition.InsideRight
            
            Adorner.Top = a_target.Top + a_position.Margin(TopSide)
            
        Case RelativePosition.BelowLeft, _
             RelativePosition.BelowRight, _
             RelativePosition.Default
            
            Adorner.Top = a_target.Top + a_target.Height + a_position.Margin(TopSide)
            
    End Select
End Sub

Public Property Get AdornerLabel() As MSForms.Label
    Set AdornerLabel = This.AdornerLabel
End Property

Public Property Set AdornerLabel(ByVal a_value As MSForms.Label)
    Set This.AdornerLabel = a_value
    Set This.LabelControl = a_value
End Property

Public Property Get AdornerLabelFormatter() As IValidationErrorFormatter
    Set AdornerLabelFormatter = This.LabelFormatter
End Property

Public Property Set AdornerLabelFormatter(ByVal a_value As IValidationErrorFormatter)
    Set This.LabelFormatter = a_value
End Property

Public Property Get AdornerLabelPosition() As IDynamicPosition
    Set AdornerLabelPosition = This.LabelPosition
End Property

Public Property Set AdornerLabelPosition(ByVal a_value As IDynamicPosition)
    Set This.LabelPosition = a_value
End Property

Public Property Get AdornerIcon() As MSForms.Image
    Set AdornerIcon = This.AdornerIcon
End Property

Public Property Set AdornerIcon(ByVal a_value As MSForms.Image)
    Set This.AdornerIcon = a_value
    Set This.IconControl = a_value
End Property

Public Property Get AdornerIconFormatter() As IValidationErrorFormatter
    Set AdornerIconFormatter = This.IconFormatter
End Property

Public Property Set AdornerIconFormatter(ByVal a_value As IValidationErrorFormatter)
    Set This.IconFormatter = a_value
End Property

Public Property Get AdornerIconPosition() As IDynamicPosition
    Set AdornerIconPosition = This.IconPosition
End Property

Public Property Set AdornerIconPosition(ByVal a_value As IDynamicPosition)
    Set This.IconPosition = a_value
End Property

' - - - - - - - - - - - - - - - -
'
' Disposable implementation
'
' - - - - - - - - - - - - - - - -

Private Sub IDisposable_Dispose()
    Set This.Target = Nothing
    Set This.AdornerIcon = Nothing
    Set This.AdornerLabel = Nothing
    Set This.IconControl = Nothing
    Set This.LabelControl = Nothing
End Sub

Private Sub IDynamicAdorner_Hide()
    
    If Not This.Visible Then Exit Sub
    
    cc_isr_Core.GuardClauses.GuardNullReference This.TargetFormatter, VBA.Information.TypeName(Me), "TargetFormatter is not set for this instance."
    
    This.TargetFormatter.Restore This.Target
    If Not This.IconControl Is Nothing Then This.IconFormatter.Restore This.IconControl
    If Not This.LabelControl Is Nothing Then This.LabelFormatter.Restore This.LabelControl

    This.Visible = False

End Sub

Private Sub IDynamicAdorner_Show(ByVal a_message As String)
    
    'If This.Visible Then Exit Sub '<~ what if we just wanted to update the message?
    cc_isr_Core.GuardClauses.GuardNullReference This.TargetFormatter, VBA.Information.TypeName(Me), "TargetFormatter is not set for this instance."
    
    This.TargetFormatter.Apply This.Target, a_message
    If Not This.LabelControl Is Nothing Then ShowAdornerLabel a_message
    If Not This.IconControl Is Nothing Then ShowAdornerIcon a_message
    
    This.Visible = True

End Sub

Private Sub ShowAdornerLabel(ByVal a_message As String)
    This.LabelFormatter.Apply This.LabelControl, a_message
    If Not This.LabelPositionSet Then
        PositionAdornerControlX This.Target, This.AdornerLabel, This.LabelPosition
        PositionAdornerControlY This.Target, This.AdornerLabel, This.LabelPosition
        This.LabelPositionSet = True
    End If
End Sub

Private Sub ShowAdornerIcon(ByVal a_message As String)
    This.IconFormatter.Apply This.IconControl, a_message
    If Not This.IconPositionSet Then
        PositionAdornerControlX This.Target, This.AdornerIcon, This.IconPosition
        PositionAdornerControlY This.Target, This.AdornerIcon, This.IconPosition
        This.IconPositionSet = True
    End If
End Sub


