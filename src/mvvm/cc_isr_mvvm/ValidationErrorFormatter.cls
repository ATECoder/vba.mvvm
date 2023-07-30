VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValidationErrorFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An object that describes how a control changes its appearance given a validation error."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   An object that describes how a control changes its appearance given a validation error.
'''             Handles MSForms and Excel.Range targets. ...which means the class is doing way too many
'''             things for its own good. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit
Implements IValidationErrorFormatter

Private Const DefaultErrorBackColor As Long = &HC0C0FF
Private Const DefaultErrorBorderColor As Long = &HC0
Private Const DefaultErrorBorderWidth As Long = 2
Private Const DefaultErrorForeColor As Long = &HC0

Private Type TState
    Applied As Boolean
    
    FormatBackgroundColor As Boolean
    FormatBorderColor As Boolean
    FormatForeColor As Boolean
    FormatFontBold As Boolean
    
    InitialBackgroundColor As Long
    InitialBorderColor As Long
    InitialBorderWidth As Double
    InitialForeColor As Long
    InitialFontBold As Boolean
    
    ErrorBackgroundColor As Long
    ErrorBorderColor As Long
    ErrorBorderWidth As Double
    ErrorForeColor As Long
    ErrorFontBold As Boolean
    
    TargetOnlyVisibleOnError As Boolean
    PropertyName As String
    
End Type

Private This As TState

''' <summary>   A builder method configure a formatter that sets the target's background color when applied. </summary>
Public Function WithErrorBackgroundColor(Optional ByVal a_color As Variant = DefaultErrorBackColor) As ValidationErrorFormatter
Attribute WithErrorBackgroundColor.VB_Description = "A builder method to create or configure a formatter that sets the target's background color when applied."
    
    This.ErrorBackgroundColor = a_color
    This.FormatBackgroundColor = True
    Set WithErrorBackgroundColor = Me

End Function

''' <summary>   A builder method to create or configure a formatter that sets the target's border color when applied. </summary>
Public Function WithErrorBorderColor(Optional ByVal a_color As Variant = DefaultErrorBorderColor) As ValidationErrorFormatter
Attribute WithErrorBorderColor.VB_Description = "A builder method to create or configure a formatter that sets the target's border color when applied."
    
    This.ErrorBorderColor = a_color
    This.FormatBorderColor = True
    Set WithErrorBorderColor = Me

End Function

''' <summary>   A builder method to create or configure a formatter that sets the target's border
''' weight when applied (Excel.Range targets only). </summary>
Public Function WithErrorBorderWidth(Optional ByVal a_width As Double = DefaultErrorBorderWidth) As ValidationErrorFormatter
Attribute WithErrorBorderWidth.VB_Description = "A builder method to create or configure a formatter that sets the target's border weight when applied (Excel.Range targets only)."
    
    This.ErrorBorderWidth = a_width
    This.FormatBorderColor = True
    Set WithErrorBorderWidth = Me
    
End Function

''' <summary>   A builder method to create or configure a formatter that sets the target's forecolor (i.e. text color) when applied. </summary>
Public Function WithErrorForeColor(Optional ByVal a_color As Variant = DefaultErrorForeColor) As ValidationErrorFormatter
Attribute WithErrorForeColor.VB_Description = "A builder method to create or configure a formatter that sets the target's forecolor (i.e. text color) when applied."
    
    This.ErrorForeColor = a_color
    This.FormatForeColor = True
    Set WithErrorForeColor = Me
    
End Function

''' <summary>   A builder method to create or configure a formatter that makes the target's font bold when applied. </summary>
Public Function WithErrorBoldFont() As ValidationErrorFormatter
Attribute WithErrorBoldFont.VB_Description = "A builder method to create or configure a formatter that makes the target's font bold when applied."
        
    This.FormatFontBold = True
    This.ErrorFontBold = True
    Set WithErrorBoldFont = Me
    
End Function

''' <summary>   A builder method to create or configure a formatter that makes the target only visible when formatter is applied. </summary>
Public Function WithTargetOnlyVisibleOnError() As ValidationErrorFormatter
Attribute WithTargetOnlyVisibleOnError.VB_Description = "A builder method to create or configure a formatter that makes the target only visible when formatter is applied."
    
    This.TargetOnlyVisibleOnError = True
    Set WithTargetOnlyVisibleOnError = Me

End Function

Private Sub StoreInitialFormat(ByVal a_value As Object)
    
    Select Case True
    
        Case TypeOf a_value Is MSForms.TextBox
        
            Dim p_textBoxTarget As MSForms.TextBox
            Set p_textBoxTarget = a_value
            This.InitialBackgroundColor = p_textBoxTarget.BackColor
            This.InitialBorderColor = p_textBoxTarget.BorderColor
            This.InitialFontBold = p_textBoxTarget.Font.Bold
            This.InitialForeColor = p_textBoxTarget.ForeColor
            
        Case TypeOf a_value Is MSForms.CheckBox 'NOTE: MSForms.OptionButton also matches this interface
        
            Dim p_checkBoxTarget As MSForms.CheckBox
            Set p_checkBoxTarget = a_value
            This.InitialFontBold = p_checkBoxTarget.Font.Bold
            This.InitialForeColor = p_checkBoxTarget.ForeColor
            
        Case TypeOf a_value Is MSForms.Label
        
            Dim p_labelTarget As MSForms.Label
            Set p_labelTarget = a_value
            This.InitialBackgroundColor = p_labelTarget.BackColor
            This.InitialBorderColor = p_labelTarget.BorderColor
            This.InitialFontBold = p_labelTarget.Font.Bold
            This.InitialForeColor = p_labelTarget.ForeColor
            
        Case TypeOf a_value Is MSForms.Frame
        
            Dim p_frameTarget As MSForms.Frame
            Set p_frameTarget = a_value
            This.InitialBackgroundColor = p_frameTarget.BackColor
            This.InitialBorderColor = p_frameTarget.BorderColor
            This.InitialFontBold = p_frameTarget.Font.Bold
            This.InitialForeColor = p_frameTarget.ForeColor
        
        Case TypeOf a_value Is MSForms.Image
        
            Dim p_imageTarget As MSForms.Image
            Set p_imageTarget = a_value
            This.InitialBackgroundColor = p_imageTarget.BackColor
            This.InitialBorderColor = p_imageTarget.BorderColor
            
        Case TypeOf a_value Is Excel.Range
        
            Dim p_rangeTarget As Excel.Range
            Set p_rangeTarget = a_value
            This.InitialBackgroundColor = p_rangeTarget.Interior.Color
            This.InitialBorderColor = p_rangeTarget.Borders.Color
            This.InitialBorderWidth = p_rangeTarget.Borders.Weight
            This.InitialFontBold = p_rangeTarget.Font.Bold
            This.InitialForeColor = p_rangeTarget.Font.Color
            
        Case Else
            Debug.Print VBA.Information.TypeName(Me) & ": Target of type " & _
                        VBA.Information.TypeName(a_value) & " is not currently supported."
            
    End Select
    
End Sub

Private Sub IValidationErrorFormatter_Apply(ByVal a_target As Object, ByVal a_message As String)

    If This.Applied Then Exit Sub
    StoreInitialFormat a_target
    
    If TypeOf a_target Is MSForms.Control Then
        Dim Control As MSForms.Control
        Set Control = a_target
        If Not This.TargetOnlyVisibleOnError Then
            Control.ControlTipText = a_message
        End If
    End If
    
    Select Case True
    
        Case TypeOf a_target Is MSForms.TextBox
            Dim TextBoxTarget As MSForms.TextBox
            Set TextBoxTarget = a_target
            If This.FormatBackgroundColor Then TextBoxTarget.BackColor = This.ErrorBackgroundColor
            If This.FormatBorderColor Then TextBoxTarget.BorderColor = This.ErrorBorderColor
            If This.FormatFontBold Then TextBoxTarget.Font.Bold = This.ErrorFontBold
            If This.FormatForeColor Then TextBoxTarget.ForeColor = This.ErrorForeColor
            This.Applied = True
            
        Case TypeOf a_target Is MSForms.CheckBox 'NOTE: MSForms.OptionButton also matches this interface
            Dim CheckBoxTarget As MSForms.CheckBox
            Set CheckBoxTarget = a_target
            If This.FormatFontBold Then CheckBoxTarget.Font.Bold = This.ErrorFontBold
            If This.FormatForeColor Then CheckBoxTarget.ForeColor = This.ErrorForeColor
            This.Applied = True
            
        Case TypeOf a_target Is MSForms.Label
            Dim LabelTarget As MSForms.Label
            Set LabelTarget = a_target
            If This.FormatBackgroundColor Then LabelTarget.BackColor = This.ErrorBackgroundColor
            If This.FormatBorderColor Then LabelTarget.BorderColor = This.ErrorBorderColor
            If This.FormatFontBold Then LabelTarget.Font.Bold = This.ErrorFontBold
            If This.FormatForeColor Then LabelTarget.ForeColor = This.ErrorForeColor
            If This.TargetOnlyVisibleOnError Then LabelTarget.Caption = a_message
            This.Applied = True
            
        Case TypeOf a_target Is MSForms.Frame
            Dim FrameTarget As MSForms.Frame
            Set FrameTarget = a_target
            If This.FormatBackgroundColor Then FrameTarget.BackColor = This.ErrorBackgroundColor
            If This.FormatBorderColor Then FrameTarget.BorderColor = This.ErrorBorderColor
            If This.FormatFontBold Then FrameTarget.Font.Bold = This.ErrorFontBold
            If This.FormatForeColor Then FrameTarget.ForeColor = This.ErrorForeColor
            This.Applied = True
            
        Case TypeOf a_target Is MSForms.Image
            Dim ImageTarget As MSForms.Image
            Set ImageTarget = a_target
            If This.FormatBackgroundColor Then ImageTarget.BackColor = This.ErrorBackgroundColor
            If This.FormatBorderColor Then ImageTarget.BorderColor = This.ErrorBorderColor
            This.Applied = True
            
        Case TypeOf a_target Is Excel.Range
            Dim RangeTarget As Excel.Range
            Set RangeTarget = a_target
            If This.FormatBackgroundColor Then RangeTarget.Interior.Color = This.ErrorBackgroundColor
            If This.FormatBorderColor Then
                RangeTarget.Borders.Color = This.ErrorBorderColor
                RangeTarget.Borders.Weight = This.ErrorBorderWidth
            End If
            If This.FormatFontBold Then RangeTarget.Font.Bold = This.ErrorFontBold
            If This.FormatForeColor Then RangeTarget.Font.Color = This.ErrorForeColor
    
    End Select
    
    If This.TargetOnlyVisibleOnError And Not Control Is Nothing Then
        Control.Visible = This.Applied
    End If

End Sub


Private Sub IValidationErrorFormatter_Restore(ByVal a_target As Object)
    If Not This.Applied Then Exit Sub
    
    If TypeOf a_target Is MSForms.Control Then
        Dim Control As MSForms.Control
        Set Control = a_target
        Control.ControlTipText = VBA.vbNullString
    End If
    
    Select Case True
    
        Case TypeOf a_target Is MSForms.TextBox
            Dim TextBoxTarget As MSForms.TextBox
            Set TextBoxTarget = a_target
            If This.FormatBackgroundColor Then TextBoxTarget.BackColor = This.InitialBackgroundColor
            If This.FormatBorderColor Then TextBoxTarget.BorderColor = This.InitialBorderColor
            If This.FormatFontBold Then TextBoxTarget.Font.Bold = This.InitialFontBold
            If This.FormatForeColor Then TextBoxTarget.ForeColor = This.InitialForeColor
            
            This.Applied = False
            
        Case TypeOf a_target Is MSForms.CheckBox 'NOTE: MSForms.OptionButton also matches this interface
            Dim CheckBoxTarget As MSForms.CheckBox
            Set CheckBoxTarget = a_target
            If This.FormatFontBold Then CheckBoxTarget.Font.Bold = This.InitialFontBold
            If This.FormatForeColor Then CheckBoxTarget.ForeColor = This.InitialForeColor
            
            This.Applied = False
            
        Case TypeOf a_target Is MSForms.Label
            Dim LabelTarget As MSForms.Label
            Set LabelTarget = a_target
            If This.FormatBackgroundColor Then LabelTarget.BackColor = This.InitialBackgroundColor
            If This.FormatBorderColor Then LabelTarget.BorderColor = This.InitialBorderColor
            If This.FormatFontBold Then LabelTarget.Font.Bold = This.InitialFontBold
            If This.FormatForeColor Then LabelTarget.ForeColor = This.InitialForeColor
            
            This.Applied = False
            
        Case TypeOf a_target Is MSForms.Frame
            Dim FrameTarget As MSForms.Frame
            Set FrameTarget = a_target
            If This.FormatBackgroundColor Then FrameTarget.BackColor = This.InitialBackgroundColor
            If This.FormatBorderColor Then FrameTarget.BorderColor = This.InitialBorderColor
            If This.FormatFontBold Then FrameTarget.Font.Bold = This.InitialFontBold
            If This.FormatForeColor Then FrameTarget.ForeColor = This.InitialForeColor
        
            This.Applied = False
            
        Case TypeOf a_target Is MSForms.Image
            Dim ImageTarget As MSForms.Image
            Set ImageTarget = a_target
            If This.FormatBackgroundColor Then ImageTarget.BackColor = This.InitialBackgroundColor
            If This.FormatBorderColor Then ImageTarget.BorderColor = This.InitialBorderColor
            
            This.Applied = False
            
        Case TypeOf a_target Is Excel.Range
            Dim RangeTarget As Excel.Range
            Set RangeTarget = a_target
            If This.FormatBackgroundColor Then RangeTarget.Interior.Color = This.InitialBackgroundColor
            If This.FormatBorderColor Then
                RangeTarget.Borders.Color = This.InitialBorderColor
                RangeTarget.Borders.Weight = This.InitialBorderWidth
            End If
            If This.FormatFontBold Then RangeTarget.Font.Bold = This.InitialFontBold
            If This.FormatForeColor Then RangeTarget.Font.Color = This.InitialForeColor
            
    End Select

    If This.TargetOnlyVisibleOnError And Not Control Is Nothing Then
        Control.Visible = This.Applied
    End If

End Sub
