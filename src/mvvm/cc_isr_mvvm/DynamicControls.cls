VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DynamicControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Builds dynamic MSForms UI components from a binding source."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Builds dynamic MSForms UI components from a binding source. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit
Implements IDynamicControlBuilder

Private Type TState
    Context As cc_isr_MVVM.IAppContext
    Container As cc_isr_MVVM.IContainerLayout
End Type

Private This As TState

''' <summary>   Initialize and returns an <see cref="IDynamicControlBuilder"/> representation of
'''             this <see cref="DynamicControls"/> object. </summary>
''' <returns>   [<see cref="IDynamicControlBuilder"/>]. </returns>
Public Function Initialize(ByVal a_context As cc_isr_MVVM.IAppContext, _
        ByVal a_container As IContainerLayout) As IDynamicControlBuilder
    
    Set Me.Context = a_context
    Set Me.Container = a_container
    
    Set Initialize = Me

End Function

Public Property Get Container() As cc_isr_MVVM.IContainerLayout
    Set Container = This.Container
End Property

Public Property Set Container(ByVal a_value As cc_isr_MVVM.IContainerLayout)
    Set This.Container = a_value
End Property

Public Property Get Context() As cc_isr_MVVM.IAppContext
    Set Context = This.Context
End Property

Public Property Set Context(ByVal a_value As cc_isr_MVVM.IAppContext)
    Set This.Context = a_value
End Property

Private Sub BindCaptionSource(ByVal a_source As Variant, ByVal a_target As Object, _
        Optional ByVal a_formatString As String, Optional ByVal a_converter As IValueConverter)
        
    If VBA.VarType(a_source) = VBA.VbVarType.vbString Then
        
        'late-bound
        a_target.Caption = a_source
        
    ElseIf TypeOf a_source Is cc_isr_MVVM.IBindingPath Then
        
        Dim p_sourcePath As IBindingPath
        Set p_sourcePath = a_source
        This.Context.Bindings.BindPropertyPath p_sourcePath.Context, p_sourcePath.Path, a_target, _
            a_converter:=a_converter, _
            a_formatString:=a_formatString
        
    Else
        cc_isr_Core.GuardClauses.GuardExpression True, VBA.Information.TypeName(Me), "Unsupported Source type.", _
            cc_isr_Core.UserDefinedErrors.InvalidArgumentError
    End If
    
End Sub

Private Function CreateLabel(ByVal a_sourceCaption As Variant, Optional ByVal a_formatString As String, _
        Optional ByVal a_converter As IValueConverter, Optional ByVal a_applyMargin As Boolean = True) As MSForms.Label
        
    Dim p_result As MSForms.Label
    Set p_result = This.Container.Add(BindingDefaults.LabelProgId, a_applyMargin:=a_applyMargin)
    p_result.AutoSize = True
    p_result.WordWrap = False
    BindCaptionSource a_sourceCaption, p_result, a_formatString, a_converter
    Set CreateLabel = p_result
    
End Function

Private Function ErrorAdornerOrDefault(ByVal a_validator As IValueValidator, ByVal a_errorAdorner As IDynamicAdorner, _
        ByVal a_target As Object) As IDynamicAdorner
        
    Dim p_result As IDynamicAdorner
    If Not a_validator Is Nothing Then
        If a_errorAdorner Is Nothing Then
            Set p_result = This.Context.Validation.AdornerFactory.Create(a_target)
        Else
            Set p_result = a_errorAdorner
        End If
    End If
    Set ErrorAdornerOrDefault = p_result
End Function

Private Function IDynamicControlBuilder_CheckBoxFor(ByVal a_sourceValue As IBindingPath, _
        ByVal a_sourceCaption As Variant) As MSForms.CheckBox
    
    Dim p_result As MSForms.CheckBox
    Set p_result = This.Container.Add(BindingDefaults.CheckBoxProgId)
    
    BindCaptionSource a_sourceCaption, p_result
    This.Context.Bindings.BindPropertyPath a_sourceValue.Context, a_sourceValue.Path, p_result
    
    Set IDynamicControlBuilder_CheckBoxFor = p_result
    
End Function

Private Function IDynamicControlBuilder_ComboBoxFor(ByVal a_sourceValue As IBindingPath, ByVal a_sourceItems As IBindingPath, Optional ByVal a_formatString As String, Optional ByVal a_converter As IValueConverter, Optional ByVal a_validator As IValueValidator, Optional ByVal a_errorAdorner As IDynamicAdorner, Optional ByVal a_titleSource As Variant) As MSForms.ComboBox
    
    If Not VBA.IsEmpty(a_titleSource) Then
        CreateLabel a_titleSource, a_applyMargin:=False
    End If
    
    Dim p_result As MSForms.ComboBox
    Set p_result = This.Container.Add(BindingDefaults.ComboBoxProgId)
        
    This.Context.Bindings.BindPropertyPath a_sourceValue.Context, a_sourceValue.Path, p_result, _
        a_formatString:=a_formatString, _
        a_converter:=a_converter, _
        a_validator:=a_validator, _
        a_validationAdorner:=ErrorAdornerOrDefault(a_validator, a_errorAdorner, p_result)

    This.Context.Bindings.BindPropertyPath a_sourceItems.Context, a_sourceItems.Path, p_result
    
    Set IDynamicControlBuilder_ComboBoxFor = p_result
    
End Function

Private Function IDynamicControlBuilder_CommandButtonFor(ByVal a_command As ICommand, ByVal a_bindingContext As Object, ByVal a_sourceCaption As Variant) As MSForms.CommandButton
    
    Dim p_result As MSForms.CommandButton
    Set p_result = This.Container.Add(BindingDefaults.CommandButtonProgId)
    
    Dim Button As MSForms.Control
    Set Button = p_result
    
    If TypeOf a_command Is CancelCommand Then
        Button.Cancel = True
    ElseIf TypeOf a_command Is AcceptCommand Then
        Button.Default = True
    End If
    
    BindCaptionSource a_sourceCaption, p_result
    This.Context.Commands.BindCommand a_bindingContext, p_result, a_command
    
    Set IDynamicControlBuilder_CommandButtonFor = p_result
    
End Function

Private Function IDynamicControlBuilder_LabelFor(ByVal a_sourceCaption As Variant, Optional ByVal a_formatString As String, _
        Optional ByVal a_converter As IValueConverter) As MSForms.Label
    Set IDynamicControlBuilder_LabelFor = CreateLabel(a_sourceCaption, a_formatString, a_converter)
End Function

Private Function IDynamicControlBuilder_ListBoxFor(ByVal a_sourceValue As IBindingPath, ByVal a_sourceItems As IBindingPath, _
        Optional ByVal a_titleSource As Variant) As MSForms.ListBox
    
    If Not VBA.IsEmpty(a_titleSource) Then
        CreateLabel a_titleSource, a_applyMargin:=False
    End If
    
    Dim p_result As MSForms.ListBox
    Set p_result = This.Container.Add(BindingDefaults.ListBoxProgId)
    
    This.Context.Bindings.BindPropertyPath a_sourceValue.Context, a_sourceValue.Path, p_result
    This.Context.Bindings.BindPropertyPath a_sourceItems.Context, a_sourceItems.Path, p_result
    
    Set IDynamicControlBuilder_ListBoxFor = p_result
    
End Function

Private Function IDynamicControlBuilder_OptionButtonFor(ByVal a_sourceValue As IBindingPath, _
        ByVal a_sourceCaption As Variant) As MSForms.OptionButton
    
    Dim p_result As MSForms.OptionButton
    Set p_result = This.Container.Add(BindingDefaults.OptionButtonProgId)
    
    BindCaptionSource a_sourceCaption, p_result
    This.Context.Bindings.BindPropertyPath a_sourceValue.Context, a_sourceValue.Path, p_result
    
    Set IDynamicControlBuilder_OptionButtonFor = p_result
    
End Function

Private Function IDynamicControlBuilder_TextAreaFor(ByVal a_sourceValue As IBindingPath, _
        Optional ByVal a_converter As IValueConverter, Optional ByVal a_validator As IValueValidator, _
        Optional ByVal a_errorAdorner As IDynamicAdorner, Optional ByVal a_titleSource As Variant) As MSForms.TextBox
    
    If Not VBA.IsEmpty(a_titleSource) Then
        CreateLabel a_titleSource, a_applyMargin:=False
    End If
    
    Dim p_result As MSForms.TextBox
    Set p_result = IDynamicControlBuilder_TextBoxFor(a_sourceValue, a_converter:=a_converter, a_validator:=a_validator, a_errorAdorner:=a_errorAdorner)
    
    p_result.MultiLine = True
    p_result.WordWrap = True
    p_result.ScrollBars = MSForms.fmScrollBars.fmScrollBarsVertical
    
    Set IDynamicControlBuilder_TextAreaFor = p_result
    
End Function

Private Function IDynamicControlBuilder_TextBoxFor(ByVal a_sourceValue As IBindingPath, _
        Optional ByVal a_formatString As String, Optional ByVal a_converter As IValueConverter, _
        Optional ByVal a_validator As IValueValidator, Optional ByVal a_errorAdorner As IDynamicAdorner, Optional ByVal a_titleSource As Variant) As MSForms.TextBox
    
    If Not VBA.IsEmpty(a_titleSource) Then
        CreateLabel a_titleSource, a_applyMargin:=False
    End If
    
    Dim p_result As MSForms.TextBox
    Set p_result = This.Container.Add(BindingDefaults.TextBoxProgId)
    
    Dim p_trigger As BindingUpdateSourceTrigger
'    ' VF: Logic error?: would always be zero and that means never // missing: if not validator is nothing then ??? trigger not picked up from validator
'    'is updated later. Why not set here and skip update later?
'    If a_validator Is Nothing Then
'        p_trigger = a_validator.Trigger '= 0 that is never
''    '    ' VF:
''    '    Else ' set value of validator
''    '        p_trigger = a_validator.Trigger
''    '    End If
'    End If
    ' VF: simplifies to unconditional
    'but causes DecimalValidator::IValueValidator_Trigger() to be called twice
    p_trigger = a_validator.Trigger
    
    
    This.Context.Bindings.BindPropertyPath a_sourceValue.Context, a_sourceValue.Path, p_result, _
        a_formatString:=a_formatString, _
        a_converter:=a_converter, _
        a_validator:=a_validator, _
        a_validationAdorner:=ErrorAdornerOrDefault(a_validator, a_errorAdorner, p_result), _
        a_updateTrigger:=p_trigger
    
    Set IDynamicControlBuilder_TextBoxFor = p_result
    
End Function

