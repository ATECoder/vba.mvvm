VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDynamicControlBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Builds dynamic MSForms UI components from a binding source."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Builds dynamic MSForms UI components from a binding source. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit

''' <summary>   Creates a multiline MSForms.TextBox control for the spercified String property binding path. </summary>
Public Function TextAreaFor(ByVal a_sourceValue As IBindingPath, Optional ByVal a_converter As IValueConverter, _
        Optional ByVal a_validator As IValueValidator, Optional ByVal a_errorAdorner As IDynamicAdorner, _
        Optional ByVal a_titleSource As Variant) As MSForms.TextBox
End Function

''' <summary>   Creates a MSForms.TextBox control for the specified String property binding path. </summary>
Public Function TextBoxFor(ByVal a_sourceValue As IBindingPath, Optional ByVal a_formatString As String, _
        Optional ByVal a_converter As IValueConverter, Optional ByVal a_validator As IValueValidator, _
        Optional ByVal a_errorAdorner As IDynamicAdorner, Optional ByVal a_titleSource As Variant) As MSForms.TextBox
End Function

''' <summary>   Creates a MSForms.Label control for the specified Caption string or String property binding path. </summary>
Public Function LabelFor(ByVal a_sourceCaption As Variant, Optional ByVal a_formatString As String, _
        Optional ByVal a_converter As IValueConverter) As MSForms.Label
End Function

''' <summary>   Creates a MSForms.ComboBox control for the specified Value property binding path; a_sourceItems should be an array property. </summary>
Public Function ComboBoxFor(ByVal a_sourceValue As IBindingPath, ByVal a_sourceItems As IBindingPath, _
        Optional ByVal a_formatString As String, Optional ByVal a_converter As IValueConverter, _
        Optional ByVal a_validator As IValueValidator, Optional ByVal a_errorAdorner As IDynamicAdorner, _
        Optional ByVal a_titleSource As Variant) As MSForms.ComboBox
End Function

''' <summary>   Creates a MSForms.ListBox control for the specified Value property binding path; a_sourceItems should be an array property. </summary>
Public Function ListBoxFor(ByVal a_sourceValue As IBindingPath, ByVal a_sourceItems As IBindingPath, _
        Optional ByVal a_titleSource As Variant) As MSForms.ListBox
End Function

''' <summary>   Creates a MSForms.OptionButton control for the specified Value (Boolean) property binding path; uses the specified Caption string or String property binding path for the control's Caption. </summary>
Public Function OptionButtonFor(ByVal a_sourceValue As IBindingPath, ByVal a_sourceCaption As Variant) As MSForms.OptionButton
Attribute OptionButtonFor.VB_Description = "Creates a MSForms.OptionButton control for the specified Value (Boolean) property binding path; uses the specified Caption string or String property binding path for the control's Caption."
End Function

''' <summary>   Creates a MSForms.CheckBoxButton control for the specified Value (Boolean) property binding path; uses the specified Caption string or String property binding path for the control's Caption. </summary>
Public Function CheckBoxFor(ByVal a_sourceValue As IBindingPath, ByVal a_sourceCaption As Variant) As MSForms.CheckBox
Attribute CheckBoxFor.VB_Description = "Creates a MSForms.CheckBoxButton control for the specified Value (Boolean) property binding path; uses the specified Caption string or String property binding path for the control's Caption."
End Function

''' <summary>   Creates a MSForms.CommandButton control for the specified ICommand, using the specified ViewModel context and Caption string or String property binding path. </summary>
Public Function CommandButtonFor(ByVal a_command As ICommand, ByVal a_bindingContext As Object, _
        ByVal a_sourceCaption As Variant) As MSForms.CommandButton
End Function
