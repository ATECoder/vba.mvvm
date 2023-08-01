VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Factory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Construction factory for remote workbooks. </summary>
''' <remarks> Dependencies:
''' AcceptCommand.cls, AppContext.cls
''' </remarks>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit

Private Type ThisData
    ''' <summary>   Singleton instance. </summary>
    ValidationManger As ValidationManager
End Type

Private This As ThisData

''' <summary>   A public creatable singleton class for creating instances of this workbook
''' classes to share with external workbooks. </summary>
''' <remarks>
''' <see href="https://learn.Microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/set-up-vb-project-using-class"/>
''' </remarks>
Private Sub Class_Initialize()
End Sub

''' <summary>   Constructs a new instance of the <see cref="AcceptCommand"/> object to share
''' with external workbooks. </summary>
''' <returns>   A new instance of the <see cref="AcceptCommand"/> object. </returns>
Public Function NewAcceptCommand() As AcceptCommand
    
    Set NewAcceptCommand = New AcceptCommand

End Function

''' <summary>   Creates a new instance of the <see cref="cc_isr_MVVM.AcceptCommand"/> as
''' an <see cref="cc_isr_MVVM.ICommand"/> interface. </summary>
''' <param name="a_view">                [IView] a view. </param>
''' <param name="a_validationManager">   [IValidationManager] a validation manager. </param>
''' <returns>   [<see cref="cc_isr_MVVM.ICommand"/>]. </returns>
Public Function CreateAcceptCommand(ByVal a_view As cc_isr_MVVM.IView, _
            ByVal a_validationManager As cc_isr_MVVM.IValidationManager) As cc_isr_MVVM.ICommand

    Dim p_result As AcceptCommand
    Set p_result = New AcceptCommand
    p_result.Initialize a_view, a_validationManager
    Set CreateAcceptCommand = p_result

End Function

''' <summary>   Creates a new instance of the <see cref="AggregateErrorFormatter"/> object to share
''' with external workbooks. Creates a new composite formatter that aggregates/combines all
''' the specified formatters. </summary>
Public Function CreateAggregateErrorFormatter(ParamArray a_formatters() As Variant) As IValidationErrorFormatter

    Dim p_result As AggregateErrorFormatter
    Set p_result = New AggregateErrorFormatter
    p_result.Initialize a_formatters
    
    Set CreateAggregateErrorFormatter = p_result
    
End Function

''' <summary>   Creates a new instance of the <see cref="AppContext"/> object to share
''' with external workbooks. </summary>
Public Function CreateAppContext(Optional ByVal a_bindings As cc_isr_MVVM.IBindingManager, _
        Optional ByVal a_commands As cc_isr_MVVM.ICommandManager, _
        Optional ByVal a_validation As cc_isr_MVVM.IValidationManager, _
        Optional ByVal a_formatterFactory As cc_isr_MVVM.IStringFormatterFactory, _
        Optional ByVal a_adornerFactory As cc_isr_MVVM.IDynamicAdornerFactory, _
        Optional ByVal a_debugOutput As Boolean = False) As AppContext
    
    Dim p_result As cc_isr_MVVM.AppContext
    Set p_result = New cc_isr_MVVM.AppContext
    p_result.Initialize a_bindings, a_commands, a_validation, a_formatterFactory, a_adornerFactory, a_debugOutput
    Set CreateAppContext = p_result
    
End Function

''' <summary>   Constructs a new instance of the <see cref="AppContext"/> object to share
''' with external workbooks. </summary>
''' <returns>   A new instance of the <see cref="AppContext"/> object. </returns>
Public Function NewAppContext() As AppContext
    Set NewAppContext = New AppContext
End Function

''' <summary>   Creates a new instance of the <see cref="BindingManager"/> object to share
''' with external workbooks. </summary>
Public Function CreateBindingManager(ByVal a_context As IAppContext, _
        ByVal a_stringFormatterFactory As IStringFormatterFactory, _
        Optional ByVal a_factory As INotifierFactory, _
        Optional ByVal a_debugOutput As Boolean = False) As IBindingManager
    
    Dim p_result As BindingManager
    Set p_result = New BindingManager
    p_result.Initialize a_context, a_stringFormatterFactory, a_factory, a_debugOutput
    Set CreateBindingManager = p_result

End Function

''' <summary>   Constructs a new instance of the <see cref="BindingManager"/> object to share
''' with external workbooks. </summary>
''' <returns>   A new instance of the <see cref="BindingManager"/> object. </returns>
Public Function NewBindingManager() As BindingManager
    Set NewBindingManager = New BindingManager
End Function

''' <summary>   Creates a new <see cref="BindingPath"/> from the specified property path string
''' and binding context. </summary>
Public Function CreateBindingPath(ByVal a_context As Object, ByVal a_path As String) As IBindingPath
    
    Dim p_result As BindingPath
    Set p_result = New BindingPath
    p_result.Initialize a_context, a_path
    Set CreateBindingPath = p_result
    
End Function

''' <summary>   Constructs a new instance of the <see cref="BindingPath"/> object to share
''' with external workbooks. </summary>
''' <returns>   A new instance of the <see cref="BindingPath"/> object. </returns>
Public Function NewBindingPath() As BindingPath
    Set NewBindingPath = New BindingPath
End Function

''' <summary>   Constructs a new instance of the <see cref="CancelCommand"/> object to share
''' with external workbooks. </summary>
''' <returns>   A new instance of the <see cref="CancelCommand"/> object. </returns>
Public Function NewCancelCommand() As CancelCommand
    
    Set NewCancelCommand = New CancelCommand

End Function

''' <summary>   Creates a new instance of the <see cref="CancelCommand"/>. </summary>
Public Function CreateCancelCommand(ByVal a_view As ICancellable) As ICommand
    Dim p_result As CancelCommand
    Set p_result = New CancelCommand
    p_result.Initialize a_view
    Set CreateCancelCommand = p_result
End Function

''' <summary>   Creates a new instance of the <see cref="CaptionPropertyBinding"/>. </summary>
Public Function CreateCaptionPropertyBinding(ByVal a_context As IAppContext, _
        ByVal a_source As IBindingPath, ByVal a_target As Object, _
        Optional ByVal a_validator As IValueValidator, _
        Optional ByVal a_converter As IValueConverter, _
        Optional ByVal a_stringFormatter As IStringFormatter, _
        Optional ByVal a_validationAdorner As IDynamicAdorner) As IPropertyBinding
    
    Dim p_result As CaptionPropertyBinding
    Set p_result = New CaptionPropertyBinding
    p_result.Initialize a_context, a_source, a_target, a_validator, a_converter, a_stringFormatter, a_validationAdorner
    Set CreateCaptionPropertyBinding = p_result
    
End Function

''' <summary>   Creates a new instance of the <see cref="CaptionPropertyBinding"/>. </summary>
Public Function NewCaptionPropertyBinding()
    NewCaptionPropertyBinding = New CaptionPropertyBinding
End Function

''' <summary>   Creates a new instance of the <see cref="CheckBoxPropertyBinding"/>. </summary>
Public Function CreateCheckBoxPropertyBinding(ByVal a_context As IAppContext, ByVal a_source As IBindingPath, _
        ByVal a_target As MSForms.CheckBox, _
        Optional ByVal a_targetProperty As Variant, _
        Optional ByVal a_mode As BindingMode = BindingMode.TwoWayBinding, _
        Optional ByVal a_updateSource As BindingUpdateSourceTrigger = OnPropertyChanged, _
        Optional ByVal a_validator As IValueValidator, _
        Optional ByVal a_converter As IValueConverter, _
        Optional ByVal a_stringFormatter As IStringFormatter, _
        Optional ByVal a_validationAdorner As IDynamicAdorner) As IPropertyBinding
    
    Dim p_result As CheckBoxPropertyBinding
    Set p_result = New CheckBoxPropertyBinding
    p_result.Initialize a_context, a_source, a_target, a_targetProperty, a_mode, a_updateSource, _
                      a_validator, a_converter, a_stringFormatter, a_validationAdorner
    Set CreateCheckBoxPropertyBinding = p_result
    
End Function

''' <summary>   Creates a new instance of the <see cref="ComboBoxPropertyBinding"/>. </summary>
Public Function CreateComboBoxPropertyBinding(ByVal a_context As IAppContext, ByVal a_source As IBindingPath, _
        ByVal a_target As MSForms.ComboBox, _
        Optional ByVal a_targetProperty As Variant, _
        Optional ByVal a_mode As BindingMode = BindingMode.TwoWayBinding, _
        Optional ByVal a_updateSource As BindingUpdateSourceTrigger = OnPropertyChanged, _
        Optional ByVal a_validator As IValueValidator, _
        Optional ByVal a_converter As IValueConverter, _
        Optional ByVal a_stringFormatter As IStringFormatter, _
        Optional ByVal a_validationAdorner As IDynamicAdorner) As IPropertyBinding
    
    Dim p_result As ComboBoxPropertyBinding
    Set p_result = New ComboBoxPropertyBinding
    p_result.Initialize a_context, a_source, a_target, a_targetProperty, a_mode, a_updateSource, _
                      a_validator, a_converter, a_stringFormatter, a_validationAdorner
    Set CreateComboBoxPropertyBinding = p_result
    
End Function

''' <summary>   Constructs a new instance of the <see cref="CommandBinding"/> object to share
''' with external workbooks. </summary>
''' <returns>   A new instance of the <see cref="CommandBinding"/> object. </returns>
Public Function NewCommandBinding() As CommandBinding
    Set NewCommandBinding = New CommandBinding
End Function

''' <summary>   Constructs a new instance of the <see cref="CommandManager"/> object to share
''' with external workbooks. </summary>
''' <returns>   A new instance of the <see cref="CommandManager"/> object. </returns>
Public Function NewCommandManager() As CommandManager
    Set NewCommandManager = New CommandManager
End Function

''' <summary>   Constructs a new instance of the <see cref="EventCommand"/> object to share
''' with external workbooks. </summary>
''' <returns>   A new instance of the <see cref="EventCommand"/> object. </returns>
Public Function NewEventCommand() As EventCommand
    Set NewEventCommand = New EventCommand
End Function

''' <summary>   Creates a new instance of the <see cref="ContainerLayout"/> object to share
''' with external workbooks. </summary>
Public Function CreateContainerLayout(ByVal a_parent As MSForms.Controls, _
        Optional ByVal a_direction As LayoutDirection, _
        Optional ByVal a_top As Long = 0, _
        Optional ByVal a_left As Long = 0, _
        Optional ByVal a_margin As Long = 3) As IContainerLayout
    Dim p_result As ContainerLayout
    Set p_result = New ContainerLayout
    p_result.Initialize a_parent, a_direction, a_top, a_left, a_margin
    Set CreateContainerLayout = p_result
End Function

''' <summary>   Constructs a new instance of the <see cref="ContainerLayout"/> object to share
''' with external workbooks. </summary>
''' <returns>   A new instance of the <see cref="ContainerLayout"/> object. </returns>
Public Function NewContainerLayout() As ContainerLayout
    Set NewContainerLayout = New ContainerLayout
End Function


''' <summary>   Constructs a new instance of the <see cref="DecimalKeyValidator"/> object to share
''' with external workbooks. </summary>
''' <returns>   A new instance of the <see cref="DecimalKeyValidator"/> object. </returns>
Public Function NewDecimalKeyValidator() As DecimalKeyValidator
    Set NewDecimalKeyValidator = New DecimalKeyValidator
End Function

''' <summary>   Creates a new instance of the <see cref="DynamicControlPosition"/> object to share
''' with external workbooks. </summary>
Public Function CreateDynamicControlPosition(ByVal a_position As RelativePosition, _
        ByVal a_leftMargin As Double, _
        Optional ByVal a_topMargin As Variant, _
        Optional ByVal a_rightMargin As Variant, _
        Optional ByVal a_bottomMargin As Variant) As IDynamicPosition
    
    Dim p_result As cc_isr_MVVM.DynamicControlPosition
    Set p_result = New cc_isr_MVVM.DynamicControlPosition
    p_result.Initialize a_position, a_leftMargin, a_topMargin, a_rightMargin, a_bottomMargin
    Set CreateDynamicControlPosition = p_result
    
End Function

''' <summary>   Creates a new instance of the <see cref="DynamicControls"/> object to share
''' with external workbooks. </summary>
Public Function CreateDynamicControls(ByVal a_context As cc_isr_MVVM.IAppContext, _
        ByVal a_container As IContainerLayout) As IDynamicControlBuilder
    
    Dim p_result As DynamicControls
    Set p_result = New DynamicControls
    p_result.Initialize a_context, a_container
    Set CreateDynamicControls = p_result
    
End Function

''' <summary>   Creates a new instance of the <see cref="Layout"/> object to share
''' with external workbooks. </summary>
Public Function CreateLayout(ByVal a_object As Object, Optional ByVal a_newWidthPercent As Long = 50, _
        Optional ByVal a_newHeightPercent As Long = 50) As Layout
    
    Dim p_result As Layout
    Set p_result = New Layout
    p_result.Initialize a_object, a_newWidthPercent, a_newHeightPercent
    Set CreateLayout = p_result
    
End Function

''' <summary>   Creates a new instance of the <see cref="LayoutPanel"/> object to share
''' with external workbooks. </summary>
Public Function CreateLayoutPanel(ByVal a_parentForm As MSForms.UserForm, _
        Optional ByVal a_parent As MSForms.Controls, _
        Optional ByVal a_height As Single, _
        Optional ByVal a_width As Single) As LayoutPanel
    
    Dim p_result As LayoutPanel
    Set p_result = New LayoutPanel
    p_result.Initialize a_parentForm, a_parent, a_height, a_width
    Set CreateLayoutPanel = p_result
    
End Function

''' <summary>   Creates a new instance of the <see cref="ListBoxPropertyBinding"/> object to share
''' with external workbooks. </summary>
Public Function CreateListBoxPropertyBinding(ByVal a_context As IAppContext, a_source As IBindingPath, _
        ByVal a_target As MSForms.ListBox, _
        Optional ByVal a_targetProperty As Variant, _
        Optional ByVal a_mode As BindingMode = BindingMode.TwoWayBinding, _
        Optional ByVal a_updateSource As BindingUpdateSourceTrigger = OnPropertyChanged, _
        Optional ByVal a_validator As IValueValidator, _
        Optional ByVal a_converter As IValueConverter, _
        Optional ByVal a_stringFormatter As IStringFormatter, _
        Optional ByVal a_validationAdorner As IDynamicAdorner) As IPropertyBinding
    
    Dim p_result As ListBoxPropertyBinding
    Set p_result = New ListBoxPropertyBinding
    p_result.Initialize a_context, a_source, a_target, a_targetProperty, a_mode, _
                      a_updateSource, a_validator, a_converter, a_stringFormatter, a_validationAdorner
    Set CreateListBoxPropertyBinding = p_result
    
End Function

''' <summary>   Creates a new instance of the <see cref="MultiPagePropertyBinding"/> object to share
''' with external workbooks. </summary>
Public Function CreateMultiPagePropertyBinding(ByVal a_context As IAppContext, a_source As IBindingPath, _
        ByVal a_target As MSForms.MultiPage, _
        Optional ByVal a_targetProperty As String = "Value", _
        Optional ByVal a_mode As BindingMode = BindingMode.TwoWayBinding, _
        Optional ByVal a_updateSource As BindingUpdateSourceTrigger = OnPropertyChanged, _
        Optional ByVal a_validator As IValueValidator, _
        Optional ByVal a_converter As IValueConverter, _
        Optional ByVal a_stringFormatter As IStringFormatter, _
        Optional ByVal a_validationAdorner As IDynamicAdorner) As IPropertyBinding
        
    Dim p_result As MultiPagePropertyBinding
    Set p_result = New MultiPagePropertyBinding
    p_result.Initialize a_context, a_source, a_target, a_targetProperty, a_mode, _
                      a_updateSource, a_validator, a_converter, a_stringFormatter, a_validationAdorner
    Set CreateMultiPagePropertyBinding = p_result
    
End Function

''' <summary>   Creates a new instance of the <see cref="OneWayPropertyBinding"/> object to share
''' with external workbooks. </summary>
Public Function CreateOneWayPropertyBinding(ByVal a_context As IAppContext, ByVal a_source As IBindingPath, _
        ByVal a_target As IBindingPath, _
        Optional ByVal a_validator As IValueValidator, _
        Optional ByVal a_converter As IValueConverter, _
        Optional ByVal a_stringFormatter As IStringFormatter, _
        Optional ByVal a_validationAdorner As IDynamicAdorner) As IPropertyBinding
    
    Dim p_result As OneWayPropertyBinding
    Set p_result = New OneWayPropertyBinding
    p_result.Initialize a_context, a_source, a_target, a_validator, a_converter, a_stringFormatter, a_validationAdorner
    Set CreateOneWayPropertyBinding = p_result
    
End Function

''' <summary>   Creates a new instance of the <see cref="OptionButtonPropertyBinding"/> object to share
''' with external workbooks. </summary>
Public Function CreateOptionButtonPropertyBinding(ByVal a_context As IAppContext, ByVal a_source As IBindingPath, _
        ByVal a_target As MSForms.OptionButton, _
        Optional ByVal a_targetProperty As Variant, _
        Optional ByVal a_mode As BindingMode = BindingMode.TwoWayBinding, _
        Optional ByVal a_updateSource As BindingUpdateSourceTrigger = OnPropertyChanged, _
        Optional ByVal a_validator As IValueValidator, _
        Optional ByVal a_converter As IValueConverter, _
        Optional ByVal a_stringFormatter As IStringFormatter, _
        Optional ByVal a_validationAdorner As IDynamicAdorner) As IPropertyBinding
    
    Dim p_result As OptionButtonPropertyBinding
    Set p_result = New OptionButtonPropertyBinding
    p_result.Initialize a_context, a_source, a_target, a_targetProperty, a_mode, _
                      a_updateSource, a_validator, a_converter, a_stringFormatter, a_validationAdorner
    Set CreateOptionButtonPropertyBinding = p_result
    
End Function

''' <summary>   Creates a new instance of the <see cref="PropertyBindingBase"/> object to share
''' with external workbooks. </summary>
Public Function CreatePropertyBindingBase(ByVal a_context As IAppContext, ByVal a_source As IBindingPath, _
        ByVal a_targetContext As Object, _
        Optional ByVal a_targetPropertyPath As String = "Value", _
        Optional ByVal a_mode As BindingMode = BindingMode.TwoWayBinding, _
        Optional ByVal a_updateSource As BindingUpdateSourceTrigger = BindingUpdateSourceTrigger.OnExit, _
        Optional ByVal a_converter As IValueConverter, _
        Optional ByVal a_stringFormatter As IStringFormatter, _
        Optional ByVal a_validator As IValueValidator, _
        Optional ByVal a_validationAdorner As IDynamicAdorner, _
        Optional ByVal a_cancelExitOnValidationError As Boolean = True) As IPropertyBinding

    Dim p_result As PropertyBindingBase
    Set p_result = New PropertyBindingBase
    p_result.Initialize a_context, a_source, a_targetContext, a_targetPropertyPath, a_mode, a_updateSource, _
            a_converter, a_stringFormatter, a_validator, a_validationAdorner, a_cancelExitOnValidationError
    Set CreatePropertyBindingBase = p_result
    
End Function


''' <summary>   Constructs a new instance of the <see cref="PropertyChangeNotifierBase"/> object to share
''' with external workbooks. </summary>
''' <returns>   A new instance of the <see cref="PropertyChangeNotifierBase"/> object. </returns>
Public Function NewPropertyChangeNotifierBase() As PropertyChangeNotifierBase
    Set NewPropertyChangeNotifierBase = New PropertyChangeNotifierBase
End Function

''' <summary>   Constructs a new instance of the <see cref="RequiredStringValidator"/> object to share
''' with external workbooks. </summary>
''' <returns>   A new instance of the <see cref="RequiredStringValidator"/> object. </returns>
Public Function NewRequiredStringValidator() As RequiredStringValidator
    Set NewRequiredStringValidator = New RequiredStringValidator
End Function

''' <summary>   Creates a new instance of the <see cref="ScrollBarPropertyBinding"/> object to share
''' with external workbooks. </summary>
Public Function CreateScrollBarPropertyBinding(ByVal a_context As IAppContext, ByVal a_source As IBindingPath, _
        ByVal a_target As MSForms.ScrollBar, _
        Optional ByVal a_targetProperty As String = "Value", _
        Optional ByVal a_mode As BindingMode = BindingMode.TwoWayBinding, _
        Optional ByVal a_updateSource As BindingUpdateSourceTrigger = OnPropertyChanged, _
        Optional ByVal a_validator As IValueValidator, _
        Optional ByVal a_converter As IValueConverter, _
        Optional ByVal a_stringFormatter As IStringFormatter, _
        Optional ByVal a_validationAdorner As IDynamicAdorner) As IPropertyBinding
        
    Dim p_result As ScrollBarPropertyBinding
    Set p_result = New ScrollBarPropertyBinding
    p_result.Initialize a_context, a_source, a_target, a_targetProperty, a_mode, _
                      a_updateSource, a_validator, a_converter, a_stringFormatter, a_validationAdorner
    Set CreateScrollBarPropertyBinding = p_result
    
End Function

''' <summary>   Creates a new instance of the <see cref="SpinButtonPropertyBinding"/> object to share
''' with external workbooks. </summary>
Public Function CreateSpinButtonPropertyBinding(ByVal a_context As IAppContext, ByVal a_source As IBindingPath, _
        ByVal a_target As MSForms.SpinButton, _
        Optional ByVal a_targetProperty As String = "Value", _
        Optional ByVal a_mode As BindingMode = BindingMode.OneWayBinding, _
        Optional ByVal a_updateSource As BindingUpdateSourceTrigger = OnPropertyChanged, _
        Optional ByVal a_validator As IValueValidator, _
        Optional ByVal a_converter As IValueConverter, _
        Optional ByVal a_stringFormatter As IStringFormatter, _
        Optional ByVal a_validationAdorner As IDynamicAdorner) As IPropertyBinding
        
    Dim p_result As SpinButtonPropertyBinding
    Set p_result = New SpinButtonPropertyBinding
    p_result.Initialize a_context, a_source, a_target, a_targetProperty, a_mode, _
                      a_updateSource, a_validator, a_converter, a_stringFormatter, a_validationAdorner
    Set CreateSpinButtonPropertyBinding = p_result
    
End Function

''' <summary>   Creates a new instance of the <see cref="StringFormatter"/> object to share
''' with external workbooks. </summary>
Public Function CreateStringFormatter(ByVal a_formatString As String) As IStringFormatter
    
    Dim p_result As StringFormatter
    Set p_result = New StringFormatter
    p_result.Initialize a_formatString
    Set CreateStringFormatter = p_result

End Function

''' <summary>   Constructs a new instance of the <see cref="StringFormatterFactory"/> object to share
''' with external workbooks. </summary>
''' <returns>   A new instance of the <see cref="StringFormatterFactory"/> object. </returns>
Public Function NewStringFormatterFactory() As StringFormatterFactory
    
    Set NewStringFormatterFactory = New StringFormatterFactory

End Function

''' <summary>   Creates a new instance of the <see cref="TabStripPropertyBinding"/> object to share
''' with external workbooks. </summary>
Public Function CreateTabStripPropertyBinding(ByVal a_context As IAppContext, a_source As IBindingPath, _
        ByVal a_target As MSForms.TabStrip, _
        Optional ByVal a_targetProperty As String = "Value", _
        Optional ByVal a_mode As BindingMode = BindingMode.TwoWayBinding, _
        Optional ByVal a_updateSource As BindingUpdateSourceTrigger = OnPropertyChanged, _
        Optional ByVal a_validator As IValueValidator, _
        Optional ByVal a_converter As IValueConverter, _
        Optional ByVal a_stringFormatter As IStringFormatter, _
        Optional ByVal a_validationAdorner As IDynamicAdorner) As IPropertyBinding

    Dim p_result As TabStripPropertyBinding
    Set p_result = New TabStripPropertyBinding
    p_result.Initialize a_context, a_source, a_target, a_targetProperty, a_mode, _
                      a_updateSource, a_validator, a_converter, a_stringFormatter, a_validationAdorner
    Set CreateTabStripPropertyBinding = p_result
    
End Function

''' <summary>   Creates a new instance of the <see cref="TextBoxPropertyBinding"/> object to share
''' with external workbooks. </summary>
Public Function CreateTextBoxPropertyBinding(ByVal a_context As IAppContext, ByVal a_source As IBindingPath, _
        ByVal a_target As Object, _
        Optional ByVal a_targetProperty As String = "Text", _
        Optional ByVal a_mode As BindingMode = BindingMode.TwoWayBinding, _
        Optional ByVal a_updateSource As BindingUpdateSourceTrigger = OnExit, _
        Optional ByVal a_validator As IValueValidator, _
        Optional ByVal a_converter As IValueConverter, _
        Optional ByVal a_stringFormatter As IStringFormatter, _
        Optional ByVal a_validationAdorner As IDynamicAdorner) As IPropertyBinding
    
    Dim p_result As TextBoxPropertyBinding
    Set p_result = New TextBoxPropertyBinding
    p_result.Initialize a_context, a_source, a_target, a_targetProperty, a_mode, _
                      a_updateSource, a_validator, a_converter, a_stringFormatter, a_validationAdorner
    Set CreateTextBoxPropertyBinding = p_result
    
End Function

''' <summary>   Creates a new instance of the <see cref="ValidationError"/> object to share
''' with external workbooks. </summary>
Public Function CreateValidationError(ByVal a_binding As IPropertyBinding, ByVal a_message As String) As IValidationError
    
    Dim p_result As ValidationError
    Set p_result = New ValidationError
    Set p_result.Binding = a_binding
    p_result.Message = a_message
    Set CreateValidationError = p_result

End Function

''' <summary>   Creates a new instance of the <see cref="ValidationErrorAdorner"/> object to share
''' with external workbooks. </summary>
Public Function CreateValidationErrorAdorner(ByVal a_target As MSForms.Control, _
        Optional ByVal a_targetFormatter As cc_isr_MVVM.IValidationErrorFormatter = Nothing, _
        Optional ByVal a_adornerLabel As Variant, _
        Optional ByVal a_labelFormatter As cc_isr_MVVM.IValidationErrorFormatter = Nothing, _
        Optional ByVal a_labelPosition As cc_isr_MVVM.IDynamicPosition = Nothing, _
        Optional ByVal a_adornerIcon As Variant, _
        Optional ByVal a_iconFormatter As cc_isr_MVVM.IValidationErrorFormatter = Nothing, _
        Optional ByVal a_iconPosition As cc_isr_MVVM.IDynamicPosition = Nothing) As IDynamicAdorner
    
    Dim p_result As ValidationErrorAdorner
    Set p_result = New ValidationErrorAdorner
    p_result.Initialize a_target, a_targetFormatter, a_adornerLabel, a_labelFormatter, a_labelPosition, _
                      a_adornerIcon, a_iconFormatter, a_iconPosition
    Set CreateValidationErrorAdorner = p_result
    
End Function


''' <summary>   Constructs a new instance of the <see cref="ValidationErrorFormatter"/> object to share
''' with external workbooks. </summary>
Public Function NewValidationErrorFormatter() As ValidationErrorFormatter
    
    Dim p_result As ValidationErrorFormatter
    Set p_result = New ValidationErrorFormatter
    Set NewValidationErrorFormatter = p_result

End Function

''' <summary>   Creates a new instance of the <see cref="ValidationErrors"/> object to share
''' with external workbooks.Creates a new collection of IValidationError objects for the specified
''' binding context (ViewModel). </summary>
Public Function CreateValidationErrors(ByVal a_bindingContext As Object) As ValidationErrors
    
    Dim p_result As ValidationErrors
    Set p_result = New ValidationErrors
    p_result.Initialize a_bindingContext
    Set CreateValidationErrors = p_result

End Function

''' <summary>   Creates a new instance of the <see cref="ValidationManager"/> object to share
''' with external workbooks. </summary>
Public Function CreateValidationManager(Optional ByVal a_errorAdornerFactory As IDynamicAdornerFactory) As ValidationManager
    Dim p_result As ValidationManager
    Set p_result = New ValidationManager
    p_result.Initialize a_errorAdornerFactory
    Set CreateValidationManager = p_result
End Function

''' <summary>   Returns a singleton instance of the <see cref="ValidationManager"/> object to share
''' with external workbooks. </summary>
Public Function ValidationManager() As ValidationManager
    If This.ValidationManger Is Nothing Then
        This.ValidationManger = Me.CreateValidationManager
    End If
    Set ValidationManager = This.ValidationManger
End Function

''' <summary>   Creates a new instance of the <see cref="ValueRangeValidator"/> object to share
''' with external workbooks. </summary>
Public Function CreateValueRangeValidator(ByVal a_min As Variant, ByVal a_max As Variant) As IValueValidator
    Dim p_result As ValueRangeValidator
    Set p_result = New ValueRangeValidator
    p_result.Min = a_min
    p_result.Max = a_max
    Set CreateValueRangeValidator = p_result
End Function

''' <summary>   Creates a new instance of the <see cref="WorksheetAreaPropertyBinding"/> object to share
''' with external workbooks. </summary>
Public Function CreateWorksheetAreaPropertyBinding(ByVal a_context As IAppContext, ByVal a_source As IBindingPath, ByVal a_target As Excel.Range, _
        Optional ByVal a_targetProperty As String = "Value", _
        Optional ByVal a_mode As BindingMode = BindingMode.TwoWayBinding, _
        Optional ByVal a_validator As IValueValidator, _
        Optional ByVal a_converter As IValueConverter, _
        Optional ByVal a_stringFormatter As IStringFormatter, _
        Optional ByVal a_validationAdorner As IDynamicAdorner) As IPropertyBinding
    
    Dim p_result As WorksheetAreaPropertyBinding
    Set p_result = New WorksheetAreaPropertyBinding
    p_result.Initialize a_context, a_source, a_target, a_targetProperty, a_mode, a_validator, _
                      a_converter, a_stringFormatter, a_validationAdorner
    Set CreateWorksheetAreaPropertyBinding = p_result
    
End Function

''' <summary>   Creates a new instance of the <see cref="WorksheetCellPropertyBinding"/> object to share
''' with external workbooks. </summary>
Public Function CreateWorksheetCellPropertyBinding(ByVal a_context As IAppContext, ByVal a_source As IBindingPath, _
        ByVal a_target As Excel.Range, _
        Optional ByVal a_targetProperty As String = "Value", _
        Optional ByVal a_mode As BindingMode = BindingMode.TwoWayBinding, _
        Optional ByVal a_validator As IValueValidator, _
        Optional ByVal a_converter As IValueConverter, _
        Optional ByVal a_stringFormatter As IStringFormatter, _
        Optional ByVal a_validationAdorner As IDynamicAdorner) As IPropertyBinding
    
    Dim p_result As WorksheetCellPropertyBinding
    Set p_result = New WorksheetCellPropertyBinding
    p_result.Initialize a_context, a_source, a_target, a_targetProperty, a_mode, a_validator, _
                      a_converter, a_stringFormatter, a_validationAdorner
    Set CreateWorksheetCellPropertyBinding = p_result
    
End Function

