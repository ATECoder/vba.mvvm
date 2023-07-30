VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Constructor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Provides constructors for remote workbooks. </summary>
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

''' <summary>   Creates a new instance of the <see cref="AcceptCommand"/> class to share
''' with external workbooks. </summary>
''' <returns>   A new instance of the <see cref="AcceptCommand"/> class. </returns>
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

    Dim result As AcceptCommand
    Set result = New AcceptCommand
    result.Initialize a_view, a_validationManager
    Set CreateAcceptCommand = result

End Function

''' <summary>   Creates a new instance of the <see cref="AggregateErrorFormatter"/> class to share
''' with external workbooks. Creates a new composite formatter that aggregates/combines all
''' the specified formatters. </summary>
Public Function CreateAggregateErrorFormatter(ParamArray a_formatters() As Variant) As IValidationErrorFormatter

    Dim result As AggregateErrorFormatter
    Set result = New AggregateErrorFormatter
    result.Initialize a_formatters
    
    Set CreateAggregateErrorFormatter = result
    
End Function

''' <summary>   Creates a new instance of the <see cref="AppContext"/> class to share
''' with external workbooks. </summary>
Public Function CreateAppContext(Optional ByVal a_bindings As cc_isr_MVVM.IBindingManager, _
        Optional ByVal a_commands As cc_isr_MVVM.ICommandManager, _
        Optional ByVal a_validation As cc_isr_MVVM.IValidationManager, _
        Optional ByVal a_formatterFactory As cc_isr_MVVM.IStringFormatterFactory, _
        Optional ByVal a_adornerFactory As cc_isr_MVVM.IDynamicAdornerFactory, _
        Optional ByVal a_debugOutput As Boolean = False) As cc_isr_MVVM.AppContext
    
    Dim result As cc_isr_MVVM.AppContext
    Set result = New cc_isr_MVVM.AppContext
    result.Initialize a_bindings, a_commands, a_validation, a_formatterFactory, a_adornerFactory, a_debugOutput
    Set CreateAppContext = result
    
End Function

''' <summary>   Creates a new instance of the <see cref="AppContext"/> class to share
''' with external workbooks. </summary>
''' <returns>   A new instance of the <see cref="AppContext"/> class. </returns>
Public Function NewAppContext() As AppContext
    Set NewAppContext = New AppContext
End Function

''' <summary>   Creates a new instance of the <see cref="BindingManager"/> class to share
''' with external workbooks. </summary>
Public Function CreateBindingManager(ByVal a_context As IAppContext, _
        ByVal a_stringFormatterFactory As IStringFormatterFactory, _
        Optional ByVal a_factory As INotifierFactory, _
        Optional ByVal a_debugOutput As Boolean = False) As IBindingManager
    
    Dim result As BindingManager
    Set result = New BindingManager
    result.Initialize a_context, a_stringFormatterFactory, a_factory, a_debugOutput
    Set CreateBindingManager = result

End Function

''' <summary>   Creates a new instance of the <see cref="BindingManager"/> class to share
''' with external workbooks. </summary>
''' <returns>   A new instance of the <see cref="BindingManager"/> class. </returns>
Public Function NewBindingManager() As BindingManager
    Set NewBindingManager = New BindingManager
End Function

''' <summary>   Creates a new <see cref="BindingPath"/> from the specified property path string
''' and binding context. </summary>
Public Function CreateBindingPath(ByVal a_context As Object, ByVal a_path As String) As IBindingPath
    
    Dim result As BindingPath
    Set result = New BindingPath
    result.Initialize a_context, a_path
    Set CreateBindingPath = result
    
End Function

''' <summary>   Creates a new instance of the <see cref="BindingPath"/> class to share
''' with external workbooks. </summary>
''' <returns>   A new instance of the <see cref="BindingPath"/> class. </returns>
Public Function NewBindingPath() As BindingPath
    Set NewBindingPath = New BindingPath
End Function

''' <summary>   Creates a new instance of the <see cref="BrowseCommand"/> class to share
''' with external workbooks. </summary>
''' <returns>   A new instance of the <see cref="BrowseCommand"/> class. </returns>
Public Function NewBrowseCommand() As BrowseCommand
    Set NewBrowseCommand = New BrowseCommand
End Function

''' <summary>   Creates a new instance of the <see cref="CancelCommand"/> class to share
''' with external workbooks. </summary>
''' <returns>   A new instance of the <see cref="CancelCommand"/> class. </returns>
Public Function NewCancelCommand() As CancelCommand
    
    Set NewCancelCommand = New CancelCommand

End Function

''' <summary>   Creates a new instance of the <see cref="CancelCommand"/>. </summary>
Public Function CreateCancelCommand(ByVal a_view As ICancellable) As ICommand
    Dim result As CancelCommand
    Set result = New CancelCommand
    result.Initialize a_view
    Set CreateCancelCommand = result
End Function

''' <summary>   Creates a new instance of the <see cref="CaptionPropertyBinding"/>. </summary>
Public Function CreateCaptionPropertyBinding(ByVal a_context As IAppContext, _
        ByVal a_source As IBindingPath, ByVal a_target As Object, _
        Optional ByVal a_validator As IValueValidator, _
        Optional ByVal a_converter As IValueConverter, _
        Optional ByVal a_stringFormatter As IStringFormatter, _
        Optional ByVal a_validationAdorner As IDynamicAdorner) As IPropertyBinding
    
    Dim result As CaptionPropertyBinding
    Set result = New CaptionPropertyBinding
    result.Initialize a_context, a_source, a_target, a_validator, a_converter, a_stringFormatter, a_validationAdorner
    Set CreateCaptionPropertyBinding = result
    
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
    
    Dim result As CheckBoxPropertyBinding
    Set result = New CheckBoxPropertyBinding
    result.Initialize a_context, a_source, a_target, a_targetProperty, a_mode, a_updateSource, _
                      a_validator, a_converter, a_stringFormatter, a_validationAdorner
    Set CreateCheckBoxPropertyBinding = result
    
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
    
    Dim result As ComboBoxPropertyBinding
    Set result = New ComboBoxPropertyBinding
    result.Initialize a_context, a_source, a_target, a_targetProperty, a_mode, a_updateSource, _
                      a_validator, a_converter, a_stringFormatter, a_validationAdorner
    Set CreateComboBoxPropertyBinding = result
    
End Function

''' <summary>   Creates a new instance of the <see cref="CommandBinding"/> class to share
''' with external workbooks. </summary>
''' <returns>   A new instance of the <see cref="CommandBinding"/> class. </returns>
Public Function NewCommandBinding() As CommandBinding
    Set NewCommandBinding = New CommandBinding
End Function

''' <summary>   Creates a new instance of the <see cref="CommandManager"/> class to share
''' with external workbooks. </summary>
''' <returns>   A new instance of the <see cref="CommandManager"/> class. </returns>
Public Function NewCommandManager() As CommandManager
    Set NewCommandManager = New CommandManager
End Function

''' <summary>   Creates a new instance of the <see cref="ContainerLayout"/> class to share
''' with external workbooks. </summary>
Public Function CreateContainerLayout(ByVal a_parent As MSForms.Controls, _
        Optional ByVal a_direction As LayoutDirection, _
        Optional ByVal a_top As Long = 0, _
        Optional ByVal a_left As Long = 0, _
        Optional ByVal a_margin As Long = 3) As IContainerLayout
    Dim result As ContainerLayout
    Set result = New ContainerLayout
    result.Initialize a_parent, a_direction, a_top, a_left, a_margin
    Set CreateContainerLayout = result
End Function

''' <summary>   Creates a new instance of the <see cref="ContainerLayout"/> class to share
''' with external workbooks. </summary>
''' <returns>   A new instance of the <see cref="ContainerLayout"/> class. </returns>
Public Function NewContainerLayout() As ContainerLayout
    Set NewContainerLayout = New ContainerLayout
End Function


''' <summary>   Creates a new instance of the <see cref="DecimalKeyValidator"/> class to share
''' with external workbooks. </summary>
''' <returns>   A new instance of the <see cref="DecimalKeyValidator"/> class. </returns>
Public Function NewDecimalKeyValidator() As DecimalKeyValidator
    Set NewDecimalKeyValidator = New DecimalKeyValidator
End Function

''' <summary>   Creates a new instance of the <see cref="DynamicControlPosition"/> class to share
''' with external workbooks. </summary>
Public Function CreateDynamicControlPosition(ByVal a_position As RelativePosition, _
        ByVal a_leftMargin As Double, _
        Optional ByVal a_topMargin As Variant, _
        Optional ByVal a_rightMargin As Variant, _
        Optional ByVal a_bottomMargin As Variant) As IDynamicPosition
    
    Dim result As cc_isr_MVVM.DynamicControlPosition
    Set result = New cc_isr_MVVM.DynamicControlPosition
    result.Initialize a_position, a_leftMargin, a_topMargin, a_rightMargin, a_bottomMargin
    Set CreateDynamicControlPosition = result
    
End Function

''' <summary>   Creates a new instance of the <see cref="DynamicControls"/> class to share
''' with external workbooks. </summary>
Public Function CreateDynamicControls(ByVal a_context As cc_isr_MVVM.IAppContext, _
        ByVal a_container As IContainerLayout) As IDynamicControlBuilder
    
    Dim result As DynamicControls
    Set result = New DynamicControls
    result.Initialize a_context, a_container
    Set CreateDynamicControls = result
    
End Function

''' <summary>   Creates a new instance of the <see cref="ExampleViewModel"/> class to share
''' with external workbooks. </summary>
Public Function CreateExampleViewModel() As ExampleViewModel

    Dim result As ExampleViewModel
    Set result = New ExampleViewModel
    Set CreateExampleViewModel = result
    
End Function

''' <summary>   Creates a new instance of the <see cref="Layout"/> class to share
''' with external workbooks. </summary>
Public Function CreateLayout(ByVal a_object As Object, Optional ByVal a_newWidthPercent As Long = 50, _
        Optional ByVal a_newHeightPercent As Long = 50) As Layout
    
    Dim result As Layout
    Set result = New Layout
    result.Initialize a_object, a_newWidthPercent, a_newHeightPercent
    Set CreateLayout = result
    
End Function

''' <summary>   Creates a new instance of the <see cref="LayoutPanel"/> class to share
''' with external workbooks. </summary>
Public Function CreateLayoutPanel(ByVal a_parentForm As MSForms.UserForm, _
        Optional ByVal a_parent As MSForms.Controls, _
        Optional ByVal a_height As Single, _
        Optional ByVal a_width As Single) As LayoutPanel
    
    Dim result As LayoutPanel
    Set result = New LayoutPanel
    result.Initialize a_parentForm, a_parent, a_height, a_width
    Set CreateLayoutPanel = result
    
End Function

''' <summary>   Creates a new instance of the <see cref="ListBoxPropertyBinding"/> class to share
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
    
    Dim result As ListBoxPropertyBinding
    Set result = New ListBoxPropertyBinding
    result.Initialize a_context, a_source, a_target, a_targetProperty, a_mode, _
                      a_updateSource, a_validator, a_converter, a_stringFormatter, a_validationAdorner
    Set CreateListBoxPropertyBinding = result
    
End Function

''' <summary>   Creates a new instance of the <see cref="MultiPagePropertyBinding"/> class to share
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
        
    Dim result As MultiPagePropertyBinding
    Set result = New MultiPagePropertyBinding
    result.Initialize a_context, a_source, a_target, a_targetProperty, a_mode, _
                      a_updateSource, a_validator, a_converter, a_stringFormatter, a_validationAdorner
    Set CreateMultiPagePropertyBinding = result
    
End Function

''' <summary>   Creates a new instance of the <see cref="OneWayPropertyBinding"/> class to share
''' with external workbooks. </summary>
Public Function CreateOneWayPropertyBinding(ByVal a_context As IAppContext, ByVal a_source As IBindingPath, _
        ByVal a_target As IBindingPath, _
        Optional ByVal a_validator As IValueValidator, _
        Optional ByVal a_converter As IValueConverter, _
        Optional ByVal a_stringFormatter As IStringFormatter, _
        Optional ByVal a_validationAdorner As IDynamicAdorner) As IPropertyBinding
    
    Dim result As OneWayPropertyBinding
    Set result = New OneWayPropertyBinding
    result.Initialize a_context, a_source, a_target, a_validator, a_converter, a_stringFormatter, a_validationAdorner
    Set CreateOneWayPropertyBinding = result
    
End Function

''' <summary>   Creates a new instance of the <see cref="OptionButtonPropertyBinding"/> class to share
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
    
    Dim result As OptionButtonPropertyBinding
    Set result = New OptionButtonPropertyBinding
    result.Initialize a_context, a_source, a_target, a_targetProperty, a_mode, _
                      a_updateSource, a_validator, a_converter, a_stringFormatter, a_validationAdorner
    Set CreateOptionButtonPropertyBinding = result
    
End Function

''' <summary>   Creates a new instance of the <see cref="PropertyBindingBase"/> class to share
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

    Dim result As PropertyBindingBase
    Set result = New PropertyBindingBase
    result.Initialize a_context, a_source, a_targetContext, a_targetPropertyPath, a_mode, a_updateSource, _
            a_converter, a_stringFormatter, a_validator, a_validationAdorner, a_cancelExitOnValidationError
    Set CreatePropertyBindingBase = result
    
End Function


''' <summary>   Creates a new instance of the <see cref="PropertyChangeNotifierBase"/> class to share
''' with external workbooks. </summary>
''' <returns>   A new instance of the <see cref="PropertyChangeNotifierBase"/> class. </returns>
Public Function NewPropertyChangeNotifierBase() As PropertyChangeNotifierBase
    Set NewPropertyChangeNotifierBase = New PropertyChangeNotifierBase
End Function

''' <summary>   Creates a new instance of the <see cref="RequiredStringValidator"/> class to share
''' with external workbooks. </summary>
''' <returns>   A new instance of the <see cref="RequiredStringValidator"/> class. </returns>
Public Function NewRequiredStringValidator() As RequiredStringValidator
    Set NewRequiredStringValidator = New RequiredStringValidator
End Function

''' <summary>   Creates a new instance of the <see cref="ScrollBarPropertyBinding"/> class to share
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
        
    Dim result As ScrollBarPropertyBinding
    Set result = New ScrollBarPropertyBinding
    result.Initialize a_context, a_source, a_target, a_targetProperty, a_mode, _
                      a_updateSource, a_validator, a_converter, a_stringFormatter, a_validationAdorner
    Set CreateScrollBarPropertyBinding = result
    
End Function

''' <summary>   Creates a new instance of the <see cref="SpinButtonPropertyBinding"/> class to share
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
        
    Dim result As SpinButtonPropertyBinding
    Set result = New SpinButtonPropertyBinding
    result.Initialize a_context, a_source, a_target, a_targetProperty, a_mode, _
                      a_updateSource, a_validator, a_converter, a_stringFormatter, a_validationAdorner
    Set CreateSpinButtonPropertyBinding = result
    
End Function

''' <summary>   Creates a new instance of the <see cref="StringFormatter"/> class to share
''' with external workbooks. </summary>
Public Function CreateStringFormatter(ByVal a_formatString As String) As IStringFormatter
    
    Dim result As StringFormatter
    Set result = New StringFormatter
    result.Initialize a_formatString
    Set CreateStringFormatter = result

End Function

''' <summary>   Creates a new instance of the <see cref="StringFormatterFactory"/> class to share
''' with external workbooks. </summary>
''' <returns>   A new instance of the <see cref="StringFormatterFactory"/> class. </returns>
Public Function NewStringFormatterFactory() As StringFormatterFactory
    
    Set NewStringFormatterFactory = New StringFormatterFactory

End Function

Public Function CreateTabStripPropertyBinding(ByVal a_context As IAppContext, a_source As IBindingPath, _
        ByVal a_target As MSForms.TabStrip, _
        Optional ByVal a_targetProperty As String = "Value", _
        Optional ByVal a_mode As BindingMode = BindingMode.TwoWayBinding, _
        Optional ByVal a_updateSource As BindingUpdateSourceTrigger = OnPropertyChanged, _
        Optional ByVal a_validator As IValueValidator, _
        Optional ByVal a_converter As IValueConverter, _
        Optional ByVal a_stringFormatter As IStringFormatter, _
        Optional ByVal a_validationAdorner As IDynamicAdorner) As IPropertyBinding

    Dim result As TabStripPropertyBinding
    Set result = New TabStripPropertyBinding
    result.Initialize a_context, a_source, a_target, a_targetProperty, a_mode, _
                      a_updateSource, a_validator, a_converter, a_stringFormatter, a_validationAdorner
    Set CreateTabStripPropertyBinding = result
    
End Function

''' <summary>   Creates a new instance of the <see cref="TextBoxPropertyBinding"/> class to share
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
    
    Dim result As TextBoxPropertyBinding
    Set result = New TextBoxPropertyBinding
    result.Initialize a_context, a_source, a_target, a_targetProperty, a_mode, _
                      a_updateSource, a_validator, a_converter, a_stringFormatter, a_validationAdorner
    Set CreateTextBoxPropertyBinding = result
    
End Function

''' <summary>   Creates a new instance of the <see cref="ValidationError"/> class to share
''' with external workbooks. </summary>
Public Function CreateValidationError(ByVal a_binding As IPropertyBinding, ByVal a_message As String) As IValidationError
    
    Dim result As ValidationError
    Set result = New ValidationError
    Set result.Binding = a_binding
    result.Message = a_message
    Set CreateValidationError = result

End Function

''' <summary>   Creates a new instance of the <see cref="ValidationErrorAdorner"/> class to share
''' with external workbooks. </summary>
Public Function CreateValidationErrorAdorner(ByVal a_target As MSForms.Control, _
        Optional ByVal a_targetFormatter As cc_isr_MVVM.IValidationErrorFormatter = Nothing, _
        Optional ByVal a_adornerLabel As Variant, _
        Optional ByVal a_labelFormatter As cc_isr_MVVM.IValidationErrorFormatter = Nothing, _
        Optional ByVal a_labelPosition As cc_isr_MVVM.IDynamicPosition = Nothing, _
        Optional ByVal a_adornerIcon As Variant, _
        Optional ByVal a_iconFormatter As cc_isr_MVVM.IValidationErrorFormatter = Nothing, _
        Optional ByVal a_iconPosition As cc_isr_MVVM.IDynamicPosition = Nothing) As IDynamicAdorner
    
    Dim result As ValidationErrorAdorner
    Set result = New ValidationErrorAdorner
    result.Initialize a_target, a_targetFormatter, a_adornerLabel, a_labelFormatter, a_labelPosition, _
                      a_adornerIcon, a_iconFormatter, a_iconPosition
    Set CreateValidationErrorAdorner = result
    
End Function


''' <summary>   Creates a new instance of the <see cref="ValidationErrorFormatter"/> class to share
''' with external workbooks. </summary>
Public Function CreateValidationErrorFormatter() As ValidationErrorFormatter
    
    Dim result As ValidationErrorFormatter
    Set result = New ValidationErrorFormatter
    Set CreateValidationErrorFormatter = result

End Function

''' <summary>   Creates a new instance of the <see cref="ValidationErrors"/> class to share
''' with external workbooks.Creates a new collection of IValidationError objects for the specified
''' binding context (ViewModel). </summary>
Public Function CreateValidationErrors(ByVal a_bindingContext As Object) As ValidationErrors
    
    Dim result As ValidationErrors
    Set result = New ValidationErrors
    result.Initialize a_bindingContext
    Set CreateValidationErrors = result

End Function

''' <summary>   Creates a new instance of the <see cref="ValidationManager"/> class to share
''' with external workbooks. </summary>
Public Function CreateValidationManager(Optional ByVal a_errorAdornerFactory As IDynamicAdornerFactory) As ValidationManager
    Dim result As ValidationManager
    Set result = New ValidationManager
    result.Initialize a_errorAdornerFactory
    Set CreateValidationManager = result
End Function

''' <summary>   Returns a singleton instance of the <see cref="ValidationManager"/> class to share
''' with external workbooks. </summary>
Public Function ValidationManager() As ValidationManager
    If This.ValidationManger Is Nothing Then
        This.ValidationManger = Me.CreateValidationManager
    End If
    Set ValidationManager = This.ValidationManger
End Function

''' <summary>   Creates a new instance of the <see cref="ValueRangeValidator"/> class to share
''' with external workbooks. </summary>
Public Function CreateValueRangeValidator(ByVal a_min As Variant, ByVal a_max As Variant) As IValueValidator
    Dim result As ValueRangeValidator
    Set result = New ValueRangeValidator
    result.Min = a_min
    result.Max = a_max
    Set CreateValueRangeValidator = result
End Function

''' <summary>   Creates a new instance of the <see cref="WorksheetAreaPropertyBinding"/> class to share
''' with external workbooks. </summary>
Public Function CreateWorksheetAreaPropertyBinding(ByVal a_context As IAppContext, ByVal a_source As IBindingPath, ByVal a_target As Excel.Range, _
        Optional ByVal a_targetProperty As String = "Value", _
        Optional ByVal a_mode As BindingMode = BindingMode.TwoWayBinding, _
        Optional ByVal a_validator As IValueValidator, _
        Optional ByVal a_converter As IValueConverter, _
        Optional ByVal a_stringFormatter As IStringFormatter, _
        Optional ByVal a_validationAdorner As IDynamicAdorner) As IPropertyBinding
    
    Dim result As WorksheetAreaPropertyBinding
    Set result = New WorksheetAreaPropertyBinding
    result.Initialize a_context, a_source, a_target, a_targetProperty, a_mode, a_validator, _
                      a_converter, a_stringFormatter, a_validationAdorner
    Set CreateWorksheetAreaPropertyBinding = result
    
End Function

''' <summary>   Creates a new instance of the <see cref="WorksheetCellPropertyBinding"/> class to share
''' with external workbooks. </summary>
Public Function CreateWorksheetCellPropertyBinding(ByVal a_context As IAppContext, ByVal a_source As IBindingPath, _
        ByVal a_target As Excel.Range, _
        Optional ByVal a_targetProperty As String = "Value", _
        Optional ByVal a_mode As BindingMode = BindingMode.TwoWayBinding, _
        Optional ByVal a_validator As IValueValidator, _
        Optional ByVal a_converter As IValueConverter, _
        Optional ByVal a_stringFormatter As IStringFormatter, _
        Optional ByVal a_validationAdorner As IDynamicAdorner) As IPropertyBinding
    
    Dim result As WorksheetCellPropertyBinding
    Set result = New WorksheetCellPropertyBinding
    result.Initialize a_context, a_source, a_target, a_targetProperty, a_mode, a_validator, _
                      a_converter, a_stringFormatter, a_validationAdorner
    Set CreateWorksheetCellPropertyBinding = result
    
End Function

