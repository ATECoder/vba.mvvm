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
''' classes. </summary>
''' <remarks>
''' <see href="https://learn.Microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/set-up-vb-project-using-class"/>
''' </remarks>
Private Sub Class_Initialize()
End Sub

''' <summary>   Constructs a new instance of the <see cref="AcceptCommand"/> object. </summary>
''' <returns>   [<see cref="AcceptCommand"/>]. </returns>
Public Function NewAcceptCommand() As AcceptCommand
    
    Set NewAcceptCommand = New AcceptCommand

End Function

''' <summary>   Creates a new instance of the <see cref="AggregateErrorFormatter"/> object.
'''             that aggregates/combines all the specified formatters. </summary>
''' <returns>   [<see cref="ValidationErrorFormatter"/>]. </returns>
Public Function NewAggregateErrorFormatter() As AggregateErrorFormatter
    
    Set NewAggregateErrorFormatter = New AggregateErrorFormatter

End Function

''' <summary>   Constructs a new instance of the <see cref="AppContext"/> object. </summary>
''' <returns>   [<see cref="AppContext"/>]. </returns>
Public Function NewAppContext() As AppContext
    
    Set NewAppContext = New AppContext

End Function

''' <summary>   Constructs a new instance of the <see cref="BindingManager"/> object. </summary>
''' <returns>   [<see cref="BindingManager"/>]. </returns>
Public Function NewBindingManager() As BindingManager
    
    Set NewBindingManager = New BindingManager

End Function

''' <summary>   Constructs a new instance of the <see cref="BindingPath"/> object. </summary>
''' <returns>   [<see cref="BindingPath"/>]. </returns>
Public Function NewBindingPath() As BindingPath
    
    Set NewBindingPath = New BindingPath

End Function

''' <summary>   Constructs a new instance of the <see cref="CancelCommand"/> object. </summary>
''' <returns>   [<see cref="CancelCommand"/>]. </returns>
Public Function NewCancelCommand() As CancelCommand
    
    Set NewCancelCommand = New CancelCommand

End Function

''' <summary>   Creates a new instance of the <see cref="CaptionPropertyBinding"/>. </summary>
''' <returns>   [<see cref="CaptionPropertyBinding"/>]. </returns>
Public Function NewCaptionPropertyBinding() As CaptionPropertyBinding
    
    Set NewCaptionPropertyBinding = New CaptionPropertyBinding

End Function

''' <summary>   Constructs a new instance of the <see cref="CheckBoxPropertyBinding"/> object. </summary>
''' <returns>   [<see cref="CheckBoxPropertyBinding"/>]. </returns>
Public Function NewCheckBoxPropertyBinding() As CheckBoxPropertyBinding
    
    Set NewCheckBoxPropertyBinding = New CheckBoxPropertyBinding

End Function

''' <summary>   Constructs a new instance of the <see cref="ComboBoxPropertyBinding"/> object. </summary>
''' <returns>   [<see cref="ComboBoxPropertyBinding"/>]. </returns>
Public Function NewComboBoxPropertyBinding() As ComboBoxPropertyBinding
    
    Set NewComboBoxPropertyBinding = New ComboBoxPropertyBinding

End Function

''' <summary>   Constructs a new instance of the <see cref="CommandBinding"/> object. </summary>
''' <returns>   [<see cref="CommandBinding"/>]. </returns>
Public Function NewCommandBinding() As CommandBinding
    
    Set NewCommandBinding = New CommandBinding

End Function

''' <summary>   Constructs a new instance of the <see cref="CommandManager"/> object. </summary>
''' <returns>   [<see cref="CommandManager"/>]. </returns>
Public Function NewCommandManager() As CommandManager
    
    Set NewCommandManager = New CommandManager

End Function

''' <summary>   Constructs a new instance of the <see cref="EventCommand"/> object. </summary>
''' <returns>   [<see cref="EventCommand"/>]. </returns>
Public Function NewEventCommand() As EventCommand
    
    Set NewEventCommand = New EventCommand

End Function

''' <summary>   Constructs a new instance of the <see cref="ContainerLayout"/> object. </summary>
''' <returns>   [<see cref="ContainerLayout"/>]. </returns>
Public Function NewContainerLayout() As ContainerLayout
    
    Set NewContainerLayout = New ContainerLayout

End Function


''' <summary>   Constructs a new instance of the <see cref="DecimalKeyValidator"/> object. </summary>
''' <returns>   [<see cref="DecimalKeyValidator"/>]. </returns>
Public Function NewDecimalKeyValidator() As DecimalKeyValidator
    
    Set NewDecimalKeyValidator = New DecimalKeyValidator

End Function

''' <summary>   Constructs a new instance of the <see cref="DynamicControlPosition"/> object. </summary>
''' <returns>   [<see cref="DynamicControlPosition"/>]. </returns>
Public Function NewDynamicControlPosition() As DynamicControlPosition
    
    Set NewDynamicControlPosition = New DynamicControlPosition

End Function

''' <summary>   Constructs a new instance of the <see cref="DynamicControls"/> object. </summary>
''' <returns>   [<see cref="DynamicControls"/>]. </returns>
Public Function NewDynamicControls() As DynamicControls
    
    Set NewDynamicControls = New DynamicControls

End Function

''' <summary>   Constructs a new instance of the <see cref="Layout"/> object. </summary>
''' <returns>   [<see cref="Layout"/>]. </returns>
Public Function NewLayout() As Layout
    
    Set NewLayout = New Layout

End Function

''' <summary>   Constructs a new instance of the <see cref="LayoutPanel"/> object. </summary>
''' <returns>   [<see cref="LayoutPanel"/>]. </returns>
Public Function NewLayoutPanel() As LayoutPanel
    
    Set NewLayoutPanel = New LayoutPanel

End Function

''' <summary>   Constructs a new instance of the <see cref="ListBoxPropertyBinding"/> object. </summary>
''' <returns>   [<see cref="ListBoxPropertyBinding"/>]. </returns>
Public Function NewListBoxPropertyBinding() As ListBoxPropertyBinding
    
    Set NewListBoxPropertyBinding = New ListBoxPropertyBinding

End Function

''' <summary>   Constructs a new instance of the <see cref="MultiPagePropertyBinding"/> object. </summary>
''' <returns>   [<see cref="MultiPagePropertyBinding"/>]. </returns>
Public Function NewMultiPagePropertyBinding() As MultiPagePropertyBinding
    
    Set NewMultiPagePropertyBinding = New MultiPagePropertyBinding

End Function

''' <summary>   Constructs a new instance of the <see cref="OneWayPropertyBinding"/> object. </summary>
''' <returns>   [<see cref="OneWayPropertyBinding"/>]. </returns>
Public Function NewOneWayPropertyBinding() As OneWayPropertyBinding
    
    Set NewOneWayPropertyBinding = New OneWayPropertyBinding

End Function

''' <summary>   Constructs a new instance of the <see cref="OptionButtonPropertyBinding"/> object. </summary>
''' <returns>   [<see cref="OptionButtonPropertyBinding"/>]. </returns>
Public Function NewOptionButtonPropertyBinding() As OptionButtonPropertyBinding
    
    Set NewOptionButtonPropertyBinding = New OptionButtonPropertyBinding

End Function

''' <summary>   Constructs a new instance of the <see cref="PropertyBindingBase"/> object. </summary>
''' <returns>   [<see cref="PropertyBindingBase"/>]. </returns>
Public Function NewPropertyBindingBase() As PropertyBindingBase
    
    Set NewPropertyBindingBase = New PropertyBindingBase

End Function

''' <summary>   Constructs a new instance of the <see cref="PropertyChangeNotifierBase"/> object. </summary>
''' <returns>   [<see cref="PropertyChangeNotifierBase"/>]. </returns>
Public Function NewPropertyChangeNotifierBase() As PropertyChangeNotifierBase
    
    Set NewPropertyChangeNotifierBase = New PropertyChangeNotifierBase

End Function

''' <summary>   Constructs a new instance of the <see cref="RequiredStringValidator"/> object. </summary>
''' <returns>   [<see cref="RequiredStringValidator"/>]. </returns>
Public Function NewRequiredStringValidator() As RequiredStringValidator
    
    Set NewRequiredStringValidator = New RequiredStringValidator

End Function

''' <summary>   Constructs a new instance of the <see cref="ScrollBarPropertyBinding"/> object. </summary>
''' <returns>   [<see cref="ScrollBarPropertyBinding"/>]. </returns>
Public Function NewScrollBarPropertyBinding() As ScrollBarPropertyBinding
    
    Set NewScrollBarPropertyBinding = New ScrollBarPropertyBinding

End Function

''' <summary>   Constructs a new instance of the <see cref="SpinButtonPropertyBinding"/> object. </summary>
''' <returns>   [<see cref="SpinButtonPropertyBinding"/>]. </returns>
Public Function NewSpinButtonPropertyBinding() As SpinButtonPropertyBinding
    
    Set NewSpinButtonPropertyBinding = New SpinButtonPropertyBinding

End Function

''' <summary>   Constructs a new instance of the <see cref="StringFormatter"/> object. </summary>
''' <returns>   [<see cref="StringFormatter"/>]. </returns>
Public Function NewStringFormatter() As StringFormatter
    
    Set NewStringFormatter = New StringFormatter

End Function

''' <summary>   Constructs a new instance of the <see cref="StringFormatterFactory"/> object. </summary>
''' <returns>   [<see cref="StringFormatterFactory"/>]. </returns>
Public Function NewStringFormatterFactory() As StringFormatterFactory
    
    Set NewStringFormatterFactory = New StringFormatterFactory

End Function

''' <summary>   Constructs a new instance of the <see cref="TabStripPropertyBinding"/> object. </summary>
''' <returns>   [<see cref="TabStripPropertyBinding"/>]. </returns>
Public Function NewTabStripPropertyBinding() As TabStripPropertyBinding
    
    Set NewTabStripPropertyBinding = New TabStripPropertyBinding

End Function

''' <summary>   Constructs a new instance of the <see cref="TextBoxPropertyBinding"/> object. </summary>
''' <returns>   [<see cref="TextBoxPropertyBinding"/>]. </returns>
Public Function NewTextBoxPropertyBinding() As TextBoxPropertyBinding
    
    Set NewTextBoxPropertyBinding = New TextBoxPropertyBinding

End Function

''' <summary>   Constructs a new instance of the <see cref="ValidationError"/> object. </summary>
''' <returns>   [<see cref="ValidationError"/>]. </returns>
Public Function NewValidationError() As ValidationError
    
    Set NewValidationError = New ValidationError

End Function

''' <summary>   Constructs a new instance of the <see cref="ValidationErrorAdorner"/> object. </summary>
''' <returns>   [<see cref="ValidationErrorAdorner"/>]. </returns>
Public Function NewValidationErrorAdorner() As ValidationErrorAdorner
    
    Set NewValidationErrorAdorner = New ValidationErrorAdorner

End Function

''' <summary>   Constructs a new instance of the <see cref="ValidationErrorFormatter"/> object. </summary>
''' <returns>   [<see cref="ValidationErrorFormatter"/>]. </returns>
Public Function NewValidationErrorFormatter() As ValidationErrorFormatter
    
    Set NewValidationErrorFormatter = New ValidationErrorFormatter

End Function

''' <summary>   Constructs a new instance of the <see cref="ValidationErrors"/> object. </summary>
''' <returns>   [<see cref="ValidationErrors"/>]. </returns>
Public Function NewValidationErrors() As ValidationErrors
    
    Set NewValidationErrors = New ValidationErrors

End Function

''' <summary>   Constructs a new instance of the <see cref="ValidationManager"/> object. </summary>
''' <returns>   [<see cref="ValidationManager"/>]. </returns>
Public Function NewValidationManager() As ValidationManager
    
    Set NewValidationManager = New ValidationManager

End Function

''' <summary>   Returns a singleton instance of the <see cref="ValidationManager"/> object. </summary>
Public Function ValidationManager() As ValidationManager
    
    If This.ValidationManger Is Nothing Then This.ValidationManger = Me.NewValidationManager
    Set ValidationManager = This.ValidationManger
    
End Function

''' <summary>   Constructs a new instance of the <see cref="ValueRangeValidator"/> object. </summary>
''' <returns>   [<see cref="ValueRangeValidator"/>]. </returns>
Public Function NewValueRangeValidator() As ValueRangeValidator
    
    Set NewValueRangeValidator = New ValueRangeValidator

End Function

''' <summary>   Constructs a new instance of the <see cref="WorksheetAreaPropertyBinding"/> object. </summary>
''' <returns>   [<see cref="WorksheetAreaPropertyBinding"/>]. </returns>
Public Function NewWorksheetAreaPropertyBinding() As WorksheetAreaPropertyBinding
    
    Set NewWorksheetAreaPropertyBinding = New WorksheetAreaPropertyBinding

End Function

''' <summary>   Constructs a new instance of the <see cref="WorksheetCellPropertyBinding"/> object. </summary>
''' <returns>   [<see cref="WorksheetCellPropertyBinding"/>]. </returns>
Public Function NewWorksheetCellPropertyBinding() As WorksheetCellPropertyBinding
    
    Set NewWorksheetCellPropertyBinding = New WorksheetCellPropertyBinding

End Function
