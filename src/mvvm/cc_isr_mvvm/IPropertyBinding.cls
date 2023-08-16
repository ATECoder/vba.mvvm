VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPropertyBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An object responsible for binding a ViewModel property path to a UI element."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   An object responsible for binding a ViewModel property path to a UI element. </summary>
''' <remarks>
''' The <c>IPropertyBinding</c> interface is mostly only useful internally. There’s little of interest here
''' that isn’t more appropriately covered by looking at the factory method for the <see cref="PropertyBindingBase"/> object.
''' </remarks>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit

''' <summary>   Indicates the binding mode/direction. </summary>
''' <remarks>
Public Enum BindingMode
    TwoWayBinding
    OneWayBinding
    OneWayToSource
    OneTimeBinding
End Enum

Public Enum BindingUpdateSourceTrigger
    ' VF: added NotSetYet  and renumbered enums
    NotSetYet = 0 ''avoid being called twice, for runtime optimization
    Never = 1
    OnPropertyChanged = 2 '
    OnKeyPress = 3 '
    OnExit = 4 '
    
    ' VF: added: required to capture Cut/Paste and Backspace/Delete in TextBox
    OnChange = 5
    
End Enum

''' <summary>   Gets a value indicating the binding mode/direction. </summary>
''' <remarks>
''' This enum value determines whether the binding synchronizes the target, the source,
''' or both. Note that binding modes <c>OneWayBinding</c> and <c>TwoWayBinding</c> both require
''' the source object to implement <c>INotifyPropertyChanged</c>.
''' </remarks>
Public Property Get Mode() As BindingMode
Attribute Mode.VB_Description = "Gets a value indicating the binding mode/direction."
End Property

''' <summary>   Gets a value indicating the binding update trigger. </summary>
''' <remarks>
''' This enum value determines when the binding gets to update its source.
''' When the Mode is <c>OneWayBinding</c> or <c>OneTimeBinding</c>, value
''' <c>UpdateSourceTrigger.Never</c> is used automatically.
'''
''' <c>OnKeyPress</c> gets to validate each individual keypress. Useful for
''' <c>TextBox</c> bindings that need a key validator.
'''
''' <c>OnExit</c> and <c>OnPropertyChanged</c> are still somewhat ambiguously defined,
''' but <c>OnExit</c> gets to keep the target control focused if there’s a validation error,
''' and <c>OnPropertyChanged</c> is [currently] firing for every <c>keypress</c> in a <c>TextBox</c>,
''' after the control’s Text property updates. Disambiguating the terms would mean breaking
''' with MSForms’ meaning of “updated”… which may actually be a good thing: <c>OnPropertyChanged</c>
''' would basically fire on exit but without a possible navigation-blocking validation error, and then
''' <c>OnKeyPress</c> mode would still need to behave like <c>OnPropertyChanged</c> as far as updating
''' the source goes.
'''
''' </remarks>
Public Property Get UpdateSourceTrigger() As BindingUpdateSourceTrigger
Attribute UpdateSourceTrigger.VB_Description = "Gets a value indicating the binding update trigger."
End Property

''' <summary>   Gets the binding path resolving to the binding source. </summary>
''' <remarks>
''' Old description. Needs to be updated:
''' Typically, that’s just a reference to the source ViewModel object. Technically,
''' it could really be any object that has any number of public properties [and ideally,
''' that implements INotifyPropertyChanged to support 2-way bindings].
''' </remarks>
Public Property Get Source() As IBindingPath
Attribute Source.VB_Description = "Gets the binding path resolving to the binding source."
End Property

''' <summary>   Gets the binding path resolving to the binding target. </summary>
''' <remarks>
''' Old description to be updated:
''' Typically, that’s just a reference to the target MSForms control. Technically, it could really be any object that has any number of public properties.
''' </remarks>
Public Property Get Target() As IBindingPath
Attribute Target.VB_Description = "Gets the binding path resolving to the binding target."
End Property

''' <summary>   Gets the name of the target property implicitly bound for the Target class type. </summary>
''' <remarks>
''' Each binding type has a default target property that is automatically
''' inferred from the type of target (and in some cases, from the data type of
''' the source property). For example, binding to a <c>TextBox</c> control
''' automatically wires up the control’s <c>Text</c> property, such that providing
''' a <c>TargetPropertyPath</c> is only needed when binding to another target property. </param>
''' </remarks>
Public Property Get DefaultTargetProperty() As String
Attribute DefaultTargetProperty.VB_Description = "Gets the name of the target property implicitly bound for the Target class type."
End Property

''' <summary>   Gets the converter (if any) used when applying the binding. </summary>
''' <remarks>
''' Each property binding can use an <c>IValueConverter</c> to “convert” a value midway
''' between the source and the target (or between the target and the source).
''' For example we can bind a <c>CheckBox</c> control to a <c>Boolean</c> property,
''' but if we need the checkbox checked when the property value is <c>False</c>, we can use
''' an <c>InverseBooleanConverter</c> to convert <c>True</c> to <c>False</c> as the binding is applied.
''' </remarks>
Public Property Get Converter() As IValueConverter
Attribute Converter.VB_Description = "Gets the converter (if any) used when applying the binding."
End Property

''' <summary>   Gets the value validator (if any) used when applying the binding. </summary>
''' <remarks>
''' When a binding is given an <c>IValueValidator</c>, it gets to start issuing validation errors,
''' which propagate to the <c>ViewModel</c> and can be used to pop a warning banner with the validation
''' error message. Note: the <c>AcceptCommand</c> class’ implementation of <c>ICommand.CanExecute</c>
''' makes it return <c>False</c> when the <c>ViewModel</c> has validation errors.
''' </remarks>
Public Property Get Validator() As IValueValidator
Attribute Validator.VB_Description = "Gets the value validator (if any) used when applying the binding."
End Property

''' <summary>   Gets the string formatter (if any) used when applying the binding to the target. </summary>
''' <remarks>
''' The <c>IAppContext.StringFormatterFactory<c> property (can be property-injected from <c>AppContext.Initialize<c>)
''' determines which formatter class type is going to be injected here. Supply a <c>VBStringFormatterFactory</c>
''' to use VB6/VBA <c>Strings.Format</c> format string syntax, or supply a <c>StringFormatterFactory</c> (or just leave
''' the defaults alone) to use the much more flexible .NET <c>String.Format</c> syntax.
''' </remarks>
Public Property Get StringFormatter() As IStringFormatter
Attribute StringFormatter.VB_Description = "Gets the string formatter (if any) used when applying the binding to the target."
End Property

''' <summary>   Determines whether the binding prevents focus leaving the control with invalid data. </summary>
''' <remarks>
''' </remarks>
Public Property Get CancelExitOnValidationError() As Boolean
Attribute CancelExitOnValidationError.VB_Description = "Determines whether the binding prevents focus leaving the control with invalid data."
End Property

''' <summary>   Applies the binding to the source to the target. </summary>
Public Sub Apply()
Attribute Apply.VB_Description = "Applies the binding."
End Sub
