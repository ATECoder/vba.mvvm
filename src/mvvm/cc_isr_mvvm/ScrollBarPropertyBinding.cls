VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ScrollBarPropertyBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   A property binding that binds any property of any source to
'''             properties of a ScrollBar UI element target. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit
Implements IPropertyBinding
Implements cc_isr_Core.IDisposable
Implements IHandlePropertyChanged

Private WithEvents TargetEventSource As MSForms.ScrollBar
Attribute TargetEventSource.VB_VarHelpID = -1

Private Type TState
    Base As PropertyBindingBase
    Handler As IHandlePropertyChanged
End Type

Private This As TState

''' <summary>   Initializes and returns an <see cref="IPropertyBinding"> representation of
'''             this <see cref="ScrollBarPropertyBinding"/> object. </summary>
''' <returns>   [<see cref="IPropertyBinding"/>]. </returns>
Public Function Initialize(ByVal a_context As IAppContext, ByVal a_source As IBindingPath, _
        ByVal a_target As MSForms.ScrollBar, _
        Optional ByVal a_targetProperty As String = "Value", _
        Optional ByVal a_mode As BindingMode = BindingMode.TwoWayBinding, _
        Optional ByVal a_updateSource As BindingUpdateSourceTrigger = OnPropertyChanged, _
        Optional ByVal a_validator As IValueValidator, _
        Optional ByVal a_converter As IValueConverter, _
        Optional ByVal a_stringFormatter As IStringFormatter, _
        Optional ByVal a_validationAdorner As IDynamicAdorner) As IPropertyBinding
        
    Dim p_bindingBase As PropertyBindingBase
    Set p_bindingBase = Factory.NewPropertyBindingBase().Initialize(a_context, a_source, a_target, a_targetProperty, _
            a_mode, a_updateSource, a_converter, a_stringFormatter, a_validator, a_validationAdorner)
    
    If Not a_validator Is Nothing Then
        p_bindingBase.AsINotifyValidationError.RegisterHandler Factory.ValidationManager
    End If
    
    Me.InjectBindingInfo p_bindingBase
    
    Set Initialize = Me
    
End Function

Public Sub InjectBindingInfo(ByVal a_bindingInfo As PropertyBindingBase)
    
    Dim p_source As String
    p_source = ThisWorkbook.VBProject.Name & "." & VBA.Information.TypeName(Me) & ".InjectBindingInfo"
    cc_isr_Core_IO.GuardClauses.GuardNullReference a_bindingInfo, p_source
    cc_isr_Core_IO.GuardClauses.GuardDoubleInitialization This.Base, p_source
    cc_isr_Core_IO.GuardClauses.GuardDoubleInitialization This.Handler, p_source
    cc_isr_Core_IO.GuardClauses.GuardExpression Not TypeOf a_bindingInfo.Target.Context Is MSForms.ScrollBar, _
        p_source, "Invalid binding target type", _
        cc_isr_Core_IO.UserDefinedErrors.InvalidArgumentError
    Set This.Base = a_bindingInfo
    Set This.Handler = a_bindingInfo

End Sub

' - - - - - - - - - - - - - - - -
'
' Disposable implementation
'
' - - - - - - - - - - - - - - - -

Private Sub IDisposable_Dispose()
    Set This.Handler = Nothing
    cc_isr_Core.DisposableExtensions.TryDispose This.Base
    Set This.Base = Nothing
End Sub

Private Sub IHandlePropertyChanged_HandlePropertyChanged(ByVal a_source As Object, ByVal a_propertyName As String)
    This.Handler.HandlePropertyChanged a_source, a_propertyName
End Sub

Private Sub IPropertyBinding_Apply()
    This.Base.Apply
End Sub

Private Property Get IPropertyBinding_CancelExitOnValidationError() As Boolean
    IPropertyBinding_CancelExitOnValidationError = This.Base.CancelExitOnValidationError
End Property

Private Property Get IPropertyBinding_Converter() As IValueConverter
    Set IPropertyBinding_Converter = This.Base.Converter
End Property

Private Property Get IPropertyBinding_DefaultTargetProperty() As String
    IPropertyBinding_DefaultTargetProperty = BindingDefaults.ScrollBarDefaultProperty
End Property

Private Property Get IPropertyBinding_Mode() As BindingMode
    IPropertyBinding_Mode = This.Base.Mode
End Property

Private Property Get IPropertyBinding_Source() As IBindingPath
    Set IPropertyBinding_Source = This.Base.Source
End Property

Private Property Get IPropertyBinding_StringFormatter() As IStringFormatter
    Set IPropertyBinding_StringFormatter = This.Base.StringFormatter
End Property

Private Property Get IPropertyBinding_Target() As IBindingPath
    Set IPropertyBinding_Target = This.Base.Target
End Property

Private Property Get IPropertyBinding_UpdateSourceTrigger() As BindingUpdateSourceTrigger
    IPropertyBinding_UpdateSourceTrigger = This.Base.UpdateSourceTrigger
End Property

Private Property Get IPropertyBinding_Validator() As IValueValidator
    Set IPropertyBinding_Validator = This.Base.Validator
End Property

Private Sub TargetEventSource_Change()
    If This.Base.UpdateSourceTrigger = OnPropertyChanged Then This.Base.ApplyToSource
End Sub

