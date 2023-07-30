VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TabStripPropertyBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder cc_isr_MVVM.Infrastructure.Bindings.PropertyBindings
'@PredeclaredId
'@Exposed
Option Explicit
Implements IPropertyBinding
Implements cc_isr_Core.IDisposable
Implements IHandlePropertyChanged

Private WithEvents TargetEventSource As MSForms.TabStrip
Attribute TargetEventSource.VB_VarHelpID = -1

Private Type TState
    Base As PropertyBindingBase
    Handler As IHandlePropertyChanged
End Type

Private This As TState

''' <summary>   Initializes this <see cref="TabStripPropertyBinding"/> object. </summary>
Public Sub Initialize(ByVal a_context As IAppContext, a_source As IBindingPath, ByVal a_target As MSForms.TabStrip, _
        Optional ByVal a_targetProperty As String = "Value", _
        Optional ByVal a_mode As BindingMode = BindingMode.TwoWayBinding, _
        Optional ByVal a_updateSource As BindingUpdateSourceTrigger = OnPropertyChanged, _
        Optional ByVal a_validator As IValueValidator, _
        Optional ByVal a_converter As IValueConverter, _
        Optional ByVal a_stringFormatter As IStringFormatter, _
        Optional ByVal a_validationAdorner As IDynamicAdorner)

    Dim p_bindingBase As PropertyBindingBase
    Set p_bindingBase = Factory.CreatePropertyBindingBase(a_context, _
            a_source, a_target, a_targetProperty, _
            a_mode, _
            a_updateSource, _
            a_validator, _
            a_converter, _
            a_stringFormatter, _
            a_validationAdorner)
    
    Me.InjectBindingInfo p_bindingBase
    
End Sub

Public Sub InjectBindingInfo(ByVal BindingInfo As PropertyBindingBase)
    
    cc_isr_Core.GuardClauses.GuardNullReference BindingInfo, VBA.Information.TypeName(Me)
    cc_isr_Core.GuardClauses.GuardDoubleInitialization This.Base, VBA.Information.TypeName(Me)
    cc_isr_Core.GuardClauses.GuardDoubleInitialization This.Handler, VBA.Information.TypeName(Me)
    Set This.Base = BindingInfo
    Set This.Handler = BindingInfo

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

Private Sub IHandlePropertyChanged_HandlePropertyChanged(ByVal Source As Object, ByVal a_propertyName As String)
    This.Handler.HandlePropertyChanged Source, a_propertyName
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
    IPropertyBinding_DefaultTargetProperty = BindingDefaults.TabStripDefaultProperty
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

