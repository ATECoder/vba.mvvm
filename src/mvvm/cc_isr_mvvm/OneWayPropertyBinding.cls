VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OneWayPropertyBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "A property binding that binds any property of any source, one-way to any property of a UI element target."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   A property binding that binds any property of any source,
'''             one-way to any property of a UI element target. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit
Implements IPropertyBinding
Implements cc_isr_Core.IDisposable
Implements IHandlePropertyChanged

Private Type TState
    Base As PropertyBindingBase
    Notifier As INotifyPropertyChanged
    Handler As IHandlePropertyChanged
End Type

Private This As TState

''' <summary>   Initializes and returns an <see cref="IPropertyBinding"> representation of
'''             this <see cref="OneWayPropertyBinding"/> object. </summary>
''' <returns>   [<see cref="IPropertyBinding"/>]. </returns>
Public Function Initialize(ByVal a_context As IAppContext, ByVal a_source As IBindingPath, _
        ByVal a_target As IBindingPath, _
        Optional ByVal a_validator As IValueValidator, _
        Optional ByVal a_converter As IValueConverter, _
        Optional ByVal a_stringFormatter As IStringFormatter, _
        Optional ByVal a_validationAdorner As IDynamicAdorner) As IPropertyBinding
    
    Dim p_bindingBase As PropertyBindingBase
    Set p_bindingBase = Factory.NewPropertyBindingBase().Initialize(a_context, a_source, a_target.Context, _
            a_target.Path, cc_isr_MVVM.BindingMode.OneWayBinding, cc_isr_MVVM.BindingUpdateSourceTrigger.Never, _
            a_converter, a_stringFormatter, a_validator, a_validationAdorner)
    
    If Not a_validator Is Nothing Then
        p_bindingBase.AsINotifyValidationError.RegisterHandler Factory.ValidationManager
    End If
    
    Me.InjectBindingInfo p_bindingBase
    
    Set Initialize = Me
    
End Function

Public Sub InjectBindingInfo(ByVal a_bindingInfo As PropertyBindingBase)

    cc_isr_Core.GuardClauses.GuardNullReference a_bindingInfo, VBA.Information.TypeName(Me)
    cc_isr_Core.GuardClauses.GuardDoubleInitialization This.Base, VBA.Information.TypeName(Me)
    cc_isr_Core.GuardClauses.GuardDoubleInitialization This.Handler, VBA.Information.TypeName(Me)
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
    This.Base.ApplyToTarget
End Sub

Private Property Get IPropertyBinding_CancelExitOnValidationError() As Boolean
    IPropertyBinding_CancelExitOnValidationError = This.Base.CancelExitOnValidationError
End Property

Private Property Get IPropertyBinding_Converter() As IValueConverter
    Set IPropertyBinding_Converter = This.Base.Converter
End Property

Private Property Get IPropertyBinding_DefaultTargetProperty() As String
    IPropertyBinding_DefaultTargetProperty = VBA.vbNullString
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

