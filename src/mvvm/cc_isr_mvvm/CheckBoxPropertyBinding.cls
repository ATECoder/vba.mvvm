VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CheckBoxPropertyBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   A property binding that binds any property of any source,
'''             two-way to <see cref="BindingDefaults.CaptionDefaultProperty"/>
'''             or <see cref="BindingDefaults.CheckBoxDefaultProperty"/> properties of
'''             a CheckBox UI element target. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit
Implements IPropertyBinding
Implements cc_isr_Core.IDisposable
Implements IHandlePropertyChanged
Private WithEvents TargetEventSource As MSForms.CheckBox
Attribute TargetEventSource.VB_VarHelpID = -1

Private Type TState
    Base As PropertyBindingBase
    Handler As IHandlePropertyChanged
End Type

Private This As TState

''' <summary>   Initializes and returns an <see cref="IPropertyBinding"> representation of
'''             this <see cref="CaptionPropertyBinding"/> object. </summary>
''' <returns>   [<see cref="IPropertyBinding"/>]. </returns>
Public Function Initialize(ByVal a_context As IAppContext, ByVal a_source As IBindingPath, _
        ByVal a_target As MSForms.CheckBox, _
        Optional ByVal a_targetProperty As Variant, _
        Optional ByVal a_mode As BindingMode = BindingMode.TwoWayBinding, _
        Optional ByVal a_updateSource As BindingUpdateSourceTrigger = OnPropertyChanged, _
        Optional ByVal a_validator As IValueValidator, _
        Optional ByVal a_converter As IValueConverter, _
        Optional ByVal a_stringFormatter As IStringFormatter, _
        Optional ByVal a_validationAdorner As IDynamicAdorner) As IPropertyBinding
    
    Dim p_localTargetProperty As String
    If VBA.IsMissing(a_targetProperty) Then
        p_localTargetProperty = Me.InferTargetPropertyFromSource(a_source)
    Else
        p_localTargetProperty = a_targetProperty
    End If
    
    Dim p_bindingBase As PropertyBindingBase
    Set p_bindingBase = Factory.NewPropertyBindingBase().Initialize(a_context, a_source, a_target, _
            p_localTargetProperty, a_mode, a_updateSource, a_converter, a_stringFormatter, _
            a_validator, a_validationAdorner)
    
    If Not a_validator Is Nothing Then
        p_bindingBase.AsINotifyValidationError.RegisterHandler Factory.ValidationManager
    End If

    Me.InjectBindingInfo p_bindingBase
    
    Set Initialize = Me
    
End Function

Public Function InferTargetPropertyFromSource(ByVal a_source As IBindingPath) As String
    
    Dim p_result As String
    Dim p_sourceValue As Variant
    
    If a_source.TryReadPropertyValue(a_outValue:=p_sourceValue) Then
        
        Select Case VBA.VarType(p_sourceValue)
        
            Case VBA.VbVarType.vbBoolean
                
                p_result = BindingDefaults.CheckBoxDefaultProperty ' "Value"
            
            Case VBA.VbVarType.vbString
                
                p_result = BindingDefaults.CaptionDefaultProperty ' "Caption"
            
            Case Else
                Debug.Print VBA.Information.TypeName(Me) & ": binding source [" & a_source.ToString() & "] involves implicit value conversion."
                p_result = BindingDefaults.CheckBoxDefaultProperty
        End Select
    
    Else
        
        p_result = BindingDefaults.CheckBoxDefaultProperty
    
    End If
    
    InferTargetPropertyFromSource = p_result

End Function

Public Sub InjectBindingInfo(ByVal a_bindingInfo As PropertyBindingBase)
    
    cc_isr_Core.GuardClauses.GuardNullReference a_bindingInfo, VBA.Information.TypeName(Me)
    cc_isr_Core.GuardClauses.GuardDoubleInitialization This.Base, VBA.Information.TypeName(Me)
    cc_isr_Core.GuardClauses.GuardDoubleInitialization This.Handler, VBA.Information.TypeName(Me)
    Set This.Base = a_bindingInfo
    Set This.Handler = a_bindingInfo
    'This.Base.AsIControlEvents.RegisterHandler Me

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
    IPropertyBinding_DefaultTargetProperty = BindingDefaults.CheckBoxDefaultProperty
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

