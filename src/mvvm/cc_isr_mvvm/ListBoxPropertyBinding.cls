VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListBoxPropertyBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit
Implements IPropertyBinding
Implements cc_isr_Core.IDisposable
Implements IHandlePropertyChanged
Implements IHandleControlEvents

Private WithEvents TargetEventSource As MSForms.ListBox
Attribute TargetEventSource.VB_VarHelpID = -1

Private Type TState
    Base As PropertyBindingBase
    Handler As IHandlePropertyChanged
End Type

Private This As TState

''' <summary>   Initializes and returns an <see cref="IPropertyBinding"> representation of
'''             this <see cref="ListBoxPropertyBinding"/> object. </summary>
Public Function Initialize(ByVal a_context As IAppContext, a_source As IBindingPath, _
        ByVal a_target As MSForms.ListBox, _
        Optional ByVal a_targetProperty As Variant, _
        Optional ByVal a_mode As BindingMode = BindingMode.TwoWayBinding, _
        Optional ByVal a_updateSource As BindingUpdateSourceTrigger = OnPropertyChanged, _
        Optional ByVal a_validator As IValueValidator, _
        Optional ByVal a_converter As IValueConverter, _
        Optional ByVal a_stringFormatter As IStringFormatter, _
        Optional ByVal a_validationAdorner As IDynamicAdorner) As IPropertyBinding
    
    Dim p_localTargetProperty As String
    If VBA.IsMissing(a_targetProperty) Then
        p_localTargetProperty = InferTargetPropertyFromSource(a_source, a_outMode:=a_mode)
    Else
        p_localTargetProperty = a_targetProperty
    End If
    
    Dim p_bindingBase As PropertyBindingBase
    Set p_bindingBase = Factory.CreatePropertyBindingBase(a_context, a_source, a_target, _
            p_localTargetProperty, _
            a_mode:=a_mode, _
            a_updateSource:=a_updateSource, _
            a_converter:=a_converter, _
            a_stringFormatter:=a_stringFormatter, _
            a_validator:=a_validator, _
            a_validationAdorner:=a_validationAdorner)
    
    If Not a_validator Is Nothing Then
        p_bindingBase.AsINotifyValidationError.RegisterHandler Factory.ValidationManager
    End If
    
    Me.InjectBindingInfo p_bindingBase
    
    Set Initialize = Me
    
End Function

Friend Function InferTargetPropertyFromSource(ByVal a_source As IBindingPath, ByRef a_outMode As cc_isr_MVVM.BindingMode) As String
    
    Dim p_result As String
    Dim p_sourceValue As Variant
    
    If a_source.TryReadPropertyValue(a_outValue:=p_sourceValue) Then
        
        Select Case True
            
            Case VBA.VarType(p_sourceValue) = VBA.VbVarType.vbLong, _
                 VBA.VarType(p_sourceValue) = VBA.VbVarType.vbInteger, _
                 VBA.VarType(p_sourceValue) = VBA.VbVarType.vbDouble, _
                 VBA.VarType(p_sourceValue) = VBA.VbVarType.vbSingle, _
                 VBA.VarType(p_sourceValue) = VBA.VbVarType.vbByte, _
                p_result = "ListIndex"
            
            Case VBA.VarType(p_sourceValue) = VBA.VbVarType.vbString
                
                p_result = "Text"
            
            Case VBA.IsArray(p_sourceValue) Or VarType(p_sourceValue) = VBA.VbVarType.vbArray
                
                Debug.Print VBA.Information.TypeName(Me) & ": binding source [" & a_source.ToString() & "] one-way."
                a_outMode = cc_isr_MVVM.BindingMode.OneWayBinding
                p_result = "List"
            
            Case Else
                Debug.Print VBA.Information.TypeName(Me) & ": binding source [" & a_source.ToString() & "] one-way to source."
                a_outMode = cc_isr_MVVM.BindingMode.OneWayToSource
                p_result = BindingDefaults.ListBoxDefaultProperty
        End Select
    
    Else
        
        Debug.Print VBA.Information.TypeName(Me) & ": Could not read source property value; binding to default target property."
        p_result = BindingDefaults.ListBoxDefaultProperty
    
    End If
    InferTargetPropertyFromSource = p_result
End Function

Public Sub InjectBindingInfo(ByVal a_bindingInfo As PropertyBindingBase)
    
    cc_isr_Core.GuardClauses.GuardNullReference a_bindingInfo, VBA.Information.TypeName(Me)
    cc_isr_Core.GuardClauses.GuardDoubleInitialization This.Base, VBA.Information.TypeName(Me)
    cc_isr_Core.GuardClauses.GuardDoubleInitialization This.Handler, VBA.Information.TypeName(Me)
    Set This.Base = a_bindingInfo
    Set This.Handler = a_bindingInfo
    This.Base.AsIControlEvents.RegisterHandler Me

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

' - - - - - - - - - - - - - - - -
'
' HandleControlEvents implementation
'
' - - - - - - - - - - - - - - - -

Private Sub IHandleControlEvents_HandleAfterUpdate()
End Sub

Private Sub IHandleControlEvents_HandleBeforeUpdate(ByRef a_cancel As Boolean)
End Sub

Private Sub IHandleControlEvents_HandleChange()
    ' VF: added: required to capture Cut/Paste and Backspace/Delete in TextBox
End Sub

Private Sub IHandleControlEvents_HandleEnter()
End Sub

Private Sub IHandleControlEvents_HandleExit(ByRef a_cancel As Boolean)
End Sub

Private Sub IHandlePropertyChanged_HandlePropertyChanged(ByVal a_source As Object, ByVal a_propertyName As String)
    This.Handler.HandlePropertyChanged a_source, a_propertyName
End Sub

' - - - - - - - - - - - - - - - -
'
' PropertyBinding implementation
'
' - - - - - - - - - - - - - - - -

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
    IPropertyBinding_DefaultTargetProperty = BindingDefaults.ListBoxDefaultProperty
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

' - - - - - - - - - - - - - - - -
'
' TargetEventSource implementation
'
' - - - - - - - - - - - - - - - -

Private Sub TargetEventSource_Change()
    If This.Base.UpdateSourceTrigger = OnPropertyChanged Then This.Base.ApplyToSource
End Sub

