VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestPropertyBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "An object that can stub an IPropertyBinding dependency in a unit test."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   An object that can stub an IPropertyBinding dependency in a unit test. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit
Implements cc_isr_MVVM.IPropertyBinding
Implements cc_isr_MVVM.IHandlePropertyChanged
Implements ITestStub

Private Type TState
    Stub As TestStubBase
    Base As cc_isr_MVVM.PropertyBindingBase
    Handler As cc_isr_MVVM.IHandlePropertyChanged
End Type

Private This As TState

Public Function Initialize(ByVal a_context As IAppContext, ByVal a_source As cc_isr_MVVM.IBindingPath, ByVal a_target As Object, _
        Optional ByVal a_targetProperty As String = "Value", _
        Optional ByVal a_mode As cc_isr_MVVM.BindingMode = BindingMode.TwoWayBinding, _
        Optional ByVal a_updateSource As cc_isr_MVVM.BindingUpdateSourceTrigger = OnExit, _
        Optional ByVal a_validator As cc_isr_MVVM.IValueValidator, _
        Optional ByVal a_converter As cc_isr_MVVM.IValueConverter, _
        Optional ByVal a_stringFormatter As cc_isr_MVVM.IStringFormatter, _
        Optional ByVal a_validationAdorner As cc_isr_MVVM.IDynamicAdorner) As ITestStub
    
    Dim Base As cc_isr_MVVM.PropertyBindingBase
    Set Base = cc_isr_MVVM.Factory.NewPropertyBindingBase().Initialize(a_context, a_source, a_target, _
            a_targetProperty, a_mode, a_updateSource, a_converter, a_stringFormatter, a_validator, a_validationAdorner)
    Me.InjectBindingInfo Base
    Set Initialize = Me
End Function

Public Sub InjectBindingInfo(ByVal a_bindingInfo As cc_isr_MVVM.PropertyBindingBase)
    
    GuardClauses.GuardNullReference a_bindingInfo, TypeName(Me)
    GuardClauses.GuardDoubleInitialization This.Base, TypeName(Me)
    GuardClauses.GuardDoubleInitialization This.Handler, TypeName(Me)
    Set This.Base = a_bindingInfo
    Set This.Handler = a_bindingInfo

End Sub

Public Property Get AsIHandlePropertyChanged() As cc_isr_MVVM.IHandlePropertyChanged
    Set AsIHandlePropertyChanged = Me
End Property

Public Property Get Source() As cc_isr_MVVM.IBindingPath
    Set Source = This.Base.Source
End Property

Public Property Set Source(ByVal a_value As cc_isr_MVVM.IBindingPath)
    Set This.Base.Source = a_value
End Property

Public Property Get Target() As cc_isr_MVVM.IBindingPath
    Set Target = This.Base.Target
End Property

Public Property Set Target(ByVal a_value As cc_isr_MVVM.IBindingPath)
    Set This.Base.Target = a_value
End Property

Public Property Get Mode() As cc_isr_MVVM.BindingMode
    Mode = This.Base.Mode
End Property

Public Property Let Mode(ByVal a_value As cc_isr_MVVM.BindingMode)
    This.Base.Mode = a_value
End Property

Public Property Get UpdateSourceTrigger() As cc_isr_MVVM.BindingUpdateSourceTrigger
    UpdateSourceTrigger = This.Base.UpdateSourceTrigger
End Property

Public Property Let UpdateSourceTrigger(ByVal a_value As cc_isr_MVVM.BindingUpdateSourceTrigger)
    This.Base.UpdateSourceTrigger = a_value
End Property

Public Property Get Validator() As cc_isr_MVVM.IValueValidator
    Set Validator = This.Base.Validator
End Property

Public Property Set Validator(ByVal a_value As cc_isr_MVVM.IValueValidator)
    Set This.Base.Validator = a_value
End Property

Public Property Get ValidationAdorner() As cc_isr_MVVM.IDynamicAdorner
    Set ValidationAdorner = This.Base.ValidationAdorner
End Property

Public Property Set ValidationAdorner(ByVal a_value As cc_isr_MVVM.IDynamicAdorner)
    Set This.Base.ValidationAdorner = a_value
End Property

Public Property Get Converter() As cc_isr_MVVM.IValueConverter
    Set Converter = This.Base.Converter
End Property

Public Property Set Converter(ByVal a_value As cc_isr_MVVM.IValueConverter)
    Set This.Base.Converter = a_value
End Property

Private Sub Class_Initialize()
    Set This.Stub = New TestStubBase
End Sub

Private Sub IHandlePropertyChanged_HandlePropertyChanged(ByVal a_source As Object, ByVal a_propertyName As String)
    This.Stub.OnInvoke "HandlePropertyChanged"
    This.Handler.HandlePropertyChanged a_source, a_propertyName
End Sub

Private Property Get IPropertyBinding_CancelExitOnValidationError() As Boolean
    IPropertyBinding_CancelExitOnValidationError = This.Base.CancelExitOnValidationError
End Property

Private Property Get IPropertyBinding_DefaultTargetProperty() As String
    This.Stub.OnInvoke "DefaultTargetProperty.Get"
    IPropertyBinding_DefaultTargetProperty = vbNullString
End Property

Private Property Get IPropertyBinding_StringFormatter() As cc_isr_MVVM.IStringFormatter
    This.Stub.OnInvoke "StringFormatter.Get"
    Set IPropertyBinding_StringFormatter = This.Base.StringFormatter
End Property

Private Property Get IPropertyBinding_Validator() As cc_isr_MVVM.IValueValidator
    This.Stub.OnInvoke "Validator.Get"
    Set IPropertyBinding_Validator = This.Base.Validator
End Property

Private Property Get ITestStub_MemberInvokes(ByVal a_memberName As String) As Long
    ITestStub_MemberInvokes = This.Stub.MemberInvokes(a_memberName)
End Property

Private Function ITestStub_ToString() As String
    ITestStub_ToString = This.Stub.ToString
End Function

Private Function ITestStub_Verify(ByVal a_assert As cc_isr_Test_Fx.Assert, ByVal a_memberName As String, _
                                  ByVal a_expectedInvokes As Long) As cc_isr_Test_Fx.Assert
    Set ITestStub_Verify = This.Stub.Verify(a_assert, a_memberName, a_expectedInvokes)
End Function

Private Property Get IPropertyBinding_Mode() As cc_isr_MVVM.BindingMode
    This.Stub.OnInvoke "Mode.Get"
    IPropertyBinding_Mode = This.Base.Mode
End Property

Private Property Get IPropertyBinding_UpdateSourceTrigger() As cc_isr_MVVM.BindingUpdateSourceTrigger
    This.Stub.OnInvoke "UpdateSourceTrigger.Get"
    IPropertyBinding_UpdateSourceTrigger = This.Base.UpdateSourceTrigger
End Property

Private Property Get IPropertyBinding_Source() As cc_isr_MVVM.IBindingPath
    This.Stub.OnInvoke "Source.Get"
    Set IPropertyBinding_Source = This.Base.Source
End Property

Private Property Get IPropertyBinding_Target() As cc_isr_MVVM.IBindingPath
    This.Stub.OnInvoke "Target.Get"
    Set IPropertyBinding_Target = This.Base.Target
End Property

Private Property Get IPropertyBinding_Converter() As cc_isr_MVVM.IValueConverter
    This.Stub.OnInvoke "Converter.Get"
    Set IPropertyBinding_Converter = This.Base.Converter
End Property

Private Sub IPropertyBinding_Apply()
    This.Stub.OnInvoke "Apply"
End Sub

