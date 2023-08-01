VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestBindingManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Binding manager tester. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit
Implements ITestStub
Implements cc_isr_MVVM.IBindingManager
Implements cc_isr_MVVM.IHandlePropertyChanged

Private Type TState
    Stub As TestStubBase
    Context As cc_isr_MVVM.IAppContext
    PropertyBindings As Collection
    NotifierFactory As cc_isr_MVVM.INotifierFactory
    StringFormatterFactory As cc_isr_MVVM.IStringFormatterFactory
    DebugOutput As Boolean
End Type

Private This As TState

Public Property Get NotifierFactory() As cc_isr_MVVM.INotifierFactory
    Set NotifierFactory = This.NotifierFactory
End Property

Private Sub Class_Initialize()
    Set This.Stub = New TestStubBase
    'Set This.Context = TestAppContext.Create
    Set This.NotifierFactory = New TestNotifierFactory
    Set This.PropertyBindings = New Collection
End Sub

Private Sub IBindingManager_Apply(ByVal a_source As Object)
    This.Stub.OnInvoke "Apply"
End Sub

'Private Function IBindingManager_BindPropertyPath(ByVal a_source As Object, ByVal a_propertyPath As String, ByVal a_target As Object, _
Optional ByVal a_targetProperty As String, _
Optional ByVal a_mode As cc_isr_MVVM.BindingMode = 0&, _
Optional ByVal a_updateTrigger As cc_isr_MVVM.BindingUpdateSourceTrigger = 1&, _
Optional ByVal a_validator As cc_isr_MVVM.IValueValidator, _
Optional ByVal a_converter As cc_isr_MVVM.IValueConverter, _
Optional ByVal a_formatString As String, _
Optional ByVal a_validationAdorner As cc_isr_MVVM.IDynamicAdorner) As IPropertyBinding
'VF: changed default a_updateTrigger due to revised enum BindingUpdateSourceTrigger


Private Function IBindingManager_BindPropertyPath(ByVal a_source As Object, ByVal a_propertyPath As String, _
        ByVal a_target As Object, _
        Optional ByVal a_targetProperty As String, _
        Optional ByVal a_mode As cc_isr_MVVM.BindingMode = 0&, _
        Optional ByVal a_updateTrigger As cc_isr_MVVM.BindingUpdateSourceTrigger = 2&, _
        Optional ByVal a_validator As cc_isr_MVVM.IValueValidator, _
        Optional ByVal a_converter As cc_isr_MVVM.IValueConverter, _
        Optional ByVal a_formatString As String, _
        Optional ByVal a_validationAdorner As cc_isr_MVVM.IDynamicAdorner) As cc_isr_MVVM.IPropertyBinding
        
    This.Stub.OnInvoke "BindPropertyPath"
    
    Dim p_testPropertyBinding As New TestPropertyBinding
    Dim Binding As IPropertyBinding
    Set Binding = p_testPropertyBinding.Initialize(This.Context, _
            cc_isr_MVVM.Factory.CreateBindingPath(a_source, a_propertyPath), a_target, _
                    a_targetProperty, a_mode, a_updateTrigger, a_validator, a_converter, _
                    This.StringFormatterFactory.Create(a_formatString), a_validationAdorner)
    This.PropertyBindings.Add Binding
    
    Set IBindingManager_BindPropertyPath = Binding
    
End Function

Private Property Get IBindingManager_DebugOutput() As Boolean
    This.Stub.OnInvoke "DebugOutput.Get"
    IBindingManager_DebugOutput = This.DebugOutput
End Property

Private Property Get IBindingManager_NotifierFactory() As cc_isr_MVVM.INotifierFactory
    This.Stub.OnInvoke "NotifierFactory.Get"
    Set IBindingManager_NotifierFactory = This.NotifierFactory
End Property

Private Property Get IBindingManager_StringFormatterFactory() As cc_isr_MVVM.IStringFormatterFactory
    This.Stub.OnInvoke "StringFormatterFactory.Get"
    Set IBindingManager_StringFormatterFactory = This.StringFormatterFactory
End Property

Private Sub IHandlePropertyChanged_HandlePropertyChanged(ByVal a_source As Object, ByVal a_propertyName As String)
    This.Stub.OnInvoke "OnPropertyChanged"
End Sub

Private Property Get ITestStub_MemberInvokes(ByVal a_memberName As String) As Long
    ITestStub_MemberInvokes = This.Stub.MemberInvokes(a_memberName)
End Property

Private Function ITestStub_ToString() As String
    ITestStub_ToString = This.Stub.ToString
End Function

Private Function ITestStub_Verify(ByVal a_assert As cc_isr_Test_Fx.Assert, ByVal a_memberName As String, _
                                  ByVal a_expectedInvokes As Long) As cc_isr_Test_Fx.Assert
    ITestStub_Verify = This.Stub.Verify(a_assert, a_memberName, a_expectedInvokes)
End Function

