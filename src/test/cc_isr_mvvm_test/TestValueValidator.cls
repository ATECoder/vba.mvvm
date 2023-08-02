VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestValueValidator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Value validator tester. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit
Implements cc_isr_MVVM.IValueValidator
Implements ITestStub

Private Type TState
    Stub As TestStubBase
    IsValid As Boolean
    Message As String
End Type

Private This As TState

''' <summary>   Gets/sets a value that determines whether this validator evaluates to True or False. </summary>
Public Property Get IsValid() As Boolean
Attribute IsValid.VB_Description = "Gets/sets a value that determines whether this validator evaluates to True or False."
    IsValid = This.IsValid
End Property

Public Property Let IsValid(ByVal a_value As Boolean)
    This.IsValid = a_value
End Property

''' <summary>   Gets/sets the validation error message for this validator. </summary>
Public Property Get Message() As String
Attribute Message.VB_Description = "Gets/sets the validation error message for this validator."
    Message = This.Message
End Property

Public Property Let Message(ByVal a_value As String)
    This.Message = a_value
End Property

Private Sub Class_Initialize()
    Set This.Stub = New TestStubBase
End Sub

Private Function IValueValidator_IsValid(ByVal a_value As Variant, ByVal a_source As cc_isr_MVVM.IBindingPath, _
        ByVal a_target As cc_isr_MVVM.IBindingPath) As Boolean
    This.Stub.OnInvoke "IsValid"
    IValueValidator_IsValid = This.IsValid
End Function

Private Property Get IValueValidator_Message() As String
    This.Stub.OnInvoke "Message"
    IValueValidator_Message = This.Message
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

Private Property Get IValueValidator_Trigger() As cc_isr_MVVM.BindingUpdateSourceTrigger
    IValueValidator_Trigger = OnExit
End Property
