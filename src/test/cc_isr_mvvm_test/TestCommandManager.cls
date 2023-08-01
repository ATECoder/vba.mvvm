VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestCommandManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Command manager tester. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit
Implements cc_isr_MVVM.ICommandManager
Implements ITestStub

Private Type TState
    Stub As TestStubBase
End Type

Private This As TState

Private Sub Class_Initialize()
    Set This.Stub = New TestStubBase
End Sub

Private Function ICommandManager_BindCommand(ByVal a_bindingContext As Object, ByVal a_target As Object, _
        ByVal a_command As cc_isr_MVVM.ICommand) As cc_isr_MVVM.ICommandBinding
    This.Stub.OnInvoke "BindCommand"
End Function

Private Sub ICommandManager_EvaluateCanExecute(ByVal a_bindingContext As Object)
    This.Stub.OnInvoke "EvaluateCanExecute"
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


