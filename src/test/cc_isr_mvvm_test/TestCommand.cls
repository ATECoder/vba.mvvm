VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "An object that can stub an ICommand dependency in a unit test."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   An object that can stub an ICommand dependency in a unit test. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit
Implements cc_isr_MVVM.ICommand
Implements ITestStub

Private Const CanExecuteDefaultValue As Boolean = True

Private Type TState
    Stub As TestStubBase
    CanExecute As Boolean
End Type

Private This As TState

''' <summary>   Controls the return value of ICommand.CanExecute. True by default. </summary>
Public Property Get CanExecute() As Boolean
Attribute CanExecute.VB_Description = "Controls the return value of ICommand.CanExecute. True by default."
    CanExecute = This.CanExecute
End Property

Public Property Let CanExecute(ByVal a_value As Boolean)
    This.CanExecute = a_value
End Property

Private Sub Class_Initialize()
    Set This.Stub = New TestStubBase
    This.CanExecute = CanExecuteDefaultValue
End Sub

Private Function ICommand_CanExecute(ByVal a_context As Object) As Boolean
    This.Stub.OnInvoke "CanExecute"
    ICommand_CanExecute = This.CanExecute
End Function

Private Property Get ICommand_Description() As String
    This.Stub.OnInvoke "Description.Get"
    ICommand_Description = "A " & TypeName(Me) & " instance tracks the number of times its members are invoked."
End Property

Private Sub ICommand_Execute(ByVal a_context As Object)
    This.Stub.OnInvoke "Execute"
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
