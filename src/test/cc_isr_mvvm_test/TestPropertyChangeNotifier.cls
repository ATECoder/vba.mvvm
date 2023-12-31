VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestPropertyChangeNotifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Property change notifier tester. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit
Implements cc_isr_MVVM.INotifyPropertyChanged
Implements ITestStub

Private Type TState
    Handlers As Collection
    Stub As TestStubBase
End Type
Private This As TState

Public Property Get Handlers() As Collection
    Set Handlers = This.Handlers
End Property

Public Property Get Stub() As ITestStub
    Set Stub = This.Stub
End Property

Private Sub Class_Initialize()
    Set This.Handlers = New Collection
    Set This.Stub = New TestStubBase
End Sub

Private Sub INotifyPropertyChanged_OnPropertyChanged(ByVal a_source As Object, ByVal a_propertyName As String)
    This.Stub.OnInvoke "OnPropertyChanged"
    Dim Handler As cc_isr_MVVM.IHandlePropertyChanged
    For Each Handler In This.Handlers
        Handler.HandlePropertyChanged a_source, a_propertyName
    Next
End Sub

Private Sub INotifyPropertyChanged_RegisterHandler(ByVal a_handler As cc_isr_MVVM.IHandlePropertyChanged)
    This.Stub.OnInvoke "RegisterHandler"
    This.Handlers.Add a_handler
End Sub

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

