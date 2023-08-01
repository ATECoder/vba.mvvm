VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "An object that can stub an IView dependency in a unit test."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   An object that can stub an IView dependency in a unit test. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit
Implements ITestStub
Implements cc_isr_MVVM.IView
Implements cc_isr_MVVM.ICancellable

Private Type TState
    Cancel As Boolean
    Stub As TestStubBase
    Shown As Boolean
    ViewModel As Object
End Type

Private This As TState

''' <summary>   Gets a value indicating whether the stub IView is in a visible (shown) state or not. </summary>
Public Property Get Visible() As Boolean
Attribute Visible.VB_Description = "Gets a value indicating whether the stub IView is in a visible (shown) state or not."
    Visible = This.Shown
End Property

''' <summary>   Gets or sets the ViewModel object this stub is working with. </summary>
Public Property Get ViewModel() As Object
Attribute ViewModel.VB_Description = "Gets or sets the ViewModel object this stub is working with."
    Set ViewModel = This.ViewModel
End Property

Public Property Set ViewModel(ByVal a_value As Object)
    Set This.ViewModel = a_value
End Property

''' <summary>   Makes IView.ShowDialog return False (simulates a cancelled dialog). </summary>
Public Sub OnCancel()
Attribute OnCancel.VB_Description = "Makes IView.ShowDialog return False (simulates a cancelled dialog)."
    This.Cancel = True
    This.Shown = False
End Sub

Private Sub Class_Initialize()
    Set This.Stub = New TestStubBase
End Sub

Private Property Get ICancellable_IsCancelled() As Boolean
    ICancellable_IsCancelled = This.Cancel
End Property

Private Sub ICancellable_OnCancel()
    This.Stub.OnInvoke "OnCancel"
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

Private Sub IView_Hide()
    This.Stub.OnInvoke "Hide"
    This.Shown = False
End Sub

Private Sub IView_Show()
    This.Stub.OnInvoke "Show"
    This.Shown = True
End Sub

Private Function IView_ShowDialog() As Boolean
    This.Stub.OnInvoke "ShowDialog"
    This.Shown = Not This.Cancel
    IView_ShowDialog = Not This.Cancel
End Function

Private Property Get IView_ViewModel() As Object
    This.Stub.OnInvoke "ViewModel.Get"
    Set IView_ViewModel = This.ViewModel
End Property
