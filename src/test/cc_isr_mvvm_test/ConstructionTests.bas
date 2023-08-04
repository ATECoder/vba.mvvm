Attribute VB_Name = "ConstructionTests"
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Unit tests: Construction. </summary>
''' <remarks>
''' 2023-08-02: All tests passed.
''' </remarks>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit

Private Type ThisData
    Assert As cc_isr_Test_Fx.Assert
    TestView As TestView
End Type

Private This As ThisData

''' <summary>   Runs before all tests. </summary>
Public Sub BeforeAll()
End Sub

''' <summary>   Runs after all tests. </summary>
Public Sub AfterAll()
End Sub

''' <summary>   Runs before each test. </summary>
Public Sub BeforeEach()
    Set This.Assert = cc_isr_Test_Fx.Assert
    Set This.TestView = New TestView
End Sub

''' <summary>   Runs after each test. </summary>
Public Sub AfterEach()
    Set This.Assert = Nothing
    Set This.TestView = Nothing
End Sub

''' <summary>   [Unit Test] Test constructing a <see cref="cc_isr_MVVM.AcceptCommand"/> object. </summary>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>]. </returns>
Public Function TestAcceptCommandShouldConstruct() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert

    On Error Resume Next

    Dim p_result As cc_isr_MVVM.AcceptCommand
    
    Set p_result = cc_isr_MVVM.Factory.NewAcceptCommand().Initialize(This.TestView, cc_isr_MVVM.Factory.NewValidationManager)
    
    Set p_outcome = This.Assert.AreEqual(0, Err.Number, "Error number " & CStr(Err.Number) & " should be 0.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.Assert.IsNotNull(p_result, TypeName(p_result) & " should not be null.")
    
    On Error GoTo 0

    Debug.Print "TestAcceptCommandShouldConstruct " & _
        IIf(p_outcome.AssertSuccessful, "passed", "failed: " & p_outcome.AssertMessage)

    Set TestAcceptCommandShouldConstruct = p_outcome
    
    
End Function

