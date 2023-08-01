Attribute VB_Name = "ConstructionTests"
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Unit tests: Construction. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit

Private Type ThisData
    Assert As cc_isr_Test_Fx.Assert
End Type

Private This As ThisData

''' <summary>   Runs before all tests. </summary>
Public Sub BeforeAll()
End Sub

''' <summary>   Runs after all tests. </summary>
Public Sub AfterAll()
End Sub

''' <summary>   Runs before each test. </summary>
Private Sub BeforeEach()
    Set Assert = cc_isr_Test_Fx.Assert
End Sub

''' <summary>   Runs after each test. </summary>
Private Sub AfterEach()
    Set Assert = Nothing
End Sub

Public Function RunTest()
    BeforeEach
    TestAcceptCommandShouldConstruct
    AfterEach
End Function

''' <summary>   [Unit Test] Test constructing <see cref=""/> . </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestAcceptCommandShouldConstruct() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert

    On Error Resume Next

    Dim p_result As cc_isr_MVVM.AcceptCommand
    
    Set p_result = cc_isr_MVVM.Factory.NewAcceptCommand()
    
    Set p_outcome = This.Assert.AreEqual(0, Err.Number, "Error number " & CStr(Err.Number) & " should be 0.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.Assert.IsNotNull(p_result, TypeName(p_result) & " should not be null.")
    
    If Not p_outcome.AssertSuccessful Then Debug.Print p_outcome.AssertMessage

    Set TestAcceptCommandShouldConstruct = p_outcome
    
    On Error GoTo 0
    
End Function

