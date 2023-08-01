Attribute VB_Name = "ValidationManagerTests"
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Unit tests: Validation manager. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit
Option Private Module

Private Type ThisData
    
    ExpectedErrNumber As Long
    ExpectedErrSource As String
    ExpectedErrorCaught As Boolean
    
    Validator As cc_isr_MVVM.IValueValidator
    
    ConcreteSUT As ValidationManager
    NotifyValidationErrorSUT As INotifyValidationError
    HandleValidationErrorSUT As IHandleValidationError
    
    BindingManager As cc_isr_MVVM.IBindingManager
    BindingManagerStub As ITestStub
    
    CommandManager As cc_isr_MVVM.ICommandManager
    CommandManagerStub As ITestStub
    
    BindingSource As TestBindingObject
    BindingSourceStub As ITestStub
    BindingTarget As TestBindingObject
    BindingTargetStub As ITestStub
    
    SourcePropertyPath As String
    TargetPropertyPath As String
    Command As TestCommand
    
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
Public Sub BeforeEach()
    
    Dim p_bindingSource As New TestBindingObject
    
    Set This.ConcreteSUT = cc_isr_MVVM.Factory.CreateValidationManager(New TestNotifierFactory)
    Set This.NotifyValidationErrorSUT = This.ConcreteSUT
    Set This.HandleValidationErrorSUT = This.ConcreteSUT
    Set This.BindingSource = New TestBindingObject ' TestBindingObject.Create(This.ConcreteSUT)
    Set This.BindingSourceStub = This.BindingSource
    Set This.BindingTarget = New TestBindingObject ' TestBindingObject.Create(This.ConcreteSUT)
    Set This.BindingTargetStub = This.BindingTarget
    Set This.Command = New TestCommand
    Set This.CommandManager = New TestCommandManager
    Set This.CommandManagerStub = This.CommandManager
    Set This.Validator = New TestValueValidator
    Dim Manager As TestBindingManager
    Set Manager = New TestBindingManager
    Set This.BindingManager = Manager
    Set This.BindingManagerStub = This.BindingManager
    This.SourcePropertyPath = "TestStringProperty"
    This.TargetPropertyPath = "TestStringProperty"
    Set This.Assert = cc_isr_Test_Fx.Assert
End Sub

''' <summary>   Runs after each test. </summary>
Public Sub AfterEach()
    Set This.ConcreteSUT = Nothing
    Set This.NotifyValidationErrorSUT = Nothing
    Set This.HandleValidationErrorSUT = Nothing
    Set This.BindingSource = Nothing
    Set This.BindingTarget = Nothing
    Set This.Command = Nothing
    Set This.Validator = Nothing
    Set This.BindingManager = Nothing
    Set This.BindingManagerStub = Nothing
    This.SourcePropertyPath = vbNullString
    This.TargetPropertyPath = vbNullString
    This.ExpectedErrNumber = 0
    This.ExpectedErrorCaught = False
    This.ExpectedErrSource = vbNullString
    Set This.Assert = Nothing
End Sub

''' <summary>   Asserts the absence of an expected error. </summary>
Private Function AssertExpectError() As cc_isr_Test_Fx.Assert
    
    Dim p_message As String
    If Err.Number = This.ExpectedErrNumber Then
        If (This.ExpectedErrSource = vbNullString) Or (Err.Source = This.ExpectedErrSource) Then
            This.ExpectedErrorCaught = True
        Else
            p_message = "An error was raised, but not from the expected source. " & _
                      "Expected: '" & TypeName(This.ConcreteSUT) & "'; Actual: '" & Err.Source & "'."
        End If
    ElseIf Err.Number <> 0 Then
        p_message = "An error was raised, but not with the expected number. Expected: '" & _
                  This.ExpectedErrNumber & "'; Actual: '" & Err.Number & "'."
    Else
        p_message = "No error was raised."
    End If
    
    Set AssertExpectError = This.Assert.IsTrue(This.ExpectedErrorCaught, p_message)
    
End Function
