Attribute VB_Name = "BindingPathTests"
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Unit tests: Binding path. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit
Option Private Module

Private Type ThisData
    
    ExpectedErrNumber As Long
    ExpectedErrSource As String
    ExpectedErrorCaught As Boolean
    
    ConcreteSUT As cc_isr_MVVM.BindingPath
    AbstractSUT As cc_isr_MVVM.IBindingPath
    
    BindingContext As TestBindingObject
    BindingSource As TestBindingObject
    PropertyName As String
    Path As String
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
    
    Dim p_context As TestBindingObject
    Set p_context = New TestBindingObject
    
    Set p_context.TestBindingObjectProperty = New TestBindingObject
    
    This.Path = "TestBindingObjectProperty.TestStringProperty"
    This.PropertyName = "TestStringProperty"
    Set This.BindingSource = p_context.TestBindingObjectProperty
    
    Set This.BindingContext = p_context
    Set This.ConcreteSUT = cc_isr_MVVM.Factory.CreateBindingPath(This.BindingContext, This.Path)
    Set This.AbstractSUT = This.ConcreteSUT
    Set This.Assert = cc_isr_Test_Fx.Assert

End Sub

''' <summary>   Runs after each test. </summary>
Public Sub AfterEach()
    
    Set This.ConcreteSUT = Nothing
    Set This.AbstractSUT = Nothing
    Set This.BindingSource = Nothing
    Set This.BindingContext = Nothing
    This.Path = vbNullString
    This.PropertyName = vbNullString
    This.ExpectedErrNumber = 0
    This.ExpectedErrorCaught = False
    This.ExpectedErrSource = vbNullString
    Set This.Assert = Nothing

End Sub

''' <summary>   Assert the absence of an expected error. </summary>
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

Public Function TestCreateGuardsNullBindingContext() As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    This.ExpectedErrNumber = cc_isr_Core.UserDefinedErrors.NullArgumentError.Code
    On Error Resume Next
    cc_isr_MVVM.Factory.CreateBindingPath Nothing, This.Path
    Set p_outcome = AssertExpectError
    On Error GoTo 0
    

    Debug.Print "TestCreateGuardsNullBindingContext " & _
        IIf(p_outcome.AssertSuccessful, "passed", "failed: " & p_outcome.AssertMessage)
    
    Set TestCreateGuardsNullBindingContext = p_outcome
    
End Function

Public Function TestCreateGuardsEmptyPath() As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    This.ExpectedErrNumber = cc_isr_Core.UserDefinedErrors.InvalidArgumentError.Code
    On Error Resume Next
    cc_isr_MVVM.Factory.CreateBindingPath This.BindingContext, vbNullString
    
    Set p_outcome = AssertExpectError
    
    On Error GoTo 0


    Debug.Print "TestCreateGuardsEmptyPath " & _
        IIf(p_outcome.AssertSuccessful, "passed", "failed: " & p_outcome.AssertMessage)

    Set TestCreateGuardsEmptyPath = p_outcome
    

End Function

Public Function TestCreateGuardsNonDefaultInstance() As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    This.ExpectedErrNumber = cc_isr_Core.UserDefinedErrors.InvalidOperationError.Code
    On Error Resume Next
    Dim p_bindingPath As BindingPath
    p_bindingPath = cc_isr_MVVM.Factory.CreateBindingPath(This.BindingContext, This.Path)
    Set p_outcome = AssertExpectError
    On Error GoTo 0


    Debug.Print "TestCreateGuardsNonDefaultInstance " & _
        IIf(p_outcome.AssertSuccessful, "passed", "failed: " & p_outcome.AssertMessage)

    Set TestCreateGuardsNonDefaultInstance = p_outcome

End Function

Public Function TestContextGuardsDefaultInstance() As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    This.ExpectedErrNumber = cc_isr_Core.UserDefinedErrors.InvalidOperationError.Code
    On Error Resume Next
    
    Dim p_bindingPath As BindingPath
    p_bindingPath = cc_isr_MVVM.Factory.CreateBindingPath(This.BindingContext, This.Path)
    ' Set BindingPath.Context = This.BindingContext
    
    Set p_outcome = AssertExpectError
    On Error GoTo 0


    Debug.Print "TestContextGuardsDefaultInstance " & _
        IIf(p_outcome.AssertSuccessful, "passed", "failed: " & p_outcome.AssertMessage)

    Set TestContextGuardsDefaultInstance = p_outcome
    
End Function

Public Function TestContextGuardsDoubleInitialization() As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    This.ExpectedErrNumber = cc_isr_Core.UserDefinedErrors.InvalidOperationError.Code
    On Error Resume Next
    Set This.ConcreteSUT.Context = New TestBindingObject
    Set p_outcome = AssertExpectError
    On Error GoTo 0


    Debug.Print "TestContextGuardsDoubleInitialization " & _
        IIf(p_outcome.AssertSuccessful, "passed", "failed: " & p_outcome.AssertMessage)
    
    Set TestContextGuardsDoubleInitialization = p_outcome

End Function

Public Function TestContextGuardsNullReference() As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    This.ExpectedErrNumber = cc_isr_Core.UserDefinedErrors.NullArgumentError.Code
    On Error Resume Next
    Set This.ConcreteSUT.Context = Nothing
    Set p_outcome = AssertExpectError
    On Error GoTo 0
    

    Debug.Print "TestContextGuardsNullReference " & _
        IIf(p_outcome.AssertSuccessful, "passed", "failed: " & p_outcome.AssertMessage)

    Set TestContextGuardsNullReference = p_outcome
    
End Function

Public Function TestPathGuardsDefaultInstance() As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    This.ExpectedErrNumber = cc_isr_Core.UserDefinedErrors.InvalidOperationError.Code
    On Error Resume Next
    
    Dim p_bindingPath As BindingPath
    p_bindingPath = cc_isr_MVVM.Factory.CreateBindingPath(This.BindingContext, This.Path)
    ' BindingPath.Path = This.Path
    
    Set p_outcome = AssertExpectError
    On Error GoTo 0


    Debug.Print "TestPathGuardsDefaultInstance " & _
        IIf(p_outcome.AssertSuccessful, "passed", "failed: " & p_outcome.AssertMessage)

    Set TestPathGuardsDefaultInstance = p_outcome

End Function

Public Function TestPathGuardsDoubleInitialization() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert
    This.ExpectedErrNumber = cc_isr_Core.UserDefinedErrors.InvalidOperationError.Code
    On Error Resume Next
    This.ConcreteSUT.Path = This.Path
    Set p_outcome = AssertExpectError
    On Error GoTo 0
    

    Debug.Print "TestPathGuardsDoubleInitialization " & _
        IIf(p_outcome.AssertSuccessful, "passed", "failed: " & p_outcome.AssertMessage)

    Set TestPathGuardsDoubleInitialization = p_outcome

End Function

Public Function TestPathGuardsEmptyString() As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    This.ExpectedErrNumber = cc_isr_Core.UserDefinedErrors.InvalidArgumentError.Code
    On Error Resume Next
    This.ConcreteSUT.Path = vbNullString
    Set p_outcome = AssertExpectError
    On Error GoTo 0


    Debug.Print "TestPathGuardsEmptyString " & _
        IIf(p_outcome.AssertSuccessful, "passed", "failed: " & p_outcome.AssertMessage)

    Set TestPathGuardsEmptyString = p_outcome

End Function

Public Function TestResolveSetsBindingSource() As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Dim p_bindingPath As BindingPath
    
    p_bindingPath = cc_isr_MVVM.Factory.NewBindingPath
    
    p_bindingPath.Path = This.Path
    Set p_bindingPath.Context = This.BindingContext
    
    Set p_outcome = This.Assert.IsFalse(p_bindingPath.Object Is Nothing, _
        "Object reference is unexpectedly set.")
        
    p_bindingPath.Resolve
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.Assert.IsTrue(This.BindingSource Is p_bindingPath.Object, _
                                            "The binding source should be set to an object.")


    Debug.Print "TestResolveSetsBindingSource " & _
        IIf(p_outcome.AssertSuccessful, "passed", "failed: " & p_outcome.AssertMessage)

    Set TestResolveSetsBindingSource = p_outcome

End Function

Public Function TestResolveSetsBindingPropertyName() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Dim p_bindingPath As BindingPath
    
    p_bindingPath = cc_isr_MVVM.Factory.NewBindingPath

    p_bindingPath.Path = This.Path
    
    Set p_bindingPath.Context = This.BindingContext
    
    Set p_outcome = This.Assert.IsFalse(p_bindingPath.PropertyName = VBA.vbNullString, _
                "PropertyName is unexpectedly non-empty.")
        
    p_bindingPath.Resolve
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.Assert.AreEqual(This.PropertyName, p_bindingPath.PropertyName, _
            "Propery name should equal the expected name")


    Debug.Print "TestResolveSetsBindingPropertyName " & _
        IIf(p_outcome.AssertSuccessful, "passed", "failed: " & p_outcome.AssertMessage)

    Set TestResolveSetsBindingPropertyName = p_outcome

End Function

Public Function TestCreateResolvesPropertyName() As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    Dim p_SUT As BindingPath
    Set p_SUT = cc_isr_MVVM.Factory.CreateBindingPath(This.BindingContext, This.Path)
    Set p_outcome = This.Assert.IsFalse(p_SUT.PropertyName = VBA.vbNullString, _
        "Property name should be empty.")


    Debug.Print "TestCreateResolvesPropertyName " & _
        IIf(p_outcome.AssertSuccessful, "passed", "failed: " & p_outcome.AssertMessage)

    Set TestCreateResolvesPropertyName = p_outcome

End Function

Public Function TestCreateResolvesBindingSource() As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    Dim p_SUT As BindingPath
    Set p_SUT = cc_isr_MVVM.Factory.CreateBindingPath(This.BindingContext, This.Path)
    
    Set p_outcome = This.Assert.IsNotNull(p_SUT.Object, _
            "The binding path object should not be nothing.")


    Debug.Print "TestCreateResolvesBindingSource " & _
        IIf(p_outcome.AssertSuccessful, "passed", "failed: " & p_outcome.AssertMessage)

    Set TestCreateResolvesBindingSource = p_outcome

End Function
