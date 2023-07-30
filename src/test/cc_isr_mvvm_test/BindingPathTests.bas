Attribute VB_Name = "BindingPathTests"
'@Folder Tests
'@TestModule
Option Explicit
Option Private Module

Private Type ThisData
    
    ExpectedErrNumber As Long
    ExpectedErrSource As String
    ExpectedErrorCaught As Boolean
    
    ConcreteSUT As BindingPath
    AbstractSUT As IBindingPath
    
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
Private Sub BeforeEach()
    
    Dim Context As TestBindingObject
    Set Context = New TestBindingObject
    
    Set Context.TestBindingObjectProperty = New TestBindingObject
    
    This.Path = "TestBindingObjectProperty.TestStringProperty"
    This.PropertyName = "TestStringProperty"
    Set This.BindingSource = Context.TestBindingObjectProperty
    
    Set This.BindingContext = Context
    Set This.ConcreteSUT = BindingPath.Create(This.BindingContext, This.Path)
    Set This.AbstractSUT = This.ConcreteSUT
    Set Assert = cc_isr_Test_Fx.Assert

End Sub

''' <summary>   Runs after each test. </summary>
Private Sub AfterEach()
    
    Set This.ConcreteSUT = Nothing
    Set This.AbstractSUT = Nothing
    Set This.BindingSource = Nothing
    Set This.BindingContext = Nothing
    This.Path = vbNullString
    This.PropertyName = vbNullString
    This.ExpectedErrNumber = 0
    This.ExpectedErrorCaught = False
    This.ExpectedErrSource = vbNullString
    Set Assert = Nothing

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
    
    Dim outcome As cc_isr_Test_Fx.Assert
    This.ExpectedErrNumber = cc_isr_Core.UserDefinedErrors.NullArgumentError.Code
    On Error Resume Next
    BindingPath.Create Nothing, This.Path
    Set outcome = AssertExpectError
    On Error GoTo 0
    
    If Not outcome.AssertSuccessful Then Debug.Print outcome.AssertMessage
    
    Set TestCreateGuardsNullBindingContext = outcome
    
End Function

Public Function TestCreateGuardsEmptyPath() As cc_isr_Test_Fx.Assert
    
    Dim outcome As cc_isr_Test_Fx.Assert
    This.ExpectedErrNumber = cc_isr_Core.UserDefinedErrors.InvalidArgumentError.Code
    On Error Resume Next
    BindingPath.Create This.BindingContext, vbNullString
    
    Set outcome = AssertExpectError
    
    On Error GoTo 0

    If Not outcome.AssertSuccessful Then Debug.Print outcome.AssertMessage

    Set TestCreateGuardsEmptyPath = outcome
    

End Function

Public Function TestCreateGuardsNonDefaultInstance() As cc_isr_Test_Fx.Assert
    
    Dim outcome As cc_isr_Test_Fx.Assert
    This.ExpectedErrNumber = cc_isr_Core.UserDefinedErrors.InvalidOperationError.Code
    On Error Resume Next
    Dim p_bindingPath As BindingPath
    p_bindingPath = cc_isr_MVVM.Constructor.CreateBindingPath
    p_bindingPath.Create This.BindingContext, This.Path
    Set outcome = AssertExpectError
    On Error GoTo 0

    If Not outcome.AssertSuccessful Then Debug.Print outcome.AssertMessage

    Set TestCreateGuardsNonDefaultInstance = outcome

End Function

Public Function TestContextGuardsDefaultInstance() As cc_isr_Test_Fx.Assert
    
    Dim outcome As cc_isr_Test_Fx.Assert
    This.ExpectedErrNumber = cc_isr_Core.UserDefinedErrors.InvalidOperationError.Code
    On Error Resume Next
    Set BindingPath.Context = This.BindingContext
    Set outcome = AssertExpectError
    On Error GoTo 0

    If Not outcome.AssertSuccessful Then Debug.Print outcome.AssertMessage

    Set TestContextGuardsDefaultInstance = outcome
    
End Function

Public Function TestContextGuardsDoubleInitialization() As cc_isr_Test_Fx.Assert
    
    Dim outcome As cc_isr_Test_Fx.Assert
    This.ExpectedErrNumber = cc_isr_Core.UserDefinedErrors.InvalidOperationError.Code
    On Error Resume Next
    Set This.ConcreteSUT.Context = New TestBindingObject
    Set outcome = AssertExpectError
    On Error GoTo 0

    If Not outcome.AssertSuccessful Then Debug.Print outcome.AssertMessage
    
    Set TestContextGuardsDoubleInitialization = outcome

End Function

Public Function TestContextGuardsNullReference() As cc_isr_Test_Fx.Assert
    
    Dim outcome As cc_isr_Test_Fx.Assert
    This.ExpectedErrNumber = cc_isr_Core.UserDefinedErrors.NullArgumentError.Code
    On Error Resume Next
    Set This.ConcreteSUT.Context = Nothing
    Set outcome = AssertExpectError
    On Error GoTo 0
    
    If Not outcome.AssertSuccessful Then Debug.Print outcome.AssertMessage

    Set TestContextGuardsNullReference = outcome
    
End Function

Public Function TestPathGuardsDefaultInstance() As cc_isr_Test_Fx.Assert
    
    Dim outcome As cc_isr_Test_Fx.Assert
    This.ExpectedErrNumber = cc_isr_Core.UserDefinedErrors.InvalidOperationError.Code
    On Error Resume Next
    BindingPath.Path = This.Path
    Set outcome = AssertExpectError
    On Error GoTo 0

    If Not outcome.AssertSuccessful Then Debug.Print outcome.AssertMessage

    Set TestPathGuardsDefaultInstance = outcome

End Function

Public Function TestPathGuardsDoubleInitialization() As cc_isr_Test_Fx.Assert

    Dim outcome As cc_isr_Test_Fx.Assert
    This.ExpectedErrNumber = cc_isr_Core.UserDefinedErrors.InvalidOperationError.Code
    On Error Resume Next
    This.ConcreteSUT.Path = This.Path
    Set outcome = AssertExpectError
    On Error GoTo 0
    
    If Not outcome.AssertSuccessful Then Debug.Print outcome.AssertMessage

    Set TestPathGuardsDoubleInitialization = outcome

End Function

Public Function TestPathGuardsEmptyString() As cc_isr_Test_Fx.Assert
    
    Dim outcome As cc_isr_Test_Fx.Assert
    This.ExpectedErrNumber = cc_isr_Core.UserDefinedErrors.InvalidArgumentError.Code
    On Error Resume Next
    This.ConcreteSUT.Path = vbNullString
    Set outcome = AssertExpectError
    On Error GoTo 0

    If Not outcome.AssertSuccessful Then Debug.Print outcome.AssertMessage

    Set TestPathGuardsEmptyString = outcome

End Function

Public Function TestResolveSetsBindingSource() As cc_isr_Test_Fx.Assert
    
    Dim outcome As cc_isr_Test_Fx.Assert
    
    Dim p_bindingPath As BindingPath
    
    p_bindingPath = cc_isr_MVVM.Constructor.CreateBindingPath
    
    p_bindingPath.Path = This.Path
    Set p_bindingPath.Context = This.BindingContext
    
    Set outcome = This.Assert.IsFalse(p_bindingPath.Object Is Nothing, _
        "Object reference is unexpectedly set.")
        
    p_bindingPath.Resolve
    
    If outcome.AssertSuccessful Then _
        Set outcome = This.Assert.IsTrue(This.BindingSource Is p_bindingPath.Object, _
                                            "The binding source should be set to an object.")

    If Not outcome.AssertSuccessful Then Debug.Print outcome.AssertMessage

    Set TestResolveSetsBindingSource = outcome

End Function

Public Function TestResolveSetsBindingPropertyName() As cc_isr_Test_Fx.Assert

    Dim outcome As cc_isr_Test_Fx.Assert
    
    Dim p_bindingPath As BindingPath
    
    p_bindingPath = cc_isr_MVVM.Constructor.CreateBindingPath

    p_bindingPath.Path = This.Path
    
    Set p_bindingPath.Context = This.BindingContext
    
    Set outcome = This.Assert.IsFalse(p_bindingPath.PropertyName = VBA.vbNullString, _
                "PropertyName is unexpectedly non-empty.")
        
    p_bindingPath.Resolve
    
    If outcome.AssertSuccessful Then _
        Set outcome = This.Assert.AreEqual(This.PropertyName, p_bindingPath.PropertyName, _
            "Propery name should equal the expected name")

    If Not outcome.AssertSuccessful Then Debug.Print outcome.AssertMessage

    Set TestResolveSetsBindingPropertyName = outcome

End Function

Public Function TestCreateResolvesPropertyName() As cc_isr_Test_Fx.Assert
    
    Dim outcome As cc_isr_Test_Fx.Assert
    Dim p_SUT As BindingPath
    Set p_SUT = BindingPath.Create(This.BindingContext, This.Path)
    Set outcome = This.Assert.IsFalse(p_SUT.PropertyName = VBA.vbNullString, _
        "Property name should be empty.")

    If Not outcome.AssertSuccessful Then Debug.Print outcome.AssertMessage

    Set TestCreateResolvesPropertyName = outcome

End Function

Public Function TestCreateResolvesBindingSource() As cc_isr_Test_Fx.Assert
    
    Dim outcome As cc_isr_Test_Fx.Assert
    Dim p_SUT As BindingPath
    Set p_SUT = BindingPath.Create(This.BindingContext, This.Path)
    
    Set outcome = This.Assert.IsNotNull(p_SUT.Object, _
            "The binding path object should not be nothing.")

    If Not outcome.AssertSuccessful Then Debug.Print outcome.AssertMessage

    Set TestCreateResolvesBindingSource = outcome

End Function
