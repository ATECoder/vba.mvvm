Attribute VB_Name = "CommandManagerTests"
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Unit tests: Command manager. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit
' Option Private Module

Private Type ThisData

    ExpectedErrNumber As Long
    ExpectedErrSource As String
    ExpectedErrorCaught As Boolean
    
    ConcreteSUT As CommandManager
    AbstractSUT As cc_isr_MVVM.ICommandManager
    
    BindingContext As TestBindingObject
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
    Set This.ConcreteSUT = cc_isr_MVVM.Factory.NewCommandManager
    Set This.AbstractSUT = This.ConcreteSUT
    Set This.BindingContext = New TestBindingObject
    Set This.Command = New TestCommand
    Set Assert = cc_isr_Test_Fx.Assert
End Sub

''' <summary>   Runs after each test. </summary>
Public Sub AfterEach()
    Set This.ConcreteSUT = Nothing
    Set This.AbstractSUT = Nothing
    Set This.BindingContext = Nothing
    Set This.Command = Nothing
    Set Assert = Nothing
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

''' <summary>   Binds the test command to the target object. </summary>
''' <param name="a_progID">   The bound object program id. </param>
''' <param name="a_target">   The instance of the target ohect created from the <ParamRef name="a_progId"/> </param>
''' <returns>   <see cref="ICommandBinding"/>. </returns>
Private Function DefaultTargetCommandBindingFor(ByVal a_progID As String, ByRef a_target As Object) As ICommandBinding

    Dim p_outcome As cc_isr_Test_Fx.Assert

    Set a_target = VBA.CreateObject(a_progID)
    
    Set p_outcome = This.AbstractSUT.BindCommand(This.BindingContext, a_target, This.Command)
    
    If Not p_outcome.AssertSuccessful Then Debug.Print p_outcome.AssertMessage
    
    Set DefaultTargetCommandBindingFor = p_outcome
    
    
End Function

''' <summary>   Asserts binding a command to the object defined by the specified program id. </summary>
''' <param name="a_progID">   The bound object program id. </param>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Private Function AssertBindCommandBindsItem(ByVal a_progID As String) As cc_isr_Test_Fx.Assert
    
    Dim p_target As Object
    Dim p_commandBinding As ICommandBinding
    p_commandBinding = DefaultTargetCommandBindingFor(a_progID, p_target)
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    p_outcome = This.Assert.IsTrue(p_commandBinding.Command Is This.Command, _
            "The bound command should be the same as the expected command.")
            
    If p_outcome.AssertSuccessful Then
    
        p_outcome = This.Assert.IsTrue(p_commandBinding.Target Is p_target, _
                "The bound object should be the same as the expected object.")
    
    End If
    
    If Not p_outcome.AssertSuccessful Then Debug.Print p_outcome.AssertMessage
    
    Set AssertBindCommandBindsItem = p_outcome

End Function

''' <summary>   [Unit Test] Tests binding a command to a command button control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestBindCommandBindsCommandButton() As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Set p_outcome = AssertBindCommandBindsItem(cc_isr_MVVM.BindingDefaults.CommandButtonProgId)

    If Not p_outcome.AssertSuccessful Then Debug.Print p_outcome.AssertMessage

    Set TestBindCommandBindsCommandButton = p_outcome
    
End Function

''' <summary>   [Unit Test] Tests binding a command to a check box control. </summary>
Public Function TestBindCommandBindsCheckBox() As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Set p_outcome = AssertBindCommandBindsItem(cc_isr_MVVM.BindingDefaults.CheckBoxProgId)

    If Not p_outcome.AssertSuccessful Then Debug.Print p_outcome.AssertMessage

    Set TestBindCommandBindsCommandButton = p_outcome
    
End Function

''' <summary>   [Unit Test] Tests binding a command to an image control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestBindCommandBindsImage() As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Set p_outcome = AssertBindCommandBindsItem(cc_isr_MVVM.BindingDefaults.ImageProgId)

    If Not p_outcome.AssertSuccessful Then Debug.Print p_outcome.AssertMessage

    Set TestBindCommandBindsCommandButton = p_outcome
    
End Function

''' <summary>   [Unit Test] Tests binding a command to a label control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestBindCommandBindsLabel() As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert

    Set p_outcome = AssertBindCommandBindsItem(cc_isr_MVVM.BindingDefaults.LabelProgId)

    If Not p_outcome.AssertSuccessful Then Debug.Print p_outcome.AssertMessage

    Set TestBindCommandBindsCommandButton = p_outcome

End Function

