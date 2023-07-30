Attribute VB_Name = "BindingManagerTests"
'@Folder Tests
'@TestModule
Option Explicit
' Option Private Module

Private Type ThisData

    ExpectedErrNumber As Long
    ExpectedErrSource As String
    ExpectedErrorCaught As Boolean
    
    ConcreteSUT As BindingManager
    AbstractSUT As IBindingManager
    HandlePropertyChangedSUT As IHandlePropertyChanged
    
    CommandManager As ICommandManager
    CommandManagerStub As ITestStub
    
    BindingSource As TestBindingObject
    BindingTarget As TestBindingObject
    
    SourceProperty As String
    TargetProperty As String
    
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
Private Sub BeforeEach()
    Set This.CommandManager = New TestCommandManager
    Set This.CommandManagerStub = This.CommandManager
    Set This.ConcreteSUT = BindingManager.Create(This.CommandManager, cc_isr_MVVM.Constructor.CreateStringFormatterFactory)
    Set This.AbstractSUT = This.ConcreteSUT
    Set This.HandlePropertyChangedSUT = This.ConcreteSUT
    Set This.BindingSource = New TestBindingObject
    Set This.BindingTarget = New TestBindingObject
    Set This.Command = New TestCommand
    This.SourcePropertyPath = "TestStringProperty"
    This.TargetPropertyPath = "TestStringProperty"
    This.SourceProperty = "TestStringProperty"
    This.TargetProperty = "TestStringProperty"
    Set Assert = cc_isr_Test_Fx.Assert
End Sub

''' <summary>   Runs after each test. </summary>
Private Sub AfterEach()
    Set This.ConcreteSUT = Nothing
    Set This.AbstractSUT = Nothing
    Set This.HandlePropertyChangedSUT = Nothing
    Set This.BindingSource = Nothing
    Set This.BindingTarget = Nothing
    Set This.Command = Nothing
    This.SourcePropertyPath = vbNullString
    This.TargetPropertyPath = vbNullString
    This.ExpectedErrNumber = 0
    This.ExpectedErrorCaught = False
    This.ExpectedErrSource = vbNullString
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

''' <summary>   [Unit Test] Tests creating guards for non default instance. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestCreateGuardsNonDefaultInstance() As cc_isr_Test_Fx.Assert

    Dim outcome As cc_isr_Test_Fx.Assert
    This.ExpectedErrNumber = cc_isr_Core.UserDefinedErrors.InvalidOperationError.Code
    Dim p_bindingManger As cc_isr_MVVM.BindingManager
    p_bindingManger = cc_isr_MVVM.Constructor.CreateBindingManager
    
    On Error Resume Next
    '@Ignore FunctionReturnValueDiscarded, FunctionReturnValueNotUsed
    p_bindingManger.Create This.CommandManager, cc_isr_MVVM.Constructor.CreateStringFormatterFactory
    Set outcome = AssertExpectError
    On Error GoTo 0
    
    If Not outcome.AssertSuccessful Then Debug.Print outcome.AssertMessage
    
    Set TestCreateGuardsNonDefaultInstance = outcome

End Function

''' <summary>   Binds the binding source to the target object. </summary>
''' <param name="a_progID">   The bound object program id. </param>
''' <param name="a_target">   The instance of the target ohect created from the <ParamRef name="a_progId"/> </param>
''' <returns>   <see cref="IPropertyBinding"/>. </returns>
Private Function DefaultPropertyPathBindingFor(ByVal a_progID As String, ByRef a_target As Object) As IPropertyBinding

    Set a_target = VBA.CreateObject(a_progID)
    
    Dim outcome As cc_isr_Test_Fx.Assert
    Set outcome = This.AbstractSUT.BindPropertyPath(This.BindingSource, _
                    This.SourcePropertyPath, a_target, This.TargetPropertyPath)
    
    If Not outcome.AssertSuccessful Then Debug.Print outcome.AssertMessage
    
    Set DefaultPropertyPathBindingFor = outcome

End Function

''' <summary>   Asserts creating a property path binding to the object defined by the specified program id. </summary>
''' <param name="a_progID">   The bound object program id. </param>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Private Function AssertCreatePropertyBindingType(ByVal a_progID As String, _
    ByVal a_type As Object) As cc_isr_Test_Fx.Assert
    
    Dim outcome As cc_isr_Test_Fx.Assert
   
    Dim p_target As Object
    
    Dim p_result As IPropertyBinding
    Set p_result = DefaultPropertyPathBindingFor(a_progID, p_target)
        
    Set outcome = This.Assert.IsTrue(TypeOf p_result Is CheckBoxPropertyBinding, _
                                 "Property type should equal expected type.")
    
    If Not outcome.AssertSuccessful Then Debug.Print outcome.AssertMessage
    
    Set AssertCreatePropertyBindingType = outcome

End Function

''' <summary>   Asserts creating a property path binding to the object defined by the specified program id. </summary>
''' <param name="a_progID">   The bound object program id. </param>
''' <param name="a_targetProeprtyPath">   [Optional, String, Null] The target property path. </param>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Private Function AssertCreatePropertyBindingType2(ByVal a_progID As String, _
        ByVal a_type As Object, _
        Optional ByVal a_targetProeprtyPath As String = VBA.vbNullString) As cc_isr_Test_Fx.Assert
    
    This.TargetPropertyPath = a_targetProeprtyPath
    
    Dim outcome As cc_isr_Test_Fx.Assert
   
    Dim p_target As Object
    
    Dim p_result As IPropertyBinding
    Set p_result = DefaultPropertyPathBindingFor(a_progID, p_target)
        
    Set outcome = This.Assert.IsTrue(TypeOf p_result Is CheckBoxPropertyBinding, _
                                 "Property type should equal expected type.")
    
    If Not outcome.AssertSuccessful Then Debug.Print outcome.AssertMessage
    
    Set AssertCreatePropertyBindingType2 = outcome

End Function


''' <summary>   Asserts creating a property path binding to the object defined by the specified program id. </summary>
''' <param name="a_progID">               The bound object program id. </param>
''' <param name="a_name">                 The expected property name. </param>
''' <param name="a_targetProeprtyPath">   [Optional, String, Null] The target property path. </param>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Private Function AssertCreatePropertyBindingName(ByVal a_progID As String, _
        ByVal a_name As String, _
        Optional ByVal a_targetProeprtyPath As String = VBA.vbNullString) As cc_isr_Test_Fx.Assert
    
    Dim outcome As cc_isr_Test_Fx.Assert
   
    This.TargetPropertyPath = a_targetProeprtyPath
   
    Dim p_target As Object
    
    Dim p_result As IPropertyBinding
    Set p_result = DefaultPropertyPathBindingFor(a_progID, p_target)
        
       
    Set outcome = This.Assert.AreEqual(a_name, p_result.Target.PropertyName, _
                                   "Property name should equal expected value.")
    
    If Not outcome.AssertSuccessful Then Debug.Print outcome.AssertMessage
    
    Set AssertCreatePropertyBindingName = outcome

End Function

''' <summary>   [Unit Test] Test creating a property path binding for a check box control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestCheckBoxTargetCreatesPropertyBinding() As cc_isr_Test_Fx.Assert

    Dim p_result As CheckBoxPropertyBinding
    
    Dim outcome As cc_isr_Test_Fx.Assert

    Set outcome = AssertCreatePropertyBindingType(FormsProgID.CheckBoxProgId, p_result)
    
    If Not outcome.AssertSuccessful Then Debug.Print outcome.AssertMessage

    Set TestCheckBoxTargetCreatesPropertyBinding = outcome
    
End Function

''' <summary>   [Unit Test] Test creating a property path binding value for a check box control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestCheckBoxTargetBindsValueByDefault() As cc_isr_Test_Fx.Assert

    Dim outcome As cc_isr_Test_Fx.Assert

    Set outcome = AssertCreatePropertyBindingName(FormsProgID.CheckBoxProgId, "Value")

    If Not outcome.AssertSuccessful Then Debug.Print outcome.AssertMessage

    Set TestCheckBoxTargetBindsValueByDefault = outcome

End Function

''' <summary>   [Unit Test] Test creating a property path binding for a check box control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestComboBoxTargetCreatesPropertyBinding() As cc_isr_Test_Fx.Assert

    Dim p_result As ComboBoxPropertyBinding
    
    Dim outcome As cc_isr_Test_Fx.Assert
    
    Set outcome = AssertCreatePropertyBindingType(FormsProgID.ComboBoxProgId, p_result)
                                                                    
    If Not outcome.AssertSuccessful Then Debug.Print outcome.AssertMessage
    
    Set TestComboBoxTargetCreatesPropertyBinding = outcome
                                                                    
End Function

''' <summary>   [Unit Test] Test creating a property path binding value for a combo box control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestComboBoxTargetBindsValueByDefault() As cc_isr_Test_Fx.Assert

    Dim outcome As cc_isr_Test_Fx.Assert
    
    Set outcome = AssertCreatePropertyBindingName(FormsProgID.ComboBoxProgId, "Value")

    If Not outcome.AssertSuccessful Then Debug.Print outcome.AssertMessage
    
    Set TestComboBoxTargetBindsValueByDefault = outcome

End Function

''' <summary>   [Unit Test] Test creating a property path binding for a list box control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestListBoxTargetCreatesPropertyBinding() As cc_isr_Test_Fx.Assert

    Dim p_result As ListBoxPropertyBinding
    
    Dim outcome As cc_isr_Test_Fx.Assert
    
    Set outcome = AssertCreatePropertyBindingType(FormsProgID.ListBoxProgId, p_result)
                                                                    
    If Not outcome.AssertSuccessful Then Debug.Print outcome.AssertMessage
    
    Set TestListBoxTargetCreatesPropertyBinding = outcome

End Function

''' <summary>   [Unit Test] Test creating a property path binding value for a list box control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestListBoxTargetBindsValueByDefault() As cc_isr_Test_Fx.Assert

    Dim outcome As cc_isr_Test_Fx.Assert
    
    Set outcome = AssertCreatePropertyBindingName(FormsProgID.ListBoxProgId, "Value")

    If Not outcome.AssertSuccessful Then Debug.Print outcome.AssertMessage
    
    Set TestListBoxTargetBindsValueByDefault = outcome

End Function

''' <summary>   [Unit Test] Test creating a property path binding for a Multi Page control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestMultiPageTargetCreatesPropertyBinding() As cc_isr_Test_Fx.Assert

    Dim p_result As MultiPagePropertyBinding
    
    Dim outcome As cc_isr_Test_Fx.Assert
    
    Set outcome = AssertCreatePropertyBindingType(FormsProgID.MultiPageProgId, p_result)
                                                                    
    If Not outcome.AssertSuccessful Then Debug.Print outcome.AssertMessage
    
    Set TestMultiPageTargetCreatesPropertyBinding = outcome

End Function

''' <summary>   [Unit Test] Test creating a property path binding value for a Multi Page control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestMultiPageTargetBindsValueByDefault() As cc_isr_Test_Fx.Assert

    Dim outcome As cc_isr_Test_Fx.Assert
    
    Set outcome = AssertCreatePropertyBindingName(FormsProgID.MultiPageProgId, "Value")

    If Not outcome.AssertSuccessful Then Debug.Print outcome.AssertMessage
    
    Set TestMultiPageTargetBindsValueByDefault = outcome

End Function

''' <summary>   [Unit Test] Test creating a property path binding for a Option Button control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestOptionButtonTargetCreatesPropertyBinding() As cc_isr_Test_Fx.Assert

    Dim p_result As OptionButtonPropertyBinding
    
    Dim outcome As cc_isr_Test_Fx.Assert
    
    Set outcome = AssertCreatePropertyBindingType(FormsProgID.OptionButtonProgId, p_result)
                                                                    
    If Not outcome.AssertSuccessful Then Debug.Print outcome.AssertMessage
    
    Set TestOptionButtonTargetCreatesPropertyBinding = outcome

End Function

''' <summary>   [Unit Test] Test creating a property path binding value for a Option Button control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestOptionButtonTargetBindsValueByDefault() As cc_isr_Test_Fx.Assert

    Dim outcome As cc_isr_Test_Fx.Assert
    
    Set outcome = AssertCreatePropertyBindingName(FormsProgID.OptionButtonProgId, "Value")

    If Not outcome.AssertSuccessful Then Debug.Print outcome.AssertMessage
    
    Set TestOptionButtonTargetBindsValueByDefault = outcome

End Function

''' <summary>   [Unit Test] Test creating a property path binding for a Scroll Bar control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestScrollBarTargetCreatesPropertyBinding() As cc_isr_Test_Fx.Assert

    Dim p_result As ScrollBarPropertyBinding
    
    Dim outcome As cc_isr_Test_Fx.Assert
    
    Set outcome = AssertCreatePropertyBindingType(FormsProgID.ScrollBarProgId, p_result)
                                                                    
    If Not outcome.AssertSuccessful Then Debug.Print outcome.AssertMessage
    
    Set TestScrollBarTargetCreatesPropertyBinding = outcome

End Function

''' <summary>   [Unit Test] Test creating a property path binding value for a Scroll Bar control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestScrollBarTargetBindsValueByDefault() As cc_isr_Test_Fx.Assert

    Dim outcome As cc_isr_Test_Fx.Assert
    
    Set outcome = AssertCreatePropertyBindingName(FormsProgID.ScrollBarProgId, "Value")

    If Not outcome.AssertSuccessful Then Debug.Print outcome.AssertMessage
    
    Set TestScrollBarTargetBindsValueByDefault = outcome

End Function

''' <summary>   [Unit Test] Test creating a property path binding for a Spin Button control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestSpinButtonTargetCreatesPropertyBinding() As cc_isr_Test_Fx.Assert

    Dim p_result As SpinButtonPropertyBinding
    
    Dim outcome As cc_isr_Test_Fx.Assert
    
    Set outcome = AssertCreatePropertyBindingType(FormsProgID.SpinButtonProgId, p_result)
                                                                    
    If Not outcome.AssertSuccessful Then Debug.Print outcome.AssertMessage
    
    Set TestSpinButtonTargetCreatesPropertyBinding = outcome

End Function

''' <summary>   [Unit Test] Test creating a property path binding value for a Spin Button control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestSpinButtonTargetBindsValueByDefault() As cc_isr_Test_Fx.Assert

    Dim outcome As cc_isr_Test_Fx.Assert
    
    Set outcome = AssertCreatePropertyBindingName(FormsProgID.SpinButtonProgId, "Value")

    If Not outcome.AssertSuccessful Then Debug.Print outcome.AssertMessage
    
    Set TestSpinButtonTargetBindsValueByDefault = outcome

End Function

''' <summary>   [Unit Test] Test creating a property path binding for a Tab Strip control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestTabStripTargetCreatesPropertyBinding() As cc_isr_Test_Fx.Assert

    Dim p_result As TabStripPropertyBinding
    
    Dim outcome As cc_isr_Test_Fx.Assert
    
    Set outcome = AssertCreatePropertyBindingType(FormsProgID.TabStripProgId, p_result)
                                                                    
    If Not outcome.AssertSuccessful Then Debug.Print outcome.AssertMessage
    
    Set TestTabStripTargetCreatesPropertyBinding = outcome
                                                                    
End Function

''' <summary>   [Unit Test] Test creating a property path binding value for a Tab Strip control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestTabStripTargetBindsValueByDefault() As cc_isr_Test_Fx.Assert

    Dim outcome As cc_isr_Test_Fx.Assert
    
    Set outcome = AssertCreatePropertyBindingName(FormsProgID.TabStripProgId, "Value")

    If Not outcome.AssertSuccessful Then Debug.Print outcome.AssertMessage
    
    Set TestTabStripTargetBindsValueByDefault = outcome

End Function

''' <summary>   [Unit Test] Test creating a property path binding for a Text Box control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestTextBoxTargetCreatesPropertyBinding() As cc_isr_Test_Fx.Assert

    Dim p_result As TextBoxPropertyBinding
    
    Dim outcome As cc_isr_Test_Fx.Assert
    
    Set outcome = AssertCreatePropertyBindingType(FormsProgID.TextBoxProgId, p_result)
    
    If Not outcome.AssertSuccessful Then Debug.Print outcome.AssertMessage
    
    Set TestTextBoxTargetCreatesPropertyBinding = outcome
                                                                    
End Function

''' <summary>   [Unit Test] Test creating a property path binding value for a Text Box control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestTextBoxTargetBindsValueByDefault() As cc_isr_Test_Fx.Assert

    Dim outcome As cc_isr_Test_Fx.Assert
    
    Set outcome = AssertCreatePropertyBindingName(FormsProgID.TextBoxProgId, "Text")
    
    If Not outcome.AssertSuccessful Then Debug.Print outcome.AssertMessage
    
    Set TestCheckBoxTargetBindsValueByDefault = outcome

End Function

''' <summary>   [Unit Test] Test creating a property path binding for a Frame control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestFrameTargetCreatesOneWayBindingWithNonDefaultTarget() As cc_isr_Test_Fx.Assert
    
    Dim outcome As cc_isr_Test_Fx.Assert
    
    This.TargetPropertyPath = "Font.Bold"
   
    Dim p_target As Object
    
    Dim p_result As IPropertyBinding
    Set p_result = DefaultPropertyPathBindingFor(FormsProgID.FrameProgId, p_target)
        
    Set outcome = This.Assert.AreEqual(TypeName(OneWayPropertyBinding), TypeName(p_result), _
                                   "Property type name should equal expected type name.")
    
    If Not outcome.AssertSuccessful Then Debug.Print outcome.AssertMessage
    
    Set TestFrameTargetCreatesOneWayBindingWithNonDefaultTarget = outcome
    
    
End Function

''' <summary>   [Unit Test] Test creating a property path binding for a label control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestLabelTargetCreatesOneWayBindingWithNonDefaultTarget() As cc_isr_Test_Fx.Assert

    Dim p_result As OneWayPropertyBinding
    
    Dim outcome As cc_isr_Test_Fx.Assert
    
    Set outcome = AssertCreatePropertyBindingType2(FormsProgID.LabelProgId, _
                                                    p_result, "Font.Bold")
                                                                    
    If Not outcome.AssertSuccessful Then Debug.Print outcome.AssertMessage
    
    Set TestLabelTargetCreatesOneWayBindingWithNonDefaultTarget = outcome

End Function

''' <summary>   [Unit Test] Test creating a property path binding for a frame control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestFrameTargetBindsCaptionPropertyByDefault() As cc_isr_Test_Fx.Assert

    Dim p_result As CaptionPropertyBinding
    
    Dim outcome As cc_isr_Test_Fx.Assert
    
    Set outcome = AssertCreatePropertyBindingType(FormsProgID.FrameProgId, p_result)
                                                                    
    If Not outcome.AssertSuccessful Then Debug.Print outcome.AssertMessage
    
    Set TestFrameTargetBindsCaptionPropertyByDefault = outcome

End Function

''' <summary>   [Unit Test] Test creating a property path binding for a label control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestLabelTargetBindsCaptionPropertyByDefault() As cc_isr_Test_Fx.Assert

    Dim p_result As CaptionPropertyBinding
    
    Dim outcome As cc_isr_Test_Fx.Assert
    
    Set outcome = AssertCreatePropertyBindingType(FormsProgID.LabelProgId, p_result)
                                                                    
    If Not outcome.AssertSuccessful Then Debug.Print outcome.AssertMessage
    
    Set TestLabelTargetBindsCaptionPropertyByDefault = outcome

End Function


''' <summary>   [Unit Test] Test non control target creates one way binding. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestNonControlTargetCreatesOneWayBinding() As cc_isr_Test_Fx.Assert
    
    Dim p_result As IPropertyBinding
    
    Set p_result = This.AbstractSUT.BindPropertyPath(This.BindingSource, This.SourcePropertyPath, _
                                                   This.BindingTarget, This.TargetPropertyPath)
    
    Dim outcome As cc_isr_Test_Fx.Assert
    
    Set outcome = This.Assert.IsTrue(TypeOf p_result Is OneWayPropertyBinding, _
                                        "Type of result should equal expected type.")
    
    If Not outcome.AssertSuccessful Then Debug.Print outcome.AssertMessage
    
    Set TestNonControlTargetCreatesOneWayBinding = outcome

End Function

''' <summary>   [Unit Test] Test non control target requires target property path. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestNonControlTargetRequiresTargetPropertyPath() As cc_isr_Test_Fx.Assert

    Dim outcome As cc_isr_Test_Fx.Assert
    This.ExpectedErrNumber = cc_isr_Core.UserDefinedErrors.InvalidArgumentError.Code
    On Error Resume Next
    This.AbstractSUT.BindPropertyPath _
        This.BindingSource, _
        This.SourcePropertyPath, _
        This.BindingTarget, _
        TargetProperty:=vbNullString
    Set outcome = AssertExpectError
    On Error GoTo 0
    
    If Not outcome.AssertSuccessful Then Debug.Print outcome.AssertMessage
    
    Set TestNonControlTargetRequiresTargetPropertyPath = outcome

End Function

''' <summary>   [Unit Test] Test . </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestAddsToPropertyBindingsCollection() As cc_isr_Test_Fx.Assert
    
    Dim Result As IPropertyBinding
    Set Result = This.AbstractSUT.BindPropertyPath(This.BindingSource, This.SourcePropertyPath, This.BindingTarget, This.TargetPropertyPath)
    Dim outcome As cc_isr_Test_Fx.Assert
    
    Set outcome = This.Assert.AreEqual(1, This.ConcreteSUT.PropertyBindings.Count, "Property binding count should match")
    If outcome.AssertSuccessful Then
        Set outcome = This.Assert.IsTrue(Result Is This.ConcreteSUT.PropertyBindings.Item(1), "type of result should match.")
    End If
    
    If Not outcome.AssertSuccessful Then Debug.Print outcome.AssertMessage
    
    Set TestAddsToPropertyBindingsCollection = outcome

End Function

''' <summary>   [Unit Test] Test . </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestHandlePropertyChangedEvaluatesCommandCanExecute() As cc_isr_Test_Fx.Assert

    Dim outcome As cc_isr_Test_Fx.Assert

    This.HandlePropertyChangedSUT.HandlePropertyChanged This.BindingSource, This.SourceProperty
    
    Set outcome = This.CommandManagerStub.Verify(This.Assert, _
                                                "EvaluateCanExecute", 1)
    If Not outcome.AssertSuccessful Then Debug.Print outcome.AssertMessage
    
    Set TestHandlePropertyChangedEvaluatesCommandCanExecute = outcome

End Function

''' <summary>   [Unit Test] Test . </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestHandlePropertyChangedEvaluatesCommandCanExecuteForAnyPropertyChange() As cc_isr_Test_Fx.Assert
    
    Dim outcome As cc_isr_Test_Fx.Assert
    
    This.HandlePropertyChangedSUT.HandlePropertyChanged This.BindingSource, This.SourceProperty
    This.HandlePropertyChangedSUT.HandlePropertyChanged This.BindingSource, "Not" & This.SourceProperty
    
    Set outcome = This.CommandManagerStub.Verify(This.Assert, _
            "EvaluateCanExecute", 2)

    If Not outcome.AssertSuccessful Then Debug.Print outcome.AssertMessage
    
    Set TestHandlePropertyChangedEvaluatesCommandCanExecuteForAnyPropertyChange = outcome

End Function


