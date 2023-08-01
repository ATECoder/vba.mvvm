Attribute VB_Name = "BindingManagerTests"
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Unit tests: Binding manager. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit
' Option Private Module

Private Type ThisData

    ExpectedErrNumber As Long
    ExpectedErrSource As String
    ExpectedErrorCaught As Boolean
    
    ConcreteSUT As cc_isr_MVVM.BindingManager
    AbstractSUT As cc_isr_MVVM.IBindingManager
    HandlePropertyChangedSUT As cc_isr_MVVM.IHandlePropertyChanged
    
    CommandManager As cc_isr_MVVM.ICommandManager
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
    Set This.ConcreteSUT = cc_isr_MVVM.Factory.CreateBindingManager(This.CommandManager, cc_isr_MVVM.Factory.NewStringFormatterFactory)
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

    Dim p_outcome As cc_isr_Test_Fx.Assert
    This.ExpectedErrNumber = cc_isr_Core.UserDefinedErrors.InvalidOperationError.Code
    Dim p_bindingManger As cc_isr_MVVM.BindingManager
    p_bindingManger = cc_isr_MVVM.Factory.NewBindingManager
    
    On Error Resume Next
    '@Ignore FunctionReturnValueDiscarded, FunctionReturnValueNotUsed
    p_bindingManger.Initialize This.CommandManager, cc_isr_MVVM.Factory.NewStringFormatterFactory
    Set p_outcome = AssertExpectError
    On Error GoTo 0
    
    If Not p_outcome.AssertSuccessful Then Debug.Print p_outcome.AssertMessage
    
    Set TestCreateGuardsNonDefaultInstance = p_outcome

End Function

''' <summary>   Binds the binding source to the target object. </summary>
''' <param name="a_progID">   The bound object program id. </param>
''' <param name="a_target">   The instance of the target ohect created from the <ParamRef name="a_progId"/> </param>
''' <returns>   <see cref="IPropertyBinding"/>. </returns>
Private Function DefaultPropertyPathBindingFor(ByVal a_progID As String, ByRef a_target As Object) As IPropertyBinding

    Set a_target = VBA.CreateObject(a_progID)
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    Set p_outcome = This.AbstractSUT.BindPropertyPath(This.BindingSource, _
                    This.SourcePropertyPath, a_target, This.TargetPropertyPath)
    
    If Not p_outcome.AssertSuccessful Then Debug.Print p_outcome.AssertMessage
    
    Set DefaultPropertyPathBindingFor = p_outcome

End Function

''' <summary>   Asserts creating a property path binding to the object defined by the specified program id. </summary>
''' <param name="a_progID">   The bound object program id. </param>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Private Function AssertCreatePropertyBindingType(ByVal a_progID As String, _
        ByVal a_type As Object) As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
   
    Dim p_target As Object
    
    Dim p_result As IPropertyBinding
    Set p_result = DefaultPropertyPathBindingFor(a_progID, p_target)
        
    Set p_outcome = This.Assert.IsTrue(TypeOf p_result Is CheckBoxPropertyBinding, _
                                 "Property type should equal expected type.")
    
    If Not p_outcome.AssertSuccessful Then Debug.Print p_outcome.AssertMessage
    
    Set AssertCreatePropertyBindingType = p_outcome

End Function

''' <summary>   Asserts creating a property path binding to the object defined by the specified program id. </summary>
''' <param name="a_progID">   The bound object program id. </param>
''' <param name="a_targetProeprtyPath">   [Optional, String, Null] The target property path. </param>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Private Function AssertCreatePropertyBindingType2(ByVal a_progID As String, _
        ByVal a_type As Object, _
        Optional ByVal a_targetProeprtyPath As String = VBA.vbNullString) As cc_isr_Test_Fx.Assert
    
    This.TargetPropertyPath = a_targetProeprtyPath
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
   
    Dim p_target As Object
    
    Dim p_result As IPropertyBinding
    Set p_result = DefaultPropertyPathBindingFor(a_progID, p_target)
        
    Set p_outcome = This.Assert.IsTrue(TypeOf p_result Is CheckBoxPropertyBinding, _
                                 "Property type should equal expected type.")
    
    If Not p_outcome.AssertSuccessful Then Debug.Print p_outcome.AssertMessage
    
    Set AssertCreatePropertyBindingType2 = p_outcome

End Function


''' <summary>   Asserts creating a property path binding to the object defined by the specified program id. </summary>
''' <param name="a_progID">               The bound object program id. </param>
''' <param name="a_name">                 The expected property name. </param>
''' <param name="a_targetProeprtyPath">   [Optional, String, Null] The target property path. </param>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Private Function AssertCreatePropertyBindingName(ByVal a_progID As String, _
        ByVal a_name As String, _
        Optional ByVal a_targetProeprtyPath As String = VBA.vbNullString) As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
   
    This.TargetPropertyPath = a_targetProeprtyPath
   
    Dim p_target As Object
    
    Dim p_result As IPropertyBinding
    Set p_result = DefaultPropertyPathBindingFor(a_progID, p_target)
        
       
    Set p_outcome = This.Assert.AreEqual(a_name, p_result.Target.PropertyName, _
                                   "Property name should equal expected value.")
    
    If Not p_outcome.AssertSuccessful Then Debug.Print p_outcome.AssertMessage
    
    Set AssertCreatePropertyBindingName = p_outcome

End Function

''' <summary>   [Unit Test] Test creating a property path binding for a check box control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestCheckBoxTargetCreatesPropertyBinding() As cc_isr_Test_Fx.Assert

    Dim p_result As CheckBoxPropertyBinding
    
    Dim p_outcome As cc_isr_Test_Fx.Assert

    Set p_outcome = AssertCreatePropertyBindingType(cc_isr_MVVM.BindingDefaults.CheckBoxProgId, p_result)
    
    If Not p_outcome.AssertSuccessful Then Debug.Print p_outcome.AssertMessage

    Set TestCheckBoxTargetCreatesPropertyBinding = p_outcome
    
End Function

''' <summary>   [Unit Test] Test creating a property path binding value for a check box control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestCheckBoxTargetBindsValueByDefault() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert

    Set p_outcome = AssertCreatePropertyBindingName(cc_isr_MVVM.BindingDefaults.CheckBoxProgId, "Value")

    If Not p_outcome.AssertSuccessful Then Debug.Print p_outcome.AssertMessage

    Set TestCheckBoxTargetBindsValueByDefault = p_outcome

End Function

''' <summary>   [Unit Test] Test creating a property path binding for a check box control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestComboBoxTargetCreatesPropertyBinding() As cc_isr_Test_Fx.Assert

    Dim p_result As ComboBoxPropertyBinding
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Set p_outcome = AssertCreatePropertyBindingType(cc_isr_MVVM.BindingDefaults.ComboBoxProgId, p_result)
                                                                    
    If Not p_outcome.AssertSuccessful Then Debug.Print p_outcome.AssertMessage
    
    Set TestComboBoxTargetCreatesPropertyBinding = p_outcome
                                                                    
End Function

''' <summary>   [Unit Test] Test creating a property path binding value for a combo box control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestComboBoxTargetBindsValueByDefault() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Set p_outcome = AssertCreatePropertyBindingName(cc_isr_MVVM.BindingDefaults.ComboBoxProgId, "Value")

    If Not p_outcome.AssertSuccessful Then Debug.Print p_outcome.AssertMessage
    
    Set TestComboBoxTargetBindsValueByDefault = p_outcome

End Function

''' <summary>   [Unit Test] Test creating a property path binding for a list box control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestListBoxTargetCreatesPropertyBinding() As cc_isr_Test_Fx.Assert

    Dim p_result As ListBoxPropertyBinding
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Set p_outcome = AssertCreatePropertyBindingType(cc_isr_MVVM.BindingDefaults.ListBoxProgId, p_result)
                                                                    
    If Not p_outcome.AssertSuccessful Then Debug.Print p_outcome.AssertMessage
    
    Set TestListBoxTargetCreatesPropertyBinding = p_outcome

End Function

''' <summary>   [Unit Test] Test creating a property path binding value for a list box control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestListBoxTargetBindsValueByDefault() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Set p_outcome = AssertCreatePropertyBindingName(cc_isr_MVVM.BindingDefaults.ListBoxProgId, "Value")

    If Not p_outcome.AssertSuccessful Then Debug.Print p_outcome.AssertMessage
    
    Set TestListBoxTargetBindsValueByDefault = p_outcome

End Function

''' <summary>   [Unit Test] Test creating a property path binding for a Multi Page control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestMultiPageTargetCreatesPropertyBinding() As cc_isr_Test_Fx.Assert

    Dim p_result As MultiPagePropertyBinding
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Set p_outcome = AssertCreatePropertyBindingType(cc_isr_MVVM.BindingDefaults.MultiPageProgId, p_result)
                                                                    
    If Not p_outcome.AssertSuccessful Then Debug.Print p_outcome.AssertMessage
    
    Set TestMultiPageTargetCreatesPropertyBinding = p_outcome

End Function

''' <summary>   [Unit Test] Test creating a property path binding value for a Multi Page control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestMultiPageTargetBindsValueByDefault() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Set p_outcome = AssertCreatePropertyBindingName(cc_isr_MVVM.BindingDefaults.MultiPageProgId, "Value")

    If Not p_outcome.AssertSuccessful Then Debug.Print p_outcome.AssertMessage
    
    Set TestMultiPageTargetBindsValueByDefault = p_outcome

End Function

''' <summary>   [Unit Test] Test creating a property path binding for a Option Button control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestOptionButtonTargetCreatesPropertyBinding() As cc_isr_Test_Fx.Assert

    Dim p_result As OptionButtonPropertyBinding
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Set p_outcome = AssertCreatePropertyBindingType(cc_isr_MVVM.BindingDefaults.OptionButtonProgId, p_result)
                                                                    
    If Not p_outcome.AssertSuccessful Then Debug.Print p_outcome.AssertMessage
    
    Set TestOptionButtonTargetCreatesPropertyBinding = p_outcome

End Function

''' <summary>   [Unit Test] Test creating a property path binding value for a Option Button control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestOptionButtonTargetBindsValueByDefault() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Set p_outcome = AssertCreatePropertyBindingName(cc_isr_MVVM.BindingDefaults.OptionButtonProgId, "Value")

    If Not p_outcome.AssertSuccessful Then Debug.Print p_outcome.AssertMessage
    
    Set TestOptionButtonTargetBindsValueByDefault = p_outcome

End Function

''' <summary>   [Unit Test] Test creating a property path binding for a Scroll Bar control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestScrollBarTargetCreatesPropertyBinding() As cc_isr_Test_Fx.Assert

    Dim p_result As ScrollBarPropertyBinding
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Set p_outcome = AssertCreatePropertyBindingType(cc_isr_MVVM.BindingDefaults.ScrollBarProgId, p_result)
                                                                    
    If Not p_outcome.AssertSuccessful Then Debug.Print p_outcome.AssertMessage
    
    Set TestScrollBarTargetCreatesPropertyBinding = p_outcome

End Function

''' <summary>   [Unit Test] Test creating a property path binding value for a Scroll Bar control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestScrollBarTargetBindsValueByDefault() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Set p_outcome = AssertCreatePropertyBindingName(cc_isr_MVVM.BindingDefaults.ScrollBarProgId, "Value")

    If Not p_outcome.AssertSuccessful Then Debug.Print p_outcome.AssertMessage
    
    Set TestScrollBarTargetBindsValueByDefault = p_outcome

End Function

''' <summary>   [Unit Test] Test creating a property path binding for a Spin Button control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestSpinButtonTargetCreatesPropertyBinding() As cc_isr_Test_Fx.Assert

    Dim p_result As SpinButtonPropertyBinding
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Set p_outcome = AssertCreatePropertyBindingType(cc_isr_MVVM.BindingDefaults.SpinButtonProgId, p_result)
                                                                    
    If Not p_outcome.AssertSuccessful Then Debug.Print p_outcome.AssertMessage
    
    Set TestSpinButtonTargetCreatesPropertyBinding = p_outcome

End Function

''' <summary>   [Unit Test] Test creating a property path binding value for a Spin Button control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestSpinButtonTargetBindsValueByDefault() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Set p_outcome = AssertCreatePropertyBindingName(cc_isr_MVVM.BindingDefaults.SpinButtonProgId, "Value")

    If Not p_outcome.AssertSuccessful Then Debug.Print p_outcome.AssertMessage
    
    Set TestSpinButtonTargetBindsValueByDefault = p_outcome

End Function

''' <summary>   [Unit Test] Test creating a property path binding for a Tab Strip control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestTabStripTargetCreatesPropertyBinding() As cc_isr_Test_Fx.Assert

    Dim p_result As TabStripPropertyBinding
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Set p_outcome = AssertCreatePropertyBindingType(cc_isr_MVVM.BindingDefaults.TabStripProgId, p_result)
                                                                    
    If Not p_outcome.AssertSuccessful Then Debug.Print p_outcome.AssertMessage
    
    Set TestTabStripTargetCreatesPropertyBinding = p_outcome
                                                                    
End Function

''' <summary>   [Unit Test] Test creating a property path binding value for a Tab Strip control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestTabStripTargetBindsValueByDefault() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Set p_outcome = AssertCreatePropertyBindingName(cc_isr_MVVM.BindingDefaults.TabStripProgId, "Value")

    If Not p_outcome.AssertSuccessful Then Debug.Print p_outcome.AssertMessage
    
    Set TestTabStripTargetBindsValueByDefault = p_outcome

End Function

''' <summary>   [Unit Test] Test creating a property path binding for a Text Box control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestTextBoxTargetCreatesPropertyBinding() As cc_isr_Test_Fx.Assert

    Dim p_result As TextBoxPropertyBinding
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Set p_outcome = AssertCreatePropertyBindingType(cc_isr_MVVM.BindingDefaults.TextBoxProgId, p_result)
    
    If Not p_outcome.AssertSuccessful Then Debug.Print p_outcome.AssertMessage
    
    Set TestTextBoxTargetCreatesPropertyBinding = p_outcome
                                                                    
End Function

''' <summary>   [Unit Test] Test creating a property path binding value for a Text Box control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestTextBoxTargetBindsValueByDefault() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Set p_outcome = AssertCreatePropertyBindingName(cc_isr_MVVM.BindingDefaults.TextBoxProgId, "Text")
    
    If Not p_outcome.AssertSuccessful Then Debug.Print p_outcome.AssertMessage
    
    Set TestCheckBoxTargetBindsValueByDefault = p_outcome

End Function

''' <summary>   [Unit Test] Test creating a property path binding for a Frame control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestFrameTargetCreatesOneWayBindingWithNonDefaultTarget() As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    This.TargetPropertyPath = "Font.Bold"
   
    Dim p_target As Object
    
    Dim p_result As IPropertyBinding
    Set p_result = DefaultPropertyPathBindingFor(cc_isr_MVVM.BindingDefaults.FrameProgId, p_target)
    
    Dim p_expecterdType As cc_isr_MVVM.OneWayPropertyBinding
    Set p_outcome = This.Assert.AreEqual(TypeName(p_expecterdType), TypeName(p_result), _
                                   "Property type name should equal expected type name.")
    
    If Not p_outcome.AssertSuccessful Then Debug.Print p_outcome.AssertMessage
    
    Set TestFrameTargetCreatesOneWayBindingWithNonDefaultTarget = p_outcome
    
    
End Function

''' <summary>   [Unit Test] Test creating a property path binding for a label control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestLabelTargetCreatesOneWayBindingWithNonDefaultTarget() As cc_isr_Test_Fx.Assert

    Dim p_result As OneWayPropertyBinding
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Set p_outcome = AssertCreatePropertyBindingType2(cc_isr_MVVM.BindingDefaults.LabelProgId, _
                                                    p_result, "Font.Bold")
                                                                    
    If Not p_outcome.AssertSuccessful Then Debug.Print p_outcome.AssertMessage
    
    Set TestLabelTargetCreatesOneWayBindingWithNonDefaultTarget = p_outcome

End Function

''' <summary>   [Unit Test] Test creating a property path binding for a frame control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestFrameTargetBindsCaptionPropertyByDefault() As cc_isr_Test_Fx.Assert

    Dim p_result As CaptionPropertyBinding
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Set p_outcome = AssertCreatePropertyBindingType(cc_isr_MVVM.BindingDefaults.FrameProgId, p_result)
                                                                    
    If Not p_outcome.AssertSuccessful Then Debug.Print p_outcome.AssertMessage
    
    Set TestFrameTargetBindsCaptionPropertyByDefault = p_outcome

End Function

''' <summary>   [Unit Test] Test creating a property path binding for a label control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestLabelTargetBindsCaptionPropertyByDefault() As cc_isr_Test_Fx.Assert

    Dim p_result As CaptionPropertyBinding
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Set p_outcome = AssertCreatePropertyBindingType(cc_isr_MVVM.BindingDefaults.LabelProgId, p_result)
                                                                    
    If Not p_outcome.AssertSuccessful Then Debug.Print p_outcome.AssertMessage
    
    Set TestLabelTargetBindsCaptionPropertyByDefault = p_outcome

End Function


''' <summary>   [Unit Test] Test non control target creates one way binding. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestNonControlTargetCreatesOneWayBinding() As cc_isr_Test_Fx.Assert
    
    Dim p_result As IPropertyBinding
    
    Set p_result = This.AbstractSUT.BindPropertyPath(This.BindingSource, This.SourcePropertyPath, _
                                                   This.BindingTarget, This.TargetPropertyPath)
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Set p_outcome = This.Assert.IsTrue(TypeOf p_result Is OneWayPropertyBinding, _
                                        "Type of result should equal expected type.")
    
    If Not p_outcome.AssertSuccessful Then Debug.Print p_outcome.AssertMessage
    
    Set TestNonControlTargetCreatesOneWayBinding = p_outcome

End Function

''' <summary>   [Unit Test] Test non control target requires target property path. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestNonControlTargetRequiresTargetPropertyPath() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert
    This.ExpectedErrNumber = cc_isr_Core.UserDefinedErrors.InvalidArgumentError.Code
    On Error Resume Next
    This.AbstractSUT.BindPropertyPath _
        This.BindingSource, _
        This.SourcePropertyPath, _
        This.BindingTarget, _
        a_targetProperty:=vbNullString
    Set p_outcome = AssertExpectError
    On Error GoTo 0
    
    If Not p_outcome.AssertSuccessful Then Debug.Print p_outcome.AssertMessage
    
    Set TestNonControlTargetRequiresTargetPropertyPath = p_outcome

End Function

''' <summary>   [Unit Test] Test . </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestAddsToPropertyBindingsCollection() As cc_isr_Test_Fx.Assert
    
    Dim Result As IPropertyBinding
    Set Result = This.AbstractSUT.BindPropertyPath(This.BindingSource, This.SourcePropertyPath, This.BindingTarget, This.TargetPropertyPath)
    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Set p_outcome = This.Assert.AreEqual(1, This.ConcreteSUT.PropertyBindings.Count, "Property binding count should match")
    If p_outcome.AssertSuccessful Then
        Set p_outcome = This.Assert.IsTrue(Result Is This.ConcreteSUT.PropertyBindings.Item(1), "type of result should match.")
    End If
    
    If Not p_outcome.AssertSuccessful Then Debug.Print p_outcome.AssertMessage
    
    Set TestAddsToPropertyBindingsCollection = p_outcome

End Function

''' <summary>   [Unit Test] Test . </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestHandlePropertyChangedEvaluatesCommandCanExecute() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert

    This.HandlePropertyChangedSUT.HandlePropertyChanged This.BindingSource, This.SourceProperty
    
    Set p_outcome = This.CommandManagerStub.Verify(This.Assert, _
                                                "EvaluateCanExecute", 1)
    If Not p_outcome.AssertSuccessful Then Debug.Print p_outcome.AssertMessage
    
    Set TestHandlePropertyChangedEvaluatesCommandCanExecute = p_outcome

End Function

''' <summary>   [Unit Test] Test . </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestHandlePropertyChangedEvaluatesCommandCanExecuteForAnyPropertyChange() As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    This.HandlePropertyChangedSUT.HandlePropertyChanged This.BindingSource, This.SourceProperty
    This.HandlePropertyChangedSUT.HandlePropertyChanged This.BindingSource, "Not" & This.SourceProperty
    
    Set p_outcome = This.CommandManagerStub.Verify(This.Assert, _
            "EvaluateCanExecute", 2)

    If Not p_outcome.AssertSuccessful Then Debug.Print p_outcome.AssertMessage
    
    Set TestHandlePropertyChangedEvaluatesCommandCanExecuteForAnyPropertyChange = p_outcome

End Function


