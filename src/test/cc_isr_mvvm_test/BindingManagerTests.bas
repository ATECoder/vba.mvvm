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
Public Sub BeforeEach()
    Set This.CommandManager = New TestCommandManager
    Set This.CommandManagerStub = This.CommandManager
    Set This.ConcreteSUT = cc_isr_MVVM.Factory.NewBindingManager().Initialize( _
                                        cc_isr_MVVM.Factory.NewAppContext().Initialize(a_debugOutput:=True), _
                                        cc_isr_MVVM.Factory.NewStringFormatterFactory)
    Set This.AbstractSUT = This.ConcreteSUT
    Set This.HandlePropertyChangedSUT = This.ConcreteSUT
    Set This.BindingSource = New TestBindingObject
    Set This.BindingTarget = New TestBindingObject
    Set This.Command = New TestCommand
    This.SourcePropertyPath = "TestStringProperty"
    This.TargetPropertyPath = "TestStringProperty"
    This.SourceProperty = "TestStringProperty"
    This.TargetProperty = "TestStringProperty"
    Set This.Assert = cc_isr_Test_Fx.Assert
End Sub

''' <summary>   Runs after each test. </summary>
Public Sub AfterEach()
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
    
    Debug.Print "TestCreateGuardsNonDefaultInstance " & _
        IIf(p_outcome.AssertSuccessful, "passed", "failed: " & p_outcome.AssertMessage)
    
    Set TestCreateGuardsNonDefaultInstance = p_outcome

End Function

''' <summary>   Binds the binding source to the target object. </summary>
''' <param name="a_progID">   The bound object program id. </param>
''' <param name="a_target">   The instance of the target ohect created from the <ParamRef name="a_progId"/> </param>
''' <returns>   <see cref="IPropertyBinding"/>. </returns>
Private Function DefaultPropertyPathBindingFor(ByVal a_progID As String, ByRef a_target As Object) As IPropertyBinding

    Set a_target = VBA.CreateObject(a_progID)
    Set DefaultPropertyPathBindingFor = This.AbstractSUT.BindPropertyPath(This.BindingSource, _
                                            This.SourcePropertyPath, a_target, This.TargetPropertyPath)

End Function

''' <summary>   Asserts creating a property path binding to the object defined by the specified program id. </summary>
''' <param name="a_progID">   The bound object program id. </param>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Private Function AssertCreatePropertyBindingType(ByVal a_progID As String, _
        ByVal a_type As Object) As cc_isr_Test_Fx.Assert
    
    Dim p_target As Object
    
    Dim p_result As IPropertyBinding
    Set p_result = DefaultPropertyPathBindingFor(a_progID, p_target)
        
    Set AssertCreatePropertyBindingType = This.Assert.AreEqual(TypeName(a_type), TypeName(p_result), _
                                 "Property type should equal expected type.")
    
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
        
    Set AssertCreatePropertyBindingType2 = This.Assert.AreEqual(TypeName(a_type), TypeName(p_result), _
                                 "Property type should equal expected type.")

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
       
    Set AssertCreatePropertyBindingName = This.Assert.AreEqual(a_name, p_result.Target.PropertyName, _
                                   "Property name should equal expected value.")
    
End Function

''' <summary>   [Unit Test] Test creating a property path binding for a check box control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestCheckBoxTargetCreatesPropertyBinding() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert

    Set p_outcome = AssertCreatePropertyBindingType(cc_isr_MVVM.BindingDefaults.CheckBoxProgId, _
                                                    cc_isr_MVVM.Factory.NewCheckBoxPropertyBinding)
    
    Debug.Print "TestCheckBoxTargetCreatesPropertyBinding " & _
        IIf(p_outcome.AssertSuccessful, "passed", "failed: " & p_outcome.AssertMessage)

    Set TestCheckBoxTargetCreatesPropertyBinding = p_outcome
    
End Function

''' <summary>   [Unit Test] Test creating a property path binding value for a check box control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestCheckBoxTargetBindsValueByDefault() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert

    ' dh: Legacy MVVM code specified 'Value' where tests return 'Caption' since the
    '     This.SourcePropertyPath and This.TargetPropertyPath are strings.
    Set p_outcome = AssertCreatePropertyBindingName(cc_isr_MVVM.BindingDefaults.CheckBoxProgId, "Caption")

    Debug.Print "TestCheckBoxTargetBindsValueByDefault " & _
        IIf(p_outcome.AssertSuccessful, "passed", "failed: " & p_outcome.AssertMessage)

    Set TestCheckBoxTargetBindsValueByDefault = p_outcome

End Function

''' <summary>   [Unit Test] Test creating a property path binding for a check box control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestComboBoxTargetCreatesPropertyBinding() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Set p_outcome = AssertCreatePropertyBindingType(cc_isr_MVVM.BindingDefaults.ComboBoxProgId, _
                                                    cc_isr_MVVM.Factory.NewComboBoxPropertyBinding)
                                                                    
    Debug.Print "TestComboBoxTargetCreatesPropertyBinding " & _
        IIf(p_outcome.AssertSuccessful, "passed", "failed: " & p_outcome.AssertMessage)
    
    Set TestComboBoxTargetCreatesPropertyBinding = p_outcome
                                                                    
End Function

''' <summary>   [Unit Test] Test creating a property path binding value for a combo box control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestComboBoxTargetBindsValueByDefault() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    ' dh: Legacy MVVM code specified 'Value' where tests return 'Text' since the
    '     This.SourcePropertyPath and This.TargetPropertyPath are strings.
    
    Set p_outcome = AssertCreatePropertyBindingName(cc_isr_MVVM.BindingDefaults.ComboBoxProgId, "Text")

    Debug.Print "TestComboBoxTargetBindsValueByDefault " & _
        IIf(p_outcome.AssertSuccessful, "passed", "failed: " & p_outcome.AssertMessage)
    
    Set TestComboBoxTargetBindsValueByDefault = p_outcome

End Function

''' <summary>   [Unit Test] Test creating a property path binding for a list box control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestListBoxTargetCreatesPropertyBinding() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Set p_outcome = AssertCreatePropertyBindingType(cc_isr_MVVM.BindingDefaults.ListBoxProgId, _
                                                    cc_isr_MVVM.Factory.NewListBoxPropertyBinding)

    Debug.Print "TestListBoxTargetCreatesPropertyBinding " & _
        IIf(p_outcome.AssertSuccessful, "passed", "failed: " & p_outcome.AssertMessage)
    
    Set TestListBoxTargetCreatesPropertyBinding = p_outcome

End Function

''' <summary>   [Unit Test] Test creating a property path binding value for a list box control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestListBoxTargetBindsValueByDefault() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    ' dh: Legacy MVVM code specified 'Value' where tests return 'Text' since the
    '     This.SourcePropertyPath and This.TargetPropertyPath are strings.
    
    Set p_outcome = AssertCreatePropertyBindingName(cc_isr_MVVM.BindingDefaults.ListBoxProgId, "Text")

    Debug.Print "TestListBoxTargetBindsValueByDefault " & _
        IIf(p_outcome.AssertSuccessful, "passed", "failed: " & p_outcome.AssertMessage)
    
    Set TestListBoxTargetBindsValueByDefault = p_outcome

End Function

''' <summary>   [Unit Test] Test creating a property path binding for a Multi Page control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestMultiPageTargetCreatesPropertyBinding() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Set p_outcome = AssertCreatePropertyBindingType(cc_isr_MVVM.BindingDefaults.MultiPageProgId, _
                                                    cc_isr_MVVM.Factory.NewMultiPagePropertyBinding)

    Debug.Print "TestMultiPageTargetCreatesPropertyBinding " & _
        IIf(p_outcome.AssertSuccessful, "passed", "failed: " & p_outcome.AssertMessage)
    
    Set TestMultiPageTargetCreatesPropertyBinding = p_outcome

End Function

''' <summary>   [Unit Test] Test creating a property path binding value for a Multi Page control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestMultiPageTargetBindsValueByDefault() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Set p_outcome = AssertCreatePropertyBindingName(cc_isr_MVVM.BindingDefaults.MultiPageProgId, "Value")

    Debug.Print "TestMultiPageTargetBindsValueByDefault " & _
        IIf(p_outcome.AssertSuccessful, "passed", "failed: " & p_outcome.AssertMessage)
    
    Set TestMultiPageTargetBindsValueByDefault = p_outcome

End Function

''' <summary>   [Unit Test] Test creating a property path binding for a Option Button control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestOptionButtonTargetCreatesPropertyBinding() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Set p_outcome = AssertCreatePropertyBindingType(cc_isr_MVVM.BindingDefaults.OptionButtonProgId, _
                                                    cc_isr_MVVM.Factory.NewOptionButtonPropertyBinding)

    Debug.Print "TestOptionButtonTargetCreatesPropertyBinding " & _
        IIf(p_outcome.AssertSuccessful, "passed", "failed: " & p_outcome.AssertMessage)
    
    Set TestOptionButtonTargetCreatesPropertyBinding = p_outcome

End Function

''' <summary>   [Unit Test] Test creating a property path binding value for a Option Button control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestOptionButtonTargetBindsValueByDefault() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    ' dh: Legacy MVVM code specified 'Value' where tests return 'Caption' since the
    '     This.SourcePropertyPath and This.TargetPropertyPath are strings.
    
    Set p_outcome = AssertCreatePropertyBindingName(cc_isr_MVVM.BindingDefaults.OptionButtonProgId, "Caption")

    Debug.Print "TestOptionButtonTargetBindsValueByDefault " & _
        IIf(p_outcome.AssertSuccessful, "passed", "failed: " & p_outcome.AssertMessage)
    
    Set TestOptionButtonTargetBindsValueByDefault = p_outcome

End Function

''' <summary>   [Unit Test] Test creating a property path binding for a Scroll Bar control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestScrollBarTargetCreatesPropertyBinding() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Set p_outcome = AssertCreatePropertyBindingType(cc_isr_MVVM.BindingDefaults.ScrollBarProgId, _
                                                    cc_isr_MVVM.Factory.NewScrollBarPropertyBinding)

    Debug.Print "TestScrollBarTargetCreatesPropertyBinding " & _
        IIf(p_outcome.AssertSuccessful, "passed", "failed: " & p_outcome.AssertMessage)
    
    Set TestScrollBarTargetCreatesPropertyBinding = p_outcome

End Function

''' <summary>   [Unit Test] Test creating a property path binding value for a Scroll Bar control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestScrollBarTargetBindsValueByDefault() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Set p_outcome = AssertCreatePropertyBindingName(cc_isr_MVVM.BindingDefaults.ScrollBarProgId, "Value")

    Debug.Print "TestScrollBarTargetBindsValueByDefault " & _
        IIf(p_outcome.AssertSuccessful, "passed", "failed: " & p_outcome.AssertMessage)
    
    Set TestScrollBarTargetBindsValueByDefault = p_outcome

End Function

''' <summary>   [Unit Test] Test creating a property path binding for a Spin Button control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestSpinButtonTargetCreatesPropertyBinding() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Set p_outcome = AssertCreatePropertyBindingType(cc_isr_MVVM.BindingDefaults.SpinButtonProgId, _
                                                    cc_isr_MVVM.Factory.NewSpinButtonPropertyBinding)

    Debug.Print "TestSpinButtonTargetCreatesPropertyBinding " & _
        IIf(p_outcome.AssertSuccessful, "passed", "failed: " & p_outcome.AssertMessage)
    
    Set TestSpinButtonTargetCreatesPropertyBinding = p_outcome

End Function

''' <summary>   [Unit Test] Test creating a property path binding value for a Spin Button control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestSpinButtonTargetBindsValueByDefault() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Set p_outcome = AssertCreatePropertyBindingName(cc_isr_MVVM.BindingDefaults.SpinButtonProgId, "Value")

    Debug.Print "TestSpinButtonTargetBindsValueByDefault " & _
        IIf(p_outcome.AssertSuccessful, "passed", "failed: " & p_outcome.AssertMessage)
    
    Set TestSpinButtonTargetBindsValueByDefault = p_outcome

End Function

''' <summary>   [Unit Test] Test creating a property path binding for a Tab Strip control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestTabStripTargetCreatesPropertyBinding() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Set p_outcome = AssertCreatePropertyBindingType(cc_isr_MVVM.BindingDefaults.TabStripProgId, _
                                                    cc_isr_MVVM.Factory.NewTabStripPropertyBinding)
                                                    
    Debug.Print "TestTabStripTargetCreatesPropertyBinding " & _
        IIf(p_outcome.AssertSuccessful, "passed", "failed: " & p_outcome.AssertMessage)
    
    Set TestTabStripTargetCreatesPropertyBinding = p_outcome
                                                                    
End Function

''' <summary>   [Unit Test] Test creating a property path binding value for a Tab Strip control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestTabStripTargetBindsValueByDefault() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Set p_outcome = AssertCreatePropertyBindingName(cc_isr_MVVM.BindingDefaults.TabStripProgId, "Value")

    Debug.Print "TestTabStripTargetBindsValueByDefault " & _
        IIf(p_outcome.AssertSuccessful, "passed", "failed: " & p_outcome.AssertMessage)
    
    Set TestTabStripTargetBindsValueByDefault = p_outcome

End Function

''' <summary>   [Unit Test] Test creating a property path binding for a Text Box control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestTextBoxTargetCreatesPropertyBinding() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Set p_outcome = AssertCreatePropertyBindingType(cc_isr_MVVM.BindingDefaults.TextBoxProgId, _
                                                    cc_isr_MVVM.Factory.NewTextBoxPropertyBinding)

    Debug.Print "TestTextBoxTargetCreatesPropertyBinding " & _
        IIf(p_outcome.AssertSuccessful, "passed", "failed: " & p_outcome.AssertMessage)
    
    Set TestTextBoxTargetCreatesPropertyBinding = p_outcome
                                                                    
End Function

''' <summary>   [Unit Test] Test creating a property path binding value for a Text Box control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestTextBoxTargetBindsValueByDefault() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Set p_outcome = AssertCreatePropertyBindingName(cc_isr_MVVM.BindingDefaults.TextBoxProgId, "Text")
    

    Debug.Print "TestCheckBoxTargetBindsValueByDefault " & _
        IIf(p_outcome.AssertSuccessful, "passed", "failed: " & p_outcome.AssertMessage)
    
    Set TestTextBoxTargetBindsValueByDefault = p_outcome

End Function

''' <summary>   [Unit Test] Test creating a property path binding for a Frame control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestFrameTargetCreatesOneWayBindingWithNonDefaultTarget() As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    This.TargetPropertyPath = "Font.Bold"
   
    Dim p_target As Object
    
    Dim p_result As IPropertyBinding
    Set p_result = DefaultPropertyPathBindingFor(cc_isr_MVVM.BindingDefaults.FrameProgId, p_target)
    
    Set p_outcome = This.Assert.AreEqual(TypeName(cc_isr_MVVM.Factory.NewOneWayPropertyBinding), _
                                         TypeName(p_result), _
                                   "Property type name should equal expected type name.")

    Debug.Print "TestFrameTargetCreatesOneWayBindingWithNonDefaultTarget " & _
        IIf(p_outcome.AssertSuccessful, "passed", "failed: " & p_outcome.AssertMessage)
    
    Set TestFrameTargetCreatesOneWayBindingWithNonDefaultTarget = p_outcome
    
    
End Function

''' <summary>   [Unit Test] Test creating a property path binding for a label control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestLabelTargetCreatesOneWayBindingWithNonDefaultTarget() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Set p_outcome = AssertCreatePropertyBindingType2(cc_isr_MVVM.BindingDefaults.LabelProgId, _
                                                    cc_isr_MVVM.Factory.NewOneWayPropertyBinding, _
                                                    "Font.Bold")

    Debug.Print "TestLabelTargetCreatesOneWayBindingWithNonDefaultTarget " & _
        IIf(p_outcome.AssertSuccessful, "passed", "failed: " & p_outcome.AssertMessage)
    
    Set TestLabelTargetCreatesOneWayBindingWithNonDefaultTarget = p_outcome

End Function

''' <summary>   [Unit Test] Test creating a property path binding for a frame control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestFrameTargetBindsCaptionPropertyByDefault() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Set p_outcome = AssertCreatePropertyBindingType(cc_isr_MVVM.BindingDefaults.FrameProgId, _
                                                    cc_isr_MVVM.Factory.NewOneWayPropertyBinding)

    Debug.Print "TestFrameTargetBindsCaptionPropertyByDefault " & _
        IIf(p_outcome.AssertSuccessful, "passed", "failed: " & p_outcome.AssertMessage)
    
    Set TestFrameTargetBindsCaptionPropertyByDefault = p_outcome

End Function

''' <summary>   [Unit Test] Test creating a property path binding for a label control. </summary>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function TestLabelTargetBindsCaptionPropertyByDefault() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Set p_outcome = AssertCreatePropertyBindingType(cc_isr_MVVM.BindingDefaults.LabelProgId, _
                                                    cc_isr_MVVM.Factory.NewOneWayPropertyBinding)

    Debug.Print "TestLabelTargetBindsCaptionPropertyByDefault " & _
        IIf(p_outcome.AssertSuccessful, "passed", "failed: " & p_outcome.AssertMessage)
    
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

    Debug.Print "TestNonControlTargetCreatesOneWayBinding " & _
        IIf(p_outcome.AssertSuccessful, "passed", "failed: " & p_outcome.AssertMessage)
    
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

    Debug.Print "TestNonControlTargetRequiresTargetPropertyPath " & _
        IIf(p_outcome.AssertSuccessful, "passed", "failed: " & p_outcome.AssertMessage)
    
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

    Debug.Print "TestAddsToPropertyBindingsCollection " & _
        IIf(p_outcome.AssertSuccessful, "passed", "failed: " & p_outcome.AssertMessage)
    
    Set TestAddsToPropertyBindingsCollection = p_outcome

End Function

''' <summary>   [Unit Test] Test . </summary>
''' <remarks>   This methods fails because the binding source does not have any bound commands. </remarks>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function PendingTestHandlePropertyChangedEvaluatesCommandCanExecute() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert

    This.HandlePropertyChangedSUT.HandlePropertyChanged This.BindingSource, This.SourceProperty
    
    Set p_outcome = This.CommandManagerStub.Verify(This.Assert, "EvaluateCanExecute", 1)

    Debug.Print "TestHandlePropertyChangedEvaluatesCommandCanExecute " & _
        IIf(p_outcome.AssertSuccessful, "passed", "failed: " & p_outcome.AssertMessage)
    
    Set TestHandlePropertyChangedEvaluatesCommandCanExecute = p_outcome

End Function

''' <summary>   [Unit Test] Test . </summary>
''' <remarks>   This methods fails because the binding source does not have any bound commands. </remarks>
''' <returns>   <see cref="cc_isr_Test_Fx.Assert"/>. </returns>
Public Function PendingTestHandlePropertyChangedEvaluatesCommandCanExecuteForAnyPropertyChange() As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    This.HandlePropertyChangedSUT.HandlePropertyChanged This.BindingSource, This.SourceProperty
    This.HandlePropertyChangedSUT.HandlePropertyChanged This.BindingSource, "Not" & This.SourceProperty
    
    Set p_outcome = This.CommandManagerStub.Verify(This.Assert, _
            "EvaluateCanExecute", 2)


    Debug.Print "TestHandlePropertyChangedEvaluatesCommandCanExecuteForAnyPropertyChange " & _
        IIf(p_outcome.AssertSuccessful, "passed", "failed: " & p_outcome.AssertMessage)
    
    Set TestHandlePropertyChangedEvaluatesCommandCanExecuteForAnyPropertyChange = p_outcome

End Function


