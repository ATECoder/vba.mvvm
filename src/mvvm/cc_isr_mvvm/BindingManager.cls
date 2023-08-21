VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BindingManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An object responsible for managing property bindings."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   An object responsible for managing property bindings. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit
Implements IBindingManager
Implements IHandlePropertyChanged
Implements cc_isr_Core.IDisposable

Private Type TState
    Context As IAppContext
    
    PropertyBindings As Collection
    NotifierFactory As INotifierFactory
    StringFormatterFactory As IStringFormatterFactory
    DebugOutput As Boolean
    BindingTargetStrategies As Dictionary
End Type

Private This As TState

''' <summary>   Initializes and returns an <see cref="IBindingManager"/> representation of
'''             this <see cref="BindingManager"/> object. </summary>
''' <returns>   [<see cref="IBindingManager"/>]. </returns>
Public Function Initialize(ByVal a_context As IAppContext, _
        ByVal a_stringFormatterFactory As IStringFormatterFactory, _
        Optional ByVal a_factory As INotifierFactory, _
        Optional ByVal a_debugOutput As Boolean = False) As IBindingManager
    
    Set Me.Context = a_context
    Me.DebugOutput = a_debugOutput
    
    If a_stringFormatterFactory Is Nothing Then
        Set Me.StringFormatterFactory = New StringFormatterFactory
    Else
        Set Me.StringFormatterFactory = a_stringFormatterFactory
    End If
    
    If a_factory Is Nothing Then
        Set Me.NotifierFactory = New NotifierBaseFactory
    Else
        Set Me.NotifierFactory = a_factory
    End If
    
    Set Initialize = Me
    
End Function

Public Property Get NotifierFactory() As INotifierFactory
    Set NotifierFactory = This.NotifierFactory
End Property

Public Property Set NotifierFactory(ByVal a_value As INotifierFactory)
    
    Dim p_source As String
    p_source = ThisWorkbook.VBProject.Name & "." & VBA.Information.TypeName(Me) & ".NotifierFactory"
    cc_isr_Core_IO.GuardClauses.GuardDoubleInitialization This.NotifierFactory, p_source
    Set This.NotifierFactory = a_value

End Property

Private Sub Apply(ByVal a_source As Object)
    
    Dim p_binding As IPropertyBinding
    For Each p_binding In This.PropertyBindings
        If a_source Is p_binding.Source.Context Then p_binding.Apply
    Next

End Sub

Public Property Get Context() As IAppContext
    Set Context = This.Context
End Property

Public Property Set Context(ByVal a_value As IAppContext)
    
    Dim p_source As String
    p_source = ThisWorkbook.VBProject.Name & "." & VBA.Information.TypeName(Me) & ".Context"
    cc_isr_Core_IO.GuardClauses.GuardNullReference a_value, p_source
    cc_isr_Core_IO.GuardClauses.GuardDoubleInitialization This.Context, p_source
    
    Set This.Context = a_value
End Property

Public Property Get Handlers() As Collection
    Set Handlers = This.PropertyBindings
End Property

Public Property Get PropertyBindings() As Collection
    Set PropertyBindings = This.PropertyBindings
End Property

Public Property Get StringFormatterFactory() As IStringFormatterFactory
    Set StringFormatterFactory = This.StringFormatterFactory
End Property

Public Property Set StringFormatterFactory(ByVal a_value As IStringFormatterFactory)
    
    Dim p_source As String
    p_source = ThisWorkbook.VBProject.Name & "." & VBA.Information.TypeName(Me) & ".StringFormatterFactory"
    cc_isr_Core_IO.GuardClauses.GuardNullReference a_value, p_source
    cc_isr_Core_IO.GuardClauses.GuardDoubleInitialization This.StringFormatterFactory, p_source
    Set This.StringFormatterFactory = a_value

End Property

Public Property Get DebugOutput() As Boolean
    DebugOutput = This.DebugOutput
End Property

Public Property Let DebugOutput(ByVal a_value As Boolean)
    This.DebugOutput = a_value
End Property

''' <summary>   Releases all held property bindings, prepares the object for proper destruction. </summary>
Public Sub Terminate()
Attribute Terminate.VB_Description = "Releases all held property bindings, prepares the object for proper destruction."
    '@Ignore VariableNotUsed
    Dim i As Long
    For i = 1 To This.PropertyBindings.Count
        This.PropertyBindings.Remove 1
    Next
    Set This.PropertyBindings = Nothing
End Sub

''' <summary>   Class initializes. </summary>
''' <remarks>   Presently, type of MS Control is not an MS Control.  Potentially, It might be better to
'''             enumerate the binding target strategies by the control name as
'''             prefixing the control name with the library name is not robust.
'''             Note, though, the type of range indeed includes the Excel library name. </remarks>
Private Sub Class_Initialize()
    Set This.BindingTargetStrategies = New Dictionary
    Set This.PropertyBindings = New Collection
    
    With This.BindingTargetStrategies
        .Add "MSForms.CheckBox", New CheckBoxBindingStrategy
        .Add "MSForms.CommandButton", New CommandButtonBindingStrategy
        .Add "MSForms.ComboBox", New ComboBoxBindingStrategy
        .Add "MSForms.Frame", New CaptionBindingStrategy
        .Add "MSForms.Label", New CaptionBindingStrategy
        .Add "MSForms.ListBox", New ListBoxBindingStrategy
        .Add "MSForms.MultiPage", New MultiPageBindingStrategy
        .Add "MSForms.OptionButton", New OptionButtonBindingStrategy
        .Add "MSForms.ScrollBar", New ScrollBarBindingStrategy
        .Add "MSForms.SpinButton", New SpinButtonBindingStrategy
        .Add "MSForms.TabStrip", New TabStripBindingStrategy
        .Add "MSForms.TextBox", New TextBoxBindingStrategy
        .Add "Excel.Range", New WorksheetCellBindingStrategy
    End With
End Sub

Private Sub Class_Terminate()
    Set This.BindingTargetStrategies = Nothing
    Set This.PropertyBindings = Nothing
End Sub

Private Function IBindingManager_BindPropertyPath(ByVal a_source As Object, ByVal a_propertyPath As String, _
        ByVal a_target As Object, _
        Optional ByVal a_targetProperty As String, _
        Optional ByVal a_mode As BindingMode = BindingMode.TwoWayBinding, _
        Optional ByVal a_updateTrigger As BindingUpdateSourceTrigger = BindingUpdateSourceTrigger.OnPropertyChanged, _
        Optional ByVal a_validator As IValueValidator, _
        Optional ByVal a_converter As IValueConverter, _
        Optional ByVal a_formatString As String, _
        Optional ByVal a_validationAdorner As IDynamicAdorner) As IPropertyBinding
    
    Dim p_formatter As IStringFormatter
    If a_formatString <> VBA.vbNullString And Not This.StringFormatterFactory Is Nothing Then
        Set p_formatter = This.StringFormatterFactory.Create(a_formatString)
    End If
    
    Dim p_binding As IPropertyBinding
    Dim p_strategy As IBindingTargetStrategy
    If TryGetBindingStrategyFor(a_target, p_strategy) Then
    
        If a_targetProperty = VBA.vbNullString Then
            
            Set p_binding = p_strategy.DefaultPropertyBindingFor(This.Context, _
                    Factory.NewBindingPath().Initialize(a_source, a_propertyPath), a_target, a_mode, _
                    a_updateTrigger, a_validator, a_converter, p_formatter, a_validationAdorner)
        
        Else
            
            Set p_binding = p_strategy.PropertyBindingFor(This.Context, _
                Factory.NewBindingPath().Initialize(a_source, a_propertyPath), _
                    Factory.NewBindingPath().Initialize(a_target, a_targetProperty), _
                    a_mode, a_updateTrigger, a_validator, a_converter, p_formatter, a_validationAdorner)
        
        End If
    Else
        
        Set p_binding = Factory.NewOneWayPropertyBinding().Initialize(This.Context, _
                Factory.NewBindingPath().Initialize(a_source, a_propertyPath), _
                Factory.NewBindingPath().Initialize(a_target, a_targetProperty), _
                a_validator:=a_validator, _
                a_converter:=a_converter, _
                a_stringFormatter:=p_formatter, _
                a_validationAdorner:=a_validationAdorner)
    
    End If
    
    If TypeOf a_source Is INotifyPropertyChanged Then
        
        Dim p_notifier As INotifyPropertyChanged
        Set p_notifier = a_source
        p_notifier.RegisterHandler p_binding
    
    End If
    
    This.PropertyBindings.Add p_binding
    Set IBindingManager_BindPropertyPath = p_binding
    If Not This.DebugOutput Then Exit Function
    
    On Error Resume Next
    If This.DebugOutput And a_targetProperty = VBA.vbNullString Then
        Debug.Print VBA.Information.TypeName(Me) & ": binding property path '" & a_propertyPath & _
            "' to the default-binding property of type '" & VBA.Information.TypeName(a_target) & "'."
    ElseIf This.DebugOutput Then
        Debug.Print VBA.Information.TypeName(Me) & ": binding property path '" & a_propertyPath & _
            "' to '" & VBA.Information.TypeName(a_target) & "." & a_targetProperty & "'."
    End If
    On Error GoTo 0
    
End Function

''' <summary>   Get a key to the list of strategies based to the specified object type. </summary>
''' <remarks>   Presently, type of MS Control is not an MS Control.  Potentially, It might be better to
'''             enumerate the binding target strategies by the control name as
'''             prefixing the control name with the library name is not robust.
'''             Note, though, the type of range indeed includes the Excel library name. </remarks>
''' <param name="a_target">   [Object] The target object, which code be an MS Forms control or
'''                           and Excel Range. </param>
''' <returns>   [String]. </returns>
Public Function GetKeyForObjectType(ByVal a_target As Object) As String

    Dim p_libraryName As String
    If TypeOf a_target Is MSForms.Control Then
        
        ' this does not capture controls such as check box.
        p_libraryName = "MSForms"
    
    ElseIf TypeOf a_target Is Excel.Range Then
        
        p_libraryName = "Excel"
    
    Else
        ' FIXME now what, prefix other target libraries here? need a better way.
        ' dh: defaulting to MS Forms for now.
        p_libraryName = "MSForms"
    
    End If

    GetKeyForObjectType = p_libraryName & IIf(p_libraryName = VBA.vbNullString, "", ".") & _
                         VBA.Information.TypeName(a_target)
    
End Function

Private Function TryGetBindingStrategyFor(ByVal a_target As Object, ByRef outStrategy As IBindingTargetStrategy) As Boolean
    
    Dim p_key As String: p_key = Me.GetKeyForObjectType(a_target)
    If This.BindingTargetStrategies.Exists(p_key) Then
        Set outStrategy = This.BindingTargetStrategies.Item(p_key)
        TryGetBindingStrategyFor = True
    Else
        ' not throwing an exception so as to test binding to a non-control element
        ' which then creates a one-way binding.
        ' see Binding Manager Tests TestNonControlTargetCreatesOneWayBinding
        TryGetBindingStrategyFor = False
#If False Then
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidArgumentError, _
                TypeName(Me) & "." & "TryGetBidingStrategyFor", _
                " Invalid key  '" & p_key & "' generated for target '" & TypeName(a_target)   '"."
#End If
    End If
    
End Function

Private Sub IBindingManager_Apply(ByVal a_source As Object)
    
    Apply a_source

End Sub

Private Property Get IBindingManager_DebugOutput() As Boolean
    IBindingManager_DebugOutput = This.DebugOutput
End Property

Private Property Get IBindingManager_NotifierFactory() As INotifierFactory
    
    Set IBindingManager_NotifierFactory = This.NotifierFactory

End Property

Private Property Get IBindingManager_StringFormatterFactory() As IStringFormatterFactory
    
    Set IBindingManager_StringFormatterFactory = This.StringFormatterFactory

End Property

Private Property Get IBindingManager_Validation() As IValidationManager
    
    Set IBindingManager_Validation = This.Context.Validation

End Property

' - - - - - - - - - - - - - - - -
'
' Disposable implementation
'
' - - - - - - - - - - - - - - - -

Private Sub IDisposable_Dispose()
    Set This.Context = Nothing
End Sub

' - - - - - - - - - - - - - - - -
'
' Handle Property Change implementation
'
' - - - - - - - - - - - - - - - -


Private Sub IHandlePropertyChanged_HandlePropertyChanged(ByVal a_source As Object, ByVal a_propertyName As String)
    This.Context.Commands.EvaluateCanExecute a_source
End Sub
