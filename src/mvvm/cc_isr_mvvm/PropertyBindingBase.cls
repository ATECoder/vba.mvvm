VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PropertyBindingBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit
Implements IPropertyBinding
Implements cc_isr_Core.IDisposable

Implements IControlEvents   '<~ relays MSForms.Control events to more specific IPropertyBinding implementations
Implements IHandleControlEvents '<~ ControlEventPunk callbacks
Implements IHandlePropertyChanged '<~ handles INotifyPropertyChange events from ViewModel (i.e. "applies" the binding)

Public Enum ApplyResult
    BindingSkipped
    BindingSuccess
    BindingValidationError
    BindingConversionError
    BindingFailed
End Enum

Private Type TState
    Applied As Boolean
    Applying As Boolean
    
    Punk As ControlEventsPunk
    Handlers As Collection
    
    Context As IAppContext
    Source As IBindingPath
    Target As IBindingPath
    
    Mode As BindingMode
    UpdateSourceTrigger As BindingUpdateSourceTrigger
    
    Converter As IValueConverter
    StringFormatter As IStringFormatter
    
    Validator As IValueValidator
    ValidationAdorner As IDynamicAdorner ' data validations (might be unset)
    CancelExitOnValidationError As Boolean
    
End Type

Private This As TState

''' <summary>   Initializes and returns an <see cref="IPropertyBinding"> representation of
'''             this <see cref="PropertyBindingBase"/> object. </summary>
''' <remarks>
''' this method initializes the <c>IBindingPath</c> objects from the given context and property
''' path strings, which makes it simpler for the calling code. Note that the target property path
''' is Optional, how is that possible?
'''
''' If you’ve spent a bit of time with the MVVM prototype code, you probably noticed the <c>PropertyBinding</c>
''' class was pretty much out of control, and extending it to support more target types would only make it worse.
''' So what I did is, I pulled the common code into a new <c>PropertyBindingBase</c> class, then moved the
''' control-specific code into its own specialized control-specific property binding implementation, and now
''' there’s a strategy pattern that’s responsible for fetching the correct implementation – so that’s how binding
''' a <c>TextBox</c> target creates a <c>TextBoxPropertyBinding</c>, and how binding a <c>Label</c> target creates
''' a <c>OneWayPropertyBinding<c>. Each control-specific binding class can handle that control’s events and implement
''' control-specific behavior accordingly.
''' </remarks>
''' <param name="a_context"> </param>
''' <param name="a_source">                        [IBindingPath] Old description: Typically, that’s just a reference to
'''                                                the source ViewModel object. </param>
''' <param name="a_targetContext">                 [Object] Typically, that’s just a reference to the target MSForms control.
'''                                                Technically, it could really be any object that has any number of public
'''                                                properties. </param>
''' <param name="a_targetPropertyPath">            [String] Each binding type has a default target property that is
'''                                                automatically inferred from the type of target (and in some cases,
'''                                                from the data type of the source property). For example, binding to a
'''                                                <c>TextBox</c> control automatically wires up the control’s <c>Text</c>
'''                                                property, such that providing a <c>TargetPropertyPath</c> is only needed
'''                                                when binding to another target property. </param>
''' <param name="a_validationAdorner">             [IDynamicAdorner] Lets the user know about validation errors. </param>
''' <param name="a_cancelExitOnValidationError">   [Boolean, True] Cancel exit on validation error. </param>
''' <returns>   [<see cref="IPropertyBinding"/>]. </returns>
Public Function Initialize(ByVal a_context As IAppContext, ByVal a_source As IBindingPath, _
        ByVal a_targetContext As Object, _
        Optional ByVal a_targetPropertyPath As String = "Value", _
        Optional ByVal a_mode As BindingMode = BindingMode.TwoWayBinding, _
        Optional ByVal a_updateSource As BindingUpdateSourceTrigger = BindingUpdateSourceTrigger.OnExit, _
        Optional ByVal a_converter As IValueConverter, _
        Optional ByVal a_stringFormatter As IStringFormatter, _
        Optional ByVal a_validator As IValueValidator, _
        Optional ByVal a_validationAdorner As IDynamicAdorner, _
        Optional ByVal a_cancelExitOnValidationError As Boolean = True) As IPropertyBinding

    Me.CancelExitOnValidationError = a_cancelExitOnValidationError
    Me.Mode = a_mode
    Me.UpdateSourceTrigger = a_updateSource
    
    Set Me.Source = a_source
    Set Me.Target = Factory.NewBindingPath().Initialize(a_targetContext, a_targetPropertyPath)
    
    Set Me.Context = a_context
    
    Set Me.Converter = a_converter
    Set Me.StringFormatter = a_stringFormatter
    
    Set Me.Validator = a_validator
    Set Me.ValidationAdorner = a_validationAdorner
    
    Set Initialize = Me
    
End Function

Public Property Get AsIControlEvents() As IControlEvents
    Set AsIControlEvents = Me
End Property

Public Property Get AsINotifyValidationError() As INotifyValidationError
    Set AsINotifyValidationError = Me
End Property

Public Property Get Source() As IBindingPath
    Set Source = This.Source
End Property

Public Property Set Source(ByVal a_value As IBindingPath)
    cc_isr_Core_IO.GuardClauses.GuardDoubleInitialization This.Source, _
        ThisWorkbook.VBProject.Name & "." & VBA.Information.TypeName(Me) & ".Source"
    cc_isr_Core_IO.GuardClauses.GuardNullReference a_value, _
        ThisWorkbook.VBProject.Name & "." & VBA.Information.TypeName(Me) & ".Source"
    Set This.Source = a_value
End Property

Public Property Get Target() As IBindingPath
    Set Target = This.Target
End Property

Public Property Set Target(ByVal a_value As IBindingPath)
    cc_isr_Core_IO.GuardClauses.GuardDoubleInitialization This.Target, _
        ThisWorkbook.VBProject.Name & "." & VBA.Information.TypeName(Me) & ".Target"
    cc_isr_Core_IO.GuardClauses.GuardNullReference a_value, _
        ThisWorkbook.VBProject.Name & "." & VBA.Information.TypeName(Me) & ".Target"
    Set This.Target = a_value
    If TypeOf This.Target.Context Is MSForms.Control Then
        On Error Resume Next
        Set This.Punk.Target = This.Target.Context
        If This.Punk.Connect Then
            This.Punk.RegisterHandler Me
        Else
            DebugMessage "Unable to connect target IConnectionPoint."
        End If
        If Err.Number <> 0 Then
            DebugMessage "Unable to handle MSForms.Control events. Binding is effectively one-way. " & Err.Description
        End If
        On Error GoTo 0
    Else
        Set This.Punk = Nothing
    End If
End Property

Public Property Get Mode() As BindingMode
    Mode = This.Mode
End Property

Public Property Let Mode(ByVal a_value As BindingMode)
    This.Mode = a_value
End Property

Public Property Get UpdateSourceTrigger() As BindingUpdateSourceTrigger
    UpdateSourceTrigger = This.UpdateSourceTrigger
End Property

Public Property Let UpdateSourceTrigger(ByVal a_value As BindingUpdateSourceTrigger)
    This.UpdateSourceTrigger = a_value
End Property

Public Property Get Converter() As IValueConverter
    Set Converter = This.Converter
End Property

Public Property Set Converter(ByVal a_value As IValueConverter)
    cc_isr_Core_IO.GuardClauses.GuardDoubleInitialization This.Converter, _
        ThisWorkbook.VBProject.Name & "." & VBA.Information.TypeName(Me) & ".Converter"
    Set This.Converter = a_value
End Property

Public Property Get Validator() As IValueValidator
    Set Validator = This.Validator
End Property

Public Property Set Validator(ByVal a_value As IValueValidator)
    cc_isr_Core_IO.GuardClauses.GuardDoubleInitialization This.Validator, _
        ThisWorkbook.VBProject.Name & "." & VBA.Information.TypeName(Me) & ".Validator"
    Set This.Validator = a_value
    If Not This.Validator Is Nothing Then This.UpdateSourceTrigger = This.Validator.Trigger
End Property

Public Property Get StringFormatter() As IStringFormatter
    Set StringFormatter = This.StringFormatter
End Property

Public Property Set StringFormatter(ByVal a_value As IStringFormatter)
    cc_isr_Core_IO.GuardClauses.GuardDoubleInitialization This.StringFormatter, _
        ThisWorkbook.VBProject.Name & "." & VBA.Information.TypeName(Me) & ".StringFormatter"
    Set This.StringFormatter = a_value
End Property

Public Property Get ValidationAdorner() As IDynamicAdorner
    Set ValidationAdorner = This.ValidationAdorner
End Property

Public Property Set ValidationAdorner(ByVal a_value As IDynamicAdorner)
    cc_isr_Core_IO.GuardClauses.GuardDoubleInitialization This.ValidationAdorner, _
        ThisWorkbook.VBProject.Name & "." & VBA.Information.TypeName(Me) & ".ValidationAdorner"
    Set This.ValidationAdorner = a_value
End Property

Public Property Get Context() As IAppContext
    Set Context = This.Context
End Property

Public Property Set Context(ByVal a_value As IAppContext)
    cc_isr_Core_IO.GuardClauses.GuardDoubleInitialization This.Context, _
        ThisWorkbook.VBProject.Name & "." & VBA.Information.TypeName(Me) & ".Context"
    cc_isr_Core_IO.GuardClauses.GuardNullReference a_value, _
        ThisWorkbook.VBProject.Name & "." & VBA.Information.TypeName(Me) & ".Context"
    Set This.Context = a_value
End Property

Public Property Get CancelExitOnValidationError() As Boolean
    CancelExitOnValidationError = This.CancelExitOnValidationError
End Property

Public Property Let CancelExitOnValidationError(ByVal a_value As Boolean)
    This.CancelExitOnValidationError = a_value
End Property

Private Function TryConvert(ByRef a_value As Variant) As Boolean
    
    If This.Converter Is Nothing Then
        TryConvert = True
        Exit Function
    End If
    
    On Error Resume Next
    a_value = This.Converter.Convert(a_value)
    If Err.Number = 0 Then
        DebugMessage "Value was successfully converted."
        TryConvert = True
    Else
        DebugMessage "IValueConverter.Convert raised an error: " & Err.Description
    End If
    On Error GoTo 0
    
End Function

Private Function TryConvertBack(ByRef a_value As Variant) As Boolean
    
    If This.Converter Is Nothing Then
        TryConvertBack = True
        Exit Function
    End If
    
    On Error Resume Next
    a_value = This.Converter.ConvertBack(a_value)
    If Err.Number = 0 Then
        DebugMessage "Value was successfully converted back."
        TryConvertBack = True
    Else
        DebugMessage "IValueConverter.ConvertBack raised an error: " & Err.Description
    End If
    On Error GoTo 0
    
End Function

''' <summary>   Applies the binding to the source to the target. </summary>
Public Sub Apply()
    
    ' <~ could break accidental recursion... leave this here...
    
    If This.Applying Then Exit Sub
    
    If This.Mode = OneTimeBinding Then
    
        If Not This.Applied Then Me.ApplyToTarget
        
    ElseIf This.Mode = OneWayBinding Or This.Mode = TwoWayBinding Then
    
        Me.ApplyToTarget
        
    ElseIf This.Mode = OneWayToSource Then
    
        Me.ApplyToSource
        
    End If

End Sub

Private Property Get CanApplyToTarget() As Boolean

    ' one-way to source and already-applied one-time binding modes do not apply to target
    CanApplyToTarget = Not This.Applying And This.Mode <> OneWayToSource And (This.Mode <> OneTimeBinding Or Not This.Applied)

End Property


''' <summary>   Reads from the source and writes to the target. </summary>
''' <remarks>
''' When it’s a ViewModel property that changes, the <c>Property Let</c> procedure invokes
''' <c>INotifyPropertyChanged.NotifyPropertyChanged</c>, the property binding was registered as an observer,
''' so <c>IHandlePropertyChanged_HandlePropertyChanged</c> is invoked; whenever the provided <c>Source</c>
''' is the binding source and the provided <c>PropertyName</c> is the source property name, <c>ApplyToTarget<cc> gets invoked.
''' </remarks>
Public Sub ApplyToTarget()

    
    If Not CanApplyToTarget Then Exit Sub
    This.Applying = True
    
    Dim p_sourceValue As Variant
    Dim UseFallbackDefault As Boolean
    
    If This.Source.Object Is Nothing Then
        UseFallbackDefault = TryGetDefaultBindingValue(a_outValue:=p_sourceValue)
    End If
    
    This.Source.Resolve
    
    If This.Source.Object Is Nothing And Not UseFallbackDefault Then
        DebugMessage "ApplyToTarget aborted: source object in path '" & This.Source.Path & "' is Nothing and there is no fallback default value for this binding."
        This.Applying = False
        Exit Sub
        
    ElseIf This.Source.Object Is Nothing Then
        DebugMessage "ApplyToTarget: source object in path '" & This.Source.Path & "' is Nothing; binding target property '" & This.Target.PropertyName & "' to default/fallback value."
        
    Else
        If Not This.Source.TryReadPropertyValue(a_outValue:=p_sourceValue) Then
            DebugMessage "ApplyToTarget failed to read current source property value."
            This.Applying = False
            Exit Sub
        End If
    End If
    
    Validate p_sourceValue, False
    p_sourceValue = Convert(p_sourceValue)
    
    This.Target.Resolve
    Dim CurrentValue As Variant
    If Not This.Target.TryReadPropertyValue(a_outValue:=CurrentValue) Then
        DebugMessage "ApplyToTarget failed to read current target property value."
        This.Applying = False
        Exit Sub
    End If
    
    Dim Proceed As Boolean
    If Not VBA.IsArray(p_sourceValue) Then
        Proceed = p_sourceValue <> CurrentValue
    Else
        Proceed = True
    End If
    
    If Proceed Then
        
        If Not This.StringFormatter Is Nothing Then
            p_sourceValue = This.StringFormatter.Format(p_sourceValue)
        End If
        
        If Not This.Target.TryWritePropertyValue(p_sourceValue) Then
            DebugMessage "ApplyToTarget failed to write to target property."
        Else
            DebugMessage "ApplyToTarget successfully applied to target."
            This.Applied = True
        End If
    End If
    This.Applying = False
End Sub

Private Property Get CanApplyToSource() As Boolean
    'one-time and one-way binding modes do not write values to the binding source (ViewModel).
    CanApplyToSource = Not This.Applying And This.Mode <> OneTimeBinding And This.Mode <> OneWayBinding
End Property

''' <summary>   Reads from the target and writes to the source.
''' ?? when This.Mode=KeyPress, the a_targetValidationValue contains the KeyAscii value.
''' </summary>
''' <remarks>
''' Reads the value from the binding target, and then writes that value to the binding source.
''' If the binding’s Mode is <c>OneTimeBinding</c> or <c>OneWayBinding</c>, we can already bail
''' out because these modes only ever write to the binding target.
''' <see href="https://rubberduckvba.blog/2020/10/25/making-mvvm-work-in-vba-part-3-bindings/"/>
''' </remarks>
Public Function ApplyToSource(Optional ByVal a_targetValidationValue As Variant) As ApplyResult

    On Error GoTo CleanFail

    If Not CanApplyToSource Then Exit Function
    This.Applying = True
    
    This.Target.Resolve
    Dim TargetValue As Variant
    If Not This.Target.TryReadPropertyValue(a_outValue:=TargetValue) Then
        DebugMessage "ApplyToSource failed to read target property value."
        ApplyToSource = ApplyResult.BindingFailed
        This.Applying = False
        Exit Function
    End If
    
    If VBA.IsMissing(a_targetValidationValue) Then
        a_targetValidationValue = TargetValue
        
    ElseIf This.UpdateSourceTrigger = BindingUpdateSourceTrigger.OnKeyPress Then
        'keypress hasn't reached the control yet:
        TargetValue = InsertKeypress(TargetValue, a_targetValidationValue)
        
    End If
    
    If Not TryConvertBack(TargetValue) Then
        ApplyToSource = ApplyResult.BindingConversionError
        DebugMessage VBA.Information.TypeName(This.Converter) & " ConvertBack failed to convert the target value."
        OnEnter
        This.Applying = False
        Exit Function
    End If
    
    If Not Validate(a_targetValidationValue) Then
        ApplyToSource = ApplyResult.BindingValidationError
        DebugMessage VBA.Information.TypeName(This.Validator) & " failed to validate the target value."
        This.Applying = False
        Exit Function
    End If
    
    This.Source.Resolve
    Dim CurrentValue As Variant
    If Not This.Source.TryReadPropertyValue(a_outValue:=CurrentValue) Then
        DebugMessage "ApplyToSource failed to read current source property value."
        ApplyToSource = ApplyResult.BindingFailed
        This.Applying = False
        Exit Function
    End If
    
    Dim Proceed As Boolean
    If Not VBA.IsArray(CurrentValue) Then
        Proceed = TargetValue <> CurrentValue
    Else
        Proceed = True
    End If
    
    If Proceed Then
        
        ClearValidationErrors
        
        If Not This.Source.TryWritePropertyValue(TargetValue) Then
            DebugMessage "ApplyToSource failed to write to source property."
            ApplyToSource = ApplyResult.BindingFailed
            Exit Function
        End If
        
        DebugMessage "Binding was successfully applied to source."
        ApplyToSource = ApplyResult.BindingSuccess
        
    End If
    
CleanExit:
    This.Applying = False
    Exit Function
    
CleanFail:
    DebugMessage "ApplyToSource failed unexpectedly."
    ApplyToSource = ApplyResult.BindingFailed
    Resume CleanExit
    Resume
End Function

Private Function InsertKeypress(ByVal a_initialValue As String, ByVal a_character As String) As String

    'FIXME this function is at the wrong abstraction level.
    
    If a_character = VBA.vbNullString Then
        InsertKeypress = a_initialValue
        Exit Function
    End If
    
    On Error Resume Next
    
    Dim p_selectionStart As Long
    p_selectionStart = This.Target.Object.SelStart
    
    Dim p_selectionLength As Long
    p_selectionLength = This.Target.Object.SelLength
    
    On Error GoTo 0
    
    Select Case True
        
        Case p_selectionStart > 0 And p_selectionLength = 0
            
            ' insert at SelStart
            InsertKeypress = cc_isr_Core.StringExtensions.Insert(a_initialValue, a_character, p_selectionStart)
        
        Case p_selectionStart > 0 And p_selectionLength > 0
            
            ' replace selection
            Dim p_remove As String
            p_remove = cc_isr_Core.StringExtensions.Substring(a_initialValue, p_selectionStart, p_selectionLength)
            InsertKeypress = VBA.Strings.Replace(a_initialValue, p_remove, a_character)
            
        Case Else
            
            'replace initial value
            InsertKeypress = a_character
    
    End Select
    
End Function

Private Function IsValid() As Boolean
    IsValid = This.Context.Validation.IsValid(This.Source.Object, This.Source.PropertyName)
End Function

Private Sub ClearValidationErrors()
    With This.Context
        .Validation.ClearValidationError This.Source
        .Commands.EvaluateCanExecute This.Source
    End With
End Sub

''' <summary>   Returns a string that represents the current object. </summary>
Public Function ToString() As String
    ToString = VBA.Information.TypeName(This.Source.Context) & "." & This.Source.Path & " -> " & VBA.Information.TypeName(This.Target.Context) & "." & This.Target.Path
End Function

Private Function ResolvePropertyPath(ByVal a_source As Object, ByVal a_propertyPath As String) As Object
    
    Dim p_parts As Variant
    p_parts = Strings.Split(a_propertyPath, ".")
    
    If UBound(p_parts) = LBound(p_parts) Then
        Set ResolvePropertyPath = a_source
    Else
        Dim p_recursiveProperty As Object
        Set p_recursiveProperty = CallByName(a_source, p_parts(0), VbGet)
        If p_recursiveProperty Is Nothing Then Exit Function
        Set ResolvePropertyPath = ResolvePropertyPath(p_recursiveProperty, Right$(a_propertyPath, Len(a_propertyPath) - Len(p_parts(0)) - 1))
    End If
    
End Function

Private Function ResolvePropertyName(ByVal a_propertyPath As String) As String
    Dim p_parts As Variant
    p_parts = Strings.Split(a_propertyPath, ".")
    ResolvePropertyName = p_parts(UBound(p_parts))
End Function

''' <summary>   Gets a default value for certain specific target properties, used when source path cannot be
'''             fully resolved, e.g., when target binds to <c>SomeObjectProperty.SomeProperty</c> and
'''             <c>SomeObjectProperty</c> is Nothing. </summary>
Private Function TryGetDefaultBindingValue(ByRef a_outValue As Variant) As Boolean

    Select Case This.Target.PropertyName
    
        Case "Text", "Caption"
            
            a_outValue = VBA.vbNullString
            TryGetDefaultBindingValue = True
            
        Case "Enabled", "Visible"
        
            a_outValue = False
            TryGetDefaultBindingValue = True
            
        Case "Value"
            
            If TypeOf This.Target Is MSForms.CheckBox Or TypeOf This.Target Is MSForms.OptionButton Then
                a_outValue = False
                TryGetDefaultBindingValue = True
            End If
            
    End Select
    
End Function

Private Function Convert(ByVal a_value As Variant) As Variant
    
    On Error GoTo CleanFail
    If Not This.Converter Is Nothing Then
        Convert = This.Converter.Convert(a_value)
    Else
        Convert = a_value
    End If
    
CleanExit:
    Exit Function
    
CleanFail:
    DebugMessage "IValueConverter failed with error " & Err.Number & ": " & Err.Description
    OnEnter
    With VBA.Information.Err
        If .Number <> 0 Then
            Debug.Print "Error " & .Number, .Description
            .Raise .Number
        End If
    End With
    ' resume execution after error-handling is finished.
    Resume CleanExit 'NOTE: ' VF: is never reached
End Function

Public Function Validate(ByVal a_targetValue As Variant, Optional ByVal a_propagate As Boolean = True) As Boolean

    If This.Validator Is Nothing Then
        Validate = True
        Exit Function
    End If
    
    On Error GoTo CleanFail
    If This.Validator.IsValid(a_targetValue, This.Source, This.Target) Then
        
        If a_propagate Then
            This.Context.Validation.ClearValidationError This.Source
            If Not This.ValidationAdorner Is Nothing Then This.ValidationAdorner.Hide
        End If
        
        Validate = True
        
    Else
        
        If a_propagate Then
            Dim p_message As String
            p_message = This.Validator.Message
            
            This.Context.Validation.OnValidationError _
                a_context:=This.Source.Context, _
                a_validationError:=Factory.NewValidationError().Initialize(Me, p_message)
            
            If Not This.ValidationAdorner Is Nothing Then This.ValidationAdorner.Show p_message
            ' VF: commented out as it leads to select all, which is not desired in text and numeric text
            '    boxes even though there may be more than one event handler
            'OnEnter
        End If
    End If
        
CleanExit:
    If a_propagate Then
        This.Context.Commands.EvaluateCanExecute This.Source.Context
    End If
    Exit Function
    
CleanFail:
    DebugMessage "IValueValidator failed with error " & Err.Number & ": " & Err.Description
    OnEnter
    Resume CleanExit
    Resume
End Function

Private Sub Format()
    If This.StringFormatter Is Nothing Then Exit Sub

    Dim BindingValue As Variant
    If Not This.Source.TryReadPropertyValue(a_outValue:=BindingValue) Then Exit Sub
        
    On Error GoTo CleanFail
    Dim FormattedValue As String
    FormattedValue = This.StringFormatter.Format(BindingValue)
    On Error GoTo 0
    
    This.Applying = True
    If Not This.Target.TryWritePropertyValue(FormattedValue) Then
        DebugMessage "Could not format the binding value."
    End If
    This.Applying = False
    
CleanExit:
    Exit Sub
CleanFail:
    DebugMessage "IStringFormatter.Format() raised an error: " & Err.Description
    Resume CleanExit
End Sub

Private Sub OnAfterUpdate()
    Format
    Dim p_handler As IHandleControlEvents
    For Each p_handler In This.Handlers
        p_handler.HandleAfterUpdate
    Next
End Sub

Private Sub OnBeforeUpdate(ByRef a_cancel As Boolean)
    OnBindingUpdateSourceOpportunity a_cancel, BindingUpdateSourceTrigger.OnPropertyChanged
    Dim p_handler As IHandleControlEvents
    For Each p_handler In This.Handlers
        p_handler.HandleBeforeUpdate a_cancel
        If a_cancel Then Exit For
    Next
End Sub
' VF: added: required to capture Cut/Paste and Backspace/Delete in TextBox
Private Sub OnChange()
    Format
    Dim p_handler As IHandleControlEvents
    For Each p_handler In This.Handlers
        p_handler.HandleChange
    Next
End Sub

Private Sub OnEnter()
    Dim p_handler As IHandleControlEvents
    For Each p_handler In This.Handlers
        p_handler.HandleEnter
    Next
End Sub

Private Sub OnExit(ByRef a_cancel As Boolean)
    OnBindingUpdateSourceOpportunity a_cancel, BindingUpdateSourceTrigger.OnExit
    Dim p_handler As IHandleControlEvents
    For Each p_handler In This.Handlers
        p_handler.HandleExit a_cancel
    Next
End Sub

Public Sub OnBindingUpdateSourceOpportunity(ByRef a_cancel As Boolean, _
        ByVal a_trigger As cc_isr_MVVM.BindingUpdateSourceTrigger, Optional ByVal TargetValidationValue As Variant)
        
    If Not CanApplyToSource Then Exit Sub
    
    Dim p_result As cc_isr_MVVM.ApplyResult
    Select Case a_trigger
        
        ' VF: ? remove OnKeyPress here?, at least in TextBox no longer used
        ' VF: changed: required to capture Cut/Paste and Backspace/Delete in TextBox
        'Case BindingUpdateSourceTrigger.OnKeyPress
        Case BindingUpdateSourceTrigger.OnKeyPress, BindingUpdateSourceTrigger.OnChange
        
            'If This.UpdateSourceTrigger = OnKeyPress And Not VBA.IsMissing(TargetValidationValue) Then
            If (This.UpdateSourceTrigger = OnKeyPress Or This.UpdateSourceTrigger = cc_isr_MVVM.BindingUpdateSourceTrigger.OnChange) _
                And Not VBA.IsMissing(TargetValidationValue) Then
                'just validate the keypress ' VF: or string pasted
                p_result = IIf(Validate(TargetValidationValue), _
                    ApplyResult.BindingSuccess, _
                    ApplyResult.BindingValidationError)
                            
            End If
            
        Case BindingUpdateSourceTrigger.OnPropertyChanged
            
            If This.UpdateSourceTrigger = OnPropertyChanged Or This.UpdateSourceTrigger = OnKeyPress Then
                p_result = ApplyToSource
            End If
        
        Case BindingUpdateSourceTrigger.OnExit
            
            If This.UpdateSourceTrigger <> Never Then
                
                Dim p_outValue As Variant
                If This.Target.TryReadPropertyValue(p_outValue) Then
                
                    p_result = IIf(Validate(p_outValue), _
                        ApplyResult.BindingSuccess, _
                        ApplyResult.BindingValidationError)
                        
                End If
            End If
            
    End Select
    
    a_cancel = This.CancelExitOnValidationError And (a_cancel Or (p_result = BindingValidationError) Or _
            (p_result = BindingConversionError))
    
End Sub

Private Sub OnValidationError()

    This.Context.Validation.OnValidationError This.Source.Context, _
        Factory.NewValidationError().Initialize(Me, This.Validator.Message)
        
    This.ValidationAdorner.Show This.Validator.Message
    
End Sub

Private Sub DebugMessage(ByVal a_message As String)

    If This.Context.Bindings.DebugOutput Then Debug.Print "PropertyBinding: [" & Me.ToString() & "] " & a_message
    
End Sub

Private Sub Class_Initialize()
    Set This.Handlers = New Collection
    Set This.Punk = New ControlEventsPunk
End Sub

Private Sub Class_Terminate()
    Set This.Handlers = Nothing
    If Not This.Punk Is Nothing Then
        This.Punk.Disconnect
        Set This.Punk = Nothing
    End If
End Sub

Private Sub IControlEvents_OnAfterUpdate()
    OnAfterUpdate
End Sub

Private Sub IControlEvents_OnBeforeUpdate(ByVal a_cancel As MSForms.IReturnBoolean)
    Dim p_cancel As Boolean
    p_cancel = a_cancel.Value
    OnBeforeUpdate p_cancel
    a_cancel.Value = p_cancel
End Sub
' VF: added: required to capture Cut/Paste and Backspace/Delete in TextBox
Private Sub IControlEvents_OnChange()
    OnChange
End Sub

Private Sub IControlEvents_OnEnter()
    OnEnter
End Sub

Private Sub IControlEvents_OnExit(ByVal a_cancel As MSForms.IReturnBoolean)
    Dim p_cancel As Boolean
    p_cancel = a_cancel.Value
    OnExit p_cancel
    a_cancel.Value = p_cancel
End Sub

Private Sub IControlEvents_RegisterHandler(ByVal a_handler As IHandleControlEvents)
    This.Handlers.Add a_handler
End Sub

' - - - - - - - - - - - - - - - -
'
' Disposable implementation
'
' - - - - - - - - - - - - - - - -

Private Sub IDisposable_Dispose()
    
    Set This.Context = Nothing
    
    This.Punk.Disconnect
    Set This.Punk = Nothing
    cc_isr_Core.DisposableExtensions.TryDispose This.ValidationAdorner
    Set This.ValidationAdorner = Nothing
    
End Sub

Private Sub IHandleControlEvents_HandleAfterUpdate()
    OnAfterUpdate
End Sub

Private Sub IHandleControlEvents_HandleBeforeUpdate(ByRef a_cancel As Boolean)
    OnBeforeUpdate a_cancel
End Sub
' VF: added: required to capture Cut/Paste and Backspace/Delete in TextBox
Private Sub IHandleControlEvents_HandleChange()
    OnChange
End Sub

Private Sub IHandleControlEvents_HandleEnter()
    OnEnter
End Sub

Private Sub IHandleControlEvents_HandleExit(ByRef a_cancel As Boolean)
    OnExit a_cancel
End Sub

Private Sub IHandlePropertyChanged_HandlePropertyChanged(ByVal a_source As Object, ByVal a_propertyName As String)
    'IHandlePropertyChanged -> INotifyPropertyChanged: source is the ViewModel, so we're propagating to Target here:
    If Not This.Applying And a_source Is This.Source.Object And a_propertyName = This.Source.PropertyName Then
        ApplyToTarget
    End If
End Sub

Private Sub IPropertyBinding_Apply()
    Apply
End Sub

Private Property Get IPropertyBinding_CancelExitOnValidationError() As Boolean
    IPropertyBinding_CancelExitOnValidationError = This.CancelExitOnValidationError
End Property

Private Property Get IPropertyBinding_Converter() As IValueConverter
    Set IPropertyBinding_Converter = This.Converter
End Property

Private Property Get IPropertyBinding_DefaultTargetProperty() As String
    IPropertyBinding_DefaultTargetProperty = BindingDefaults.BindingBaseDefaultProperty
End Property

Private Property Get IPropertyBinding_Mode() As BindingMode
    IPropertyBinding_Mode = This.Mode
End Property

Private Property Get IPropertyBinding_Source() As IBindingPath
    Set IPropertyBinding_Source = This.Source
End Property

Private Property Get IPropertyBinding_StringFormatter() As IStringFormatter
    Set IPropertyBinding_StringFormatter = This.StringFormatter
End Property

Private Property Get IPropertyBinding_Target() As IBindingPath
    Set IPropertyBinding_Target = This.Target
End Property

Private Property Get IPropertyBinding_UpdateSourceTrigger() As BindingUpdateSourceTrigger
    IPropertyBinding_UpdateSourceTrigger = This.UpdateSourceTrigger
End Property

Private Property Get IPropertyBinding_Validator() As IValueValidator
    Set IPropertyBinding_Validator = This.Validator
End Property


