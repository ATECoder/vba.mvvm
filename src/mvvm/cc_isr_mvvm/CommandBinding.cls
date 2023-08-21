VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommandBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An object responsible for binding a command to a UI element."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   An object responsible for binding a command to a UI element. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Implements ICommandBinding
Implements cc_isr_Core.IDisposable
Option Explicit

Private Type TCommandBinding
    ViewModel As Object
    Target As Object
    Command As ICommand
End Type

Private WithEvents CommandButtonEvents As MSForms.CommandButton
Attribute CommandButtonEvents.VB_VarHelpID = -1
Private WithEvents CheckBoxEvents As MSForms.CheckBox
Attribute CheckBoxEvents.VB_VarHelpID = -1
Private WithEvents ImageEvents As MSForms.Image
Attribute ImageEvents.VB_VarHelpID = -1
Private WithEvents LabelEvents As MSForms.Label
Attribute LabelEvents.VB_VarHelpID = -1

Private This As TCommandBinding

''' <summary>   Initializes and returns an <see cref="IPropertyBinding"/> representation of
'''             this <see cref="CommandBinding"/> object. </summary>
''' <param name="a_target">       [Object] The command target object, e.g., <see cref="MSForms.CommandButton"/>. </param>
''' <param name="a_command">      [ICommand] an object representing the command.  </param>
''' <param name="a_viewModel">    [Object] the ViewModel object. </param>
''' <returns>    [ICommandBinding]. </returns>
Friend Function Initialize(ByVal a_target As Object, ByVal a_command As ICommand, _
        ByVal a_viewModel As Object) As ICommandBinding
    
    Set Me.ViewModel = a_viewModel
    Set Me.Target = a_target
    Set Me.Command = a_command
    Set Initialize = Me
    
End Function

''' <summary>   Initializes and returns this <see cref="CommandBinding"/> object
'''             for a <see cref="MSForms.CommandButton"/> target. </summary>
''' <param name="a_target">       [<see cref="MSForms.CommandButton"/>] The command target object. </param>
''' <param name="a_command">      [ICommand] an object representing the command.  </param>
''' <param name="a_viewModel">    [Object] the ViewModel object. </param>
''' <returns>    [ICommandBinding]. </returns>
Public Function ForCommandButton(ByVal a_target As MSForms.CommandButton, ByVal a_command As ICommand, _
        ByVal a_viewModel As Object) As ICommandBinding
    
    Set ForCommandButton = Me.Initialize(a_target, a_command, a_viewModel)

End Function

''' <summary>   Initializes and returns this <see cref="CommandBinding"/> object
'''             for a <see cref="MSForms.CheckBox"/> target. </summary>
''' <param name="a_target">       [<see cref="MSForms.CheckBox"/>] The command target object. </param>
''' <param name="a_command">      [ICommand] an object representing the command.  </param>
''' <param name="a_viewModel">    [Object] the ViewModel object. </param>
''' <returns>    [ICommandBinding]. </returns>
Public Function ForCheckBox(ByVal a_target As MSForms.CheckBox, ByVal a_command As ICommand, _
        ByVal a_viewModel As Object) As ICommandBinding
    
    Set ForCheckBox = Me.Initialize(a_target, a_command, a_viewModel)

End Function

''' <summary>   Initializes and returns this <see cref="CommandBinding"/> object
'''             for a <see cref="MSForms.Image"/> target. </summary>
''' <param name="a_target">       [<see cref="MSForms.Image"/>] The command target object. </param>
''' <param name="a_command">      [ICommand] an object representing the command.  </param>
''' <param name="a_viewModel">    [Object] the ViewModel object. </param>
''' <returns>    [ICommandBinding]. </returns>
Public Function ForImage(ByVal a_target As MSForms.Image, ByVal a_command As ICommand, _
        ByVal a_viewModel As Object) As ICommandBinding
    
    Set ForImage = Me.Initialize(a_target, a_command, a_viewModel)

End Function

''' <summary>   Initializes and returns this <see cref="CommandBinding"/> object
'''             for a <see cref="MSForms.Label"/> target. </summary>
''' <param name="a_target">       [<see cref="MSForms.Label"/>] The command target object. </param>
''' <param name="a_command">      [ICommand] an object representing the command.  </param>
''' <param name="a_viewModel">    [Object] the ViewModel object. </param>
''' <returns>    [ICommandBinding]. </returns>
Public Function ForLabel(ByVal a_target As MSForms.Label, ByVal a_command As ICommand, _
        ByVal a_viewModel As Object) As ICommandBinding
    
    Set ForLabel = Me.Initialize(a_target, a_command, a_viewModel)

End Function

Public Property Get ViewModel() As Object
    Set ViewModel = This.ViewModel
End Property

Public Property Set ViewModel(ByVal a_value As Object)
    Set This.ViewModel = a_value
End Property

Public Property Get Target() As Object
    Set Target = This.Target
End Property

Public Property Set Target(ByVal a_value As Object)
    
    Dim p_source As String
    p_source = ThisWorkbook.VBProject.Name & "." & VBA.Information.TypeName(Me) & ".Target"
    cc_isr_Core_IO.GuardClauses.GuardDoubleInitialization This.Target, p_source
    Set This.Target = a_value
    
    Select Case True
        Case TypeOf a_value Is MSForms.CommandButton
            Set CommandButtonEvents = a_value
            
        Case TypeOf a_value Is MSForms.CheckBox
            Set CheckBoxEvents = a_value
            
        Case TypeOf a_value Is MSForms.Image
            
            Set ImageEvents = a_value
            
        Case TypeOf a_value Is MSForms.Label
            
            Set LabelEvents = a_value
            
        Case Else
        
            cc_isr_Core_IO.GuardClauses.GuardExpression True, p_source, _
                "Type '" & VBA.Information.TypeName(a_value) & _
                "' does not support command bindings at the moment.", _
                cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError
                
    End Select
    
End Property

Public Property Get Command() As ICommand
    Set Command = This.Command
End Property

Public Property Set Command(ByVal a_value As ICommand)
    
    Dim p_source As String
    p_source = ThisWorkbook.VBProject.Name & "." & VBA.Information.TypeName(Me) & ".Command"
    cc_isr_Core_IO.GuardClauses.GuardDoubleInitialization This.Command, p_source
    cc_isr_Core_IO.GuardClauses.GuardNullReference a_value, p_source
    Set This.Command = a_value

End Property

Private Property Get AsInterface() As ICommandBinding
    Set AsInterface = Me
End Property

Private Sub OnExecute()
    If Not This.Command Is Nothing Then
        This.Command.Execute This.ViewModel
    Else
        Debug.Print "BUG in " & VBA.Information.TypeName(Me) & ": Command is 'Nothing', cannot execute."
        Debug.Assert False ' should not happen, break here if it does.
    End If
End Sub

Private Sub EvaluateCanExecute(ByVal a_source As Object)
    If This.Target Is Nothing Then Exit Sub
    If This.Command Is Nothing Then
        This.Target.Enabled = False
    Else
        On Error Resume Next
        This.Target.Enabled = This.Command.CanExecute(a_source)
        'a validation error formatter might have taken over the tooltip?
        This.Target.ControlTipText = IIf(This.Target.ControlTipText = VBA.vbNullString, This.Command.Description, This.Target.ControlTipText)
        On Error GoTo 0
    End If
End Sub

''' <summary>   Returns a string that represents the current object. </summary>
Public Function ToString() As String
    ToString = VBA.Information.TypeName(This.Target) & " -> " & VBA.Information.TypeName(This.Command)
End Function

' - - - - - - - - - - - - - - - -
'
' Command binding implementation
'
' - - - - - - - - - - - - - - - -

Private Property Get ICommandBinding_Target() As Object
    Set ICommandBinding_Target = This.Target
End Property

Private Property Get ICommandBinding_Command() As ICommand
    Set ICommandBinding_Command = This.Command
End Property

Private Sub ICommandBinding_EvaluateCanExecute(ByVal a_context As Object)
    EvaluateCanExecute a_context
End Sub

' - - - - - - - - - - - - - - - -
'
' Event handlers
'
' - - - - - - - - - - - - - - - -

Private Sub CheckBoxEvents_Click()

    Dim p_source As String
    p_source = ThisWorkbook.VBProject.Name & "." & VBA.Information.TypeName(Me) & ".CheckBoxEvents_Click"
    cc_isr_Core_IO.GuardClauses.GuardExpression Not TypeOf This.Target Is MSForms.CheckBox, _
            p_source, _
            "Invalid command binding type; must be a CheckBox", _
            cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError
            
    OnExecute

End Sub

Private Sub CommandButtonEvents_Click()

    Dim p_source As String
    p_source = ThisWorkbook.VBProject.Name & "." & VBA.Information.TypeName(Me) & ".CommandButtonEvents_Click"
    cc_isr_Core_IO.GuardClauses.GuardExpression Not TypeOf This.Target Is MSForms.CommandButton, p_source, _
            "Invalid command binding type; must be a CommandButton", _
            cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError
            
    OnExecute

End Sub

Private Sub ImageEvents_Click()

    Dim p_source As String
    p_source = ThisWorkbook.VBProject.Name & "." & VBA.Information.TypeName(Me) & ".ImageEvents_Click"
    cc_isr_Core_IO.GuardClauses.GuardExpression Not TypeOf This.Target Is MSForms.Image, p_source, _
            "Invalid command binding type; must be an Image", _
            cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError
            
    OnExecute
End Sub

Private Sub LabelEvents_Click()

    Dim p_source As String
    p_source = ThisWorkbook.VBProject.Name & "." & VBA.Information.TypeName(Me) & ".LabelEvents_Click"
    cc_isr_Core_IO.GuardClauses.GuardExpression Not TypeOf This.Target Is MSForms.Label, p_source, _
            "Invalid command binding type; must be a Label", _
            cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError
            
    OnExecute
End Sub

' - - - - - - - - - - - - - - - -
'
' Disposable implementation
'
' - - - - - - - - - - - - - - - -

Private Sub IDisposable_Dispose()
    Set This.Command = Nothing
    Set This.Target = Nothing
    Set This.ViewModel = Nothing
End Sub

