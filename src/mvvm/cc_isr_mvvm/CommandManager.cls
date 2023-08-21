VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommandManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An object responsible for managing command bindings."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   An object responsible for managing command bindings. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit
Implements ICommandManager
Implements cc_isr_Core.IDisposable

Private Type TState
    CommandBindings As Collection
    DebugOutput As Boolean
End Type

Private This As TState

''' <summary>   Gets/sets a value indicating whether the object produces debugger output. </summary>
Public Property Get DebugOutput() As Boolean
Attribute DebugOutput.VB_Description = "Gets/sets a value indicating whether the object produces debugger output."
    DebugOutput = This.DebugOutput
End Property

Public Property Let DebugOutput(ByVal a_value As Boolean)
    This.DebugOutput = a_value
End Property

Private Sub EvaluateCanExecute(ByVal a_bindingContext As Object)
    Dim p_binding As ICommandBinding
    For Each p_binding In This.CommandBindings
        p_binding.EvaluateCanExecute a_bindingContext
    Next
End Sub

Private Sub DebugMessage(ByVal a_message As String)
    If This.DebugOutput Then Debug.Print VBA.Information.TypeName(Me) & ": " & a_message
End Sub

Private Sub Class_Initialize()
    Set This.CommandBindings = New Collection
End Sub

Private Function ICommandManager_BindCommand(ByVal a_bindingContext As Object, ByVal a_target As Object, ByVal a_command As ICommand) As ICommandBinding
    
    Dim p_source As String
    p_source = ThisWorkbook.VBProject.Name & "." & VBA.Information.TypeName(Me) & ".ICommandManager_BindCommand"
    cc_isr_Core_IO.GuardClauses.GuardNullReference a_bindingContext, p_source
    cc_isr_Core_IO.GuardClauses.GuardNullReference a_target, p_source
    cc_isr_Core_IO.GuardClauses.GuardNullReference a_command, p_source
    
    Dim p_binding As ICommandBinding
    
    Select Case True
    
        Case TypeOf a_target Is MSForms.CommandButton
            Set p_binding = Factory.NewCommandBinding().ForCommandButton(a_target, a_command, a_bindingContext)
            
        Case TypeOf a_target Is MSForms.Image
            Set p_binding = Factory.NewCommandBinding().ForImage(a_target, a_command, a_bindingContext)
            
        Case TypeOf a_target Is MSForms.Label
            Set p_binding = Factory.NewCommandBinding().ForLabel(a_target, a_command, a_bindingContext)
            
        Case TypeOf a_target Is MSForms.CheckBox
            Set p_binding = Factory.NewCommandBinding().ForCheckBox(a_target, a_command, a_bindingContext)
            
        Case Else
            cc_isr_Core_IO.GuardClauses.GuardExpression True, p_source, _
                "Target type '" & VBA.Information.TypeName(a_target) & _
                "' does not currently support command bindings.", _
                cc_isr_Core_IO.UserDefinedErrors.InvalidOperationError
            
    End Select
    
    This.CommandBindings.Add p_binding
    Set ICommandManager_BindCommand = p_binding
    
    DebugMessage "Created CommandBinding for " & VBA.Information.TypeName(a_command) & "."
    
End Function

Private Sub ICommandManager_EvaluateCanExecute(ByVal a_bindingContext As Object)
    EvaluateCanExecute a_bindingContext
End Sub

' - - - - - - - - - - - - - - - -
'
' Disposable implementation
'
' - - - - - - - - - - - - - - - -

Private Sub IDisposable_Dispose()
    Dispose
End Sub

''' <summary>   Releases all held ICommandBinding references, prepares the object for proper destruction. </summary>
Public Sub Dispose()
    '@Ignore VariableNotUsed
    Dim Index As Long
    For Index = 1 To This.CommandBindings.Count
        cc_isr_Core.DisposableExtensions.TryDispose This.CommandBindings.Item(This.CommandBindings.Count)
        This.CommandBindings.Remove This.CommandBindings.Count
    Next
    Set This.CommandBindings = Nothing
End Sub

