VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICommandManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Represents an object responsible for managing command bindings in an application context."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Represents an object responsible for managing command bindings in an application context. </summary>
''' <remarks>
''' This manager knows about all the command bindings. This object is automatically created when you create an
''' <c>AppContext</c> instance, but the <c>AppContext</c> can be injected with any <c>ICommandManager</c>
''' implementation as needed.
''' </remarks>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit

''' <summary>   Binds the specified command to user interaction with the specified target UI element (e.g. a MSForms.CommandButton control). </summary>
Public Function BindCommand(ByVal a_bindingContext As Object, ByVal a_target As Object, ByVal a_command As ICommand) As ICommandBinding
Attribute BindCommand.VB_Description = "Binds the specified command to user interaction with the specified target UI element (e.g. a MSForms.CommandButton control)."
End Function

''' <summary>   Causes all commands to evaluate whether they can be executed. </summary>
Public Sub EvaluateCanExecute(ByVal a_bindingContext As Object)
Attribute EvaluateCanExecute.VB_Description = "Causes all commands to evaluate whether they can be executed."
End Sub
