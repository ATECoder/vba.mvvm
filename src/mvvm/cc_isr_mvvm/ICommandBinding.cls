VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICommandBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An object responsible for binding a command to a UI element."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   An object responsible for binding a command to a UI element. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit

''' <summary>   Gets the event source object bound to the command. </summary>
Public Property Get Target() As Object
Attribute Target.VB_Description = "Gets the event source object bound to the command."
End Property

''' <summary>   Gets the command bound to the event source. </summary>
Public Property Get Command() As ICommand
Attribute Command.VB_Description = "Gets the command bound to the event source."
End Property

''' <summary>   Evaluates whether the command can execute given the binding context. </summary>
Public Sub EvaluateCanExecute(ByVal a_context As Object)
Attribute EvaluateCanExecute.VB_Description = "Evaluates whether the command can execute given the binding context."
End Sub
