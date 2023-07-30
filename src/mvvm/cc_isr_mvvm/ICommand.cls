VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An object that represents an executable command."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   An object that represents an executable command. </summary>
''' <remarks>
''' Using the command pattern (even without MVVM command bindings), a control event handler, such as Button_Click,
''' would still be responsible for kicking the “create a new item” logic into action… but now that logic would be
''' living in some ICommand implementation, encapsulating its dependencies and state (and thus moving these outside
''' of the form’s code-behind but not into global scope now).
''' This makes a command as an abstraction that has:
''' - A user-friendly description of what the command does.
''' - A function that takes a context object and returns a Boolean value that indicates whether the command can
'''   currently be executed.
''' - An Execute procedure that takes a context object and, well, executes the command.
''' </remarks>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit

''' <summary>   Returns True if the command is enabled given the provided binding context (ViewModel). </summary>
''' <param name="a_context">   [Object] encapsulates the state, the data we’re working with. In
'''                            MVVM that would be the ViewModel instance.</param>
Public Function CanExecute(ByVal a_context As Object) As Boolean
Attribute CanExecute.VB_Description = "Returns True if the command is enabled given the provided binding context (ViewModel)."
End Function

''' <summary>   Executes the command given the provided binding context (ViewModel). </summary>
''' <remarks>
''' </remarks>
Public Sub Execute(ByVal a_context As Object)
Attribute Execute.VB_Description = "Executes the command given the provided binding context (ViewModel)."
End Sub

''' <summary>   Gets a user-friendly description of the command. </summary>
''' <remarks>
''' MVVM command bindings use the Description property to set the ControlToolTip string of a
''' binding’s target CommandButton object, and automatically invokes the CanExecute method
''' as property bindings update, which automatically enables or disables the bound command
''' button control: the command pattern works very, very well with Model-View-ViewModel, but
''' nothing says we cannot use the command pattern without it.
''' </remarks>
Public Property Get Description() As String
Attribute Description.VB_Description = "Gets a user-friendly description of the command."
End Property
