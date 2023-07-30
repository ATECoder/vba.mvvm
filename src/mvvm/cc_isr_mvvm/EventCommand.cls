VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   A command that allows the ViewModel to supply the implementation.. </summary>
''' <remarks>
''' Not all commands are created equal: a command like CancelCommand is generic enough that it can
''' work with any ICancellable object, and an AcceptCommand can work with any implementation of the
''' IView interface. On the other hand, something feels wrong about systematically implementing any
''' & all commands in their own classes.
'''
''' Having each command neatly factored into its own class module is a great way to implement complex
''' commands, but can be overkill when things are relatively trivial – very often the ViewModel class
''' already has access to every object a command needs, and having a way to make the ViewModel itself
''' implement the command would solve this.
'''
''' I’m going to introduce an EventCommand class into the MVVM infrastructure code, to do exactly this:
'''
''' <code>
''' Private WithEvents PseudoDelegateCommand As EventCommand
'''
''' Private Sub Class_Initialize()
'''     Set PseudoDelegateCommand = EventCommand.Create("Full circle!")
''' End Sub
'''
''' Private Sub PseudoDelegateCommand_OnCanExecute(ByVal a_context As Object, a_outResult As Boolean)
'''     'supply the ICommand.CanExecute implementation here.
'''     'assign outResult to False to disable the command (it's True by default).
'''     'in principle, the Context *is* the ViewModel instance, so this assertion should hold:
'''     Debug.Assert Me Is Context
'''     'it also means the Context parameter should probably be ignored.
''' End Sub
'''
''' Private Sub PseudoDelegateCommand_OnExecute(ByVal a_context As Object)
'''     ' supply the ICommand.Execute implementation here.
'''     ' in principle, the Context *is* the ViewModel instance, so this assertion should hold:
'''     Debug.Assert Me Is Context
'''     ' it also means the Context parameter should probably be ignored.
'''     ' EventCommand is useful for commands that are specific to a particular ViewModel,
'''     ' and don't really need to have their implementation extracted into their own class.
''' End Sub
''' </code>
''' </remarks>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit
Implements ICommand
 
Private Type TState
    Description As String
End Type
 
Private This As TState
 
Public Event OnCanExecute(ByVal a_context As Object, ByRef a_outResult As Boolean)

Public Event OnExecute(ByVal a_context As Object)
 
''' <summary>   Gets/sets the command's Description. </summary>
Public Property Get Description() As String
    Description = This.Description
End Property
 
Public Property Let Description(ByVal a_value As String)
    This.Description = a_value
End Property
 
Private Function ICommand_CanExecute(ByVal a_context As Object) As Boolean
    
    Dim p_outResult As Boolean
    p_outResult = True
    RaiseEvent OnCanExecute(a_context, p_outResult)
    ICommand_CanExecute = p_outResult

End Function
 
Private Property Get ICommand_Description() As String
    ICommand_Description = This.Description
End Property
 
Private Sub ICommand_Execute(ByVal a_context As Object)
    RaiseEvent OnExecute(a_context)
End Sub

