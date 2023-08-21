VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CancelCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "A command that closes (hides) a cancellable View in a canceled state."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   A command that closes (hides) a cancellable View in a canceled state. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit
Implements ICommand

Private Type TState
    View As ICancellable
End Type

Private This As TState

''' <summary>   Initializes and returns am <see cref="ICommand"/> representation of
'''             this <see cref="CancelCommand"/> object. </summary>
''' <returns>   [<see cref="ICommand"/>]. </returns>
Public Function Initialize(ByVal a_view As ICancellable) As ICommand
Attribute Initialize.VB_Description = "Initializes a new instance of this command."
    
    Set Me.View = a_view
    
    Set Initialize = Me
    
End Function

Public Property Get View() As ICancellable
    Set View = This.View
End Property

Public Property Set View(ByVal a_value As ICancellable)
    
    Dim p_source As String
    p_source = ThisWorkbook.VBProject.Name & "." & VBA.Information.TypeName(Me) & ".View"
    cc_isr_Core_IO.GuardClauses.GuardDoubleInitialization This.View, p_source
    Set This.View = a_value

End Property

Private Function ICommand_CanExecute(ByVal a_context As Object) As Boolean
    ICommand_CanExecute = True
End Function

Private Property Get ICommand_Description() As String
    ICommand_Description = "Cancel pending changes and close."
End Property

Private Sub ICommand_Execute(ByVal a_context As Object)
    This.View.OnCancel
End Sub
