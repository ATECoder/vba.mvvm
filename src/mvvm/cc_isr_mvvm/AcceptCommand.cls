VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AcceptCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "A command that closes (hides) a View, enabled when the ViewModel is valid."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   A command that closes (hides) a View, enabled when the ViewModel is valid. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit
Implements ICommand

Private Type TState
    View As IView
    ValidationManager As cc_isr_MVVM.IValidationManager
End Type

Private This As TState

''' <summary>   Initializes this <see cref=AcceptCommand"/> object. </summary>
Public Sub Initialize(ByVal a_view As cc_isr_MVVM.IView, ByVal a_validationManager As cc_isr_MVVM.IValidationManager)

    cc_isr_Core.GuardClauses.GuardNullReference a_view, VBA.Information.TypeName(Me), "View cannot be Nothing."
    cc_isr_Core.GuardClauses.GuardNullReference a_validationManager, VBA.Information.TypeName(Me), "ValidationManager cannot be Nothing."
    Set Me.View = a_view
    Set Me.ValidationManager = a_validationManager

End Sub


Public Property Get View() As cc_isr_MVVM.IView
    Set View = This.View
End Property

Public Property Set View(ByVal a_value As cc_isr_MVVM.IView)
    
    cc_isr_Core.GuardClauses.GuardDoubleInitialization This.View, VBA.Information.TypeName(Me)
    Set This.View = a_value

End Property

Public Property Get ValidationManager() As IValidationManager
    
    Set ValidationManager = This.ValidationManager

End Property

Public Property Set ValidationManager(ByVal a_value As IValidationManager)
    
    cc_isr_Core.GuardClauses.GuardDoubleInitialization This.ValidationManager, VBA.Information.TypeName(Me)
    Set This.ValidationManager = a_value

End Property

Private Function ICommand_CanExecute(ByVal a_context As Object) As Boolean
    
    ICommand_CanExecute = This.ValidationManager.IsValid(a_context)

End Function

Private Property Get ICommand_Description() As String
    
    ICommand_Description = "Accept changes and close."

End Property

Private Sub ICommand_Execute(ByVal a_context As Object)
    
    This.View.Hide

End Sub
