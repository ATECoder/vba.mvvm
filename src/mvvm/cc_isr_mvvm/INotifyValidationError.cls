VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "INotifyValidationError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An observer that can notify registered handlers when applying a binding throws a validation error."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   An observer that can notify registered handlers when applying a binding throws a validation error. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit

''' <summary>   Registers the specified handler. </summary>
Public Sub RegisterHandler(ByVal a_handler As IHandleValidationError)
Attribute RegisterHandler.VB_Description = "Registers the specified handler."
End Sub

''' <summary>   Notifies registered handlers. </summary>
Public Sub OnValidationErrorsChanged(ByVal a_bindingContext As Object, _
    Optional ByVal a_validationError As IValidationError)
End Sub
