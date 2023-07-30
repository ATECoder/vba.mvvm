VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IHandleValidationError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An object that can be registered as a listener for data validation errors."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   An object that can be registered as a listener for data validation errors. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit

''' <summary>   A callback invoked when a validation error is added or removed from the validation manager. </summary>
Public Sub HandleValidationErrorsChanged(ByVal a_bindingPath As IBindingPath, Optional ByVal a_validationError As IValidationError)
Attribute HandleValidationErrorsChanged.VB_Description = "A callback invoked when a validation error is added or removed from the validation manager."
End Sub
