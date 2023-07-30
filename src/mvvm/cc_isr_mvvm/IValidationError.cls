VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IValidationError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An object representing a binding validation error."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   An object representing a binding validation error. </summary>
''' <remarks>
''' An IValidationError is a simple object that gives us a Message (from the IValueValidator that
''' caused the binding to fail validation) and the IPropertyBinding that couldn’t be applied.
''' </remarks>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit

''' <summary>   The validation error message. </summary>
Public Property Get Message() As String
Attribute Message.VB_Description = "The validation error message."
End Property

''' <summary>   The binding that failed to validate. </summary>
Public Property Get Binding() As IPropertyBinding
Attribute Binding.VB_Description = "The binding that failed to validate."
End Property
