VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValidationError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Represents a binding validation error."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Represents a binding validation error. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit

Implements IValidationError

Private Type TState
    Binding As IPropertyBinding
    Message As String
End Type

Private This As TState

''' <summary>   Initializes this <see cref="ValidationError"/> object. </summary>
Public Sub Initialize(ByVal a_binding As IPropertyBinding, ByVal a_message As String)
Attribute Initialize.VB_Description = "Creates a new instance of a validation error object."
    
    Set Me.Binding = a_binding
    Me.Message = a_message

End Sub

''' <summary>   The validation error message. </summary>
Public Property Get Message() As String
Attribute Message.VB_Description = "The validation error message."
    Message = This.Message
End Property

Public Property Let Message(ByVal a_value As String)
    This.Message = a_value
End Property

''' <summary>   Gets the binding responsible for this validation error. </summary>
Public Property Get Binding() As IPropertyBinding
Attribute Binding.VB_Description = "Gets the binding responsible for this validation error."
    Set Binding = This.Binding
End Property

Public Property Set Binding(ByVal a_value As IPropertyBinding)
    Set This.Binding = a_value
End Property

Private Property Get IValidationError_Message() As String
    IValidationError_Message = This.Message
End Property

Private Property Get IValidationError_Binding() As IPropertyBinding
    Set IValidationError_Binding = This.Binding
End Property
