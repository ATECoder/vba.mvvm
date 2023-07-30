VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IValidationManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Represents an object responsible for managing the validation errors in a ViewModel."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Represents an object responsible for managing the validation errors in a ViewModel. </summary>
''' <remarks>
''' The role of the validation manager is to encapsulate the validation state and expose methods to add and clear
''' validation errors; the IsValid indexed property returns a Boolean given a context (the ViewModel) and an
''' optional property name: in order to know whether the entire context is valid, omit the PropertyName argument.
''' </remarks>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit

''' <summary>   True if there are no validation errors, False otherwise. </summary>
''' <remarks>
''' he IsValid indexed property returns a Boolean given a context (the ViewModel) and an
''' optional property name: in order to know whether the entire context is valid, omit the PropertyName argument.
''' </remarks>
Public Property Get IsValid(ByVal a_context As Object, Optional ByVal a_propertyName As String) As Boolean
Attribute IsValid.VB_Description = "True if there are no validation errors, False otherwise."
End Property

''' <summary>   Gets the IValidationError for the specified binding path, if it exists. </summary>
''' <remarks>
''' An IValidationError is a simple object that gives us a Message (from the IValueValidator that
''' caused the binding to fail validation) and the IPropertyBinding that couldn’t be applied.
''' </remarks>
Public Property Get ValidationError(ByVal a_bindingPath As IBindingPath) As IValidationError
Attribute ValidationError.VB_Description = "Gets the IValidationError for the specified binding path, if it exists."
End Property

''' <summary>   Gets the factory used for creating the default validation error adorners. </summary>
''' <remarks>
''' </remarks>
Public Property Get AdornerFactory() As IDynamicAdornerFactory
Attribute AdornerFactory.VB_Description = "Gets the factory used for creating the default validation error adorners."
End Property

''' <summary>   Adds the specified validation error to the validation context. </summary>
''' <remarks>
''' OnValidationError and ClearValidationError respectively add and remove a validation error for a
''' particular source property, and the validation manager keeps validation errors in a dictionary
''' keyed with the ViewModel object (a string representation of its object pointer), such that each
''' ViewModel can be deemed “valid” or “invalid” individually/separately.
''' </remarks>
Public Sub OnValidationError(ByVal a_context As Object, ByVal a_validationError As IValidationError)
Attribute OnValidationError.VB_Description = "Adds the specified validation error to the validation context."
End Sub

''' <summary>   Removes the validation error for the specified binding path, if it exists. </summary>
''' <remarks>
''' OnValidationError and ClearValidationError respectively add and remove a validation error for a
''' particular source property, and the validation manager keeps validation errors in a dictionary
''' keyed with the ViewModel object (a string representation of its object pointer), such that each
''' ViewModel can be deemed “valid” or “invalid” individually/separately.
''' </remarks>
Public Sub ClearValidationError(ByVal a_bindingPath As IBindingPath)
Attribute ClearValidationError.VB_Description = "Removes the validation error for the specified binding path, if it exists."
End Sub
