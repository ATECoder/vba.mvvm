VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValidationErrors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "A collection of IValidationError objects for a given binding context (ViewModel)."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   A collection of IValidationError objects for a given binding context (ViewModel). </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit

Private Type TState
    BindingContext As Object
    ValidationErrors As Collection
End Type
Private This As TState

''' <summary>   Initializes a collection of <see cref="IValidationError"/> objects for the specified
'''             binding context (ViewModel) and returns this <see cref=ValidationErrors"/> object. </summary>
''' <returns>   [<see cref="ValidationErrors"/>]. </returns>
Public Function Initialize(ByVal a_bindingContext As Object) As ValidationErrors
Attribute Initialize.VB_Description = "Creates a new collection of IValidationError objects for the specified binding context (ViewModel)."
    
    Set Me.BindingContext = a_bindingContext
    
    Set Initialize = Me

End Function

''' <summary>   Gets/sets the binding context (ViewModel) for this instance. </summary>
Public Property Get BindingContext() As Object
Attribute BindingContext.VB_Description = "Gets/sets the binding context (ViewModel) for this instance."
    Set BindingContext = This.BindingContext
End Property

Public Property Set BindingContext(ByVal a_value As Object)
    Set This.BindingContext = a_value
End Property

''' <summary>   Gets the IValidationError for the specified property;
''' returns Nothing if there is no validation error. </summary>
Public Property Get Item(ByVal a_propertyName As String) As IValidationError
Attribute Item.VB_Description = "Gets the IValidationError for the specified property; returns Nothing if there is no validation error."
Attribute Item.VB_UserMemId = 0

    ' QUESTION: ' VF: Can this be set via Rubberduck?   Attribute Item.VB_UserMemId = 0
    ' I thought this were a trick I would be telling you as I think you did not mention it anywhere on your blog,  well you know them all ;-)
    ' actually, this trick needs some explanations for other to understand how come this works
    Dim p_result As IValidationError
    On Error Resume Next
    Set p_result = This.ValidationErrors.Item(a_propertyName)
    On Error GoTo 0 'key not found: returns Nothing
    Set Item = p_result
End Property

''' <summary>   Gets the number of properties with a validation error in this binding context. </summary>
Public Property Get Count() As Long
Attribute Count.VB_Description = "Gets the number of properties with a validation error in this binding context."
    Count = This.ValidationErrors.Count
End Property

''' <summary>   Provides custom collection support for For Each enumeration. </summary>
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_Description = "Provides custom collection support for For Each enumeration."
Attribute NewEnum.VB_UserMemId = -4

    ' QUESTION: ' VF: Can this be set via Rubberduck?
    ' Attribute NewEnum.VB_Description = "Provides custom collection support for For Each enumeration."
    ' Attribute NewEnum.VB_UserMemId = -4
    ' I thought this were a trick I would be telling you as I think you did not mention it anywhere on your blog,  well you know them all ;-)
    ' actually, this trick needs some explanations for other to understand how come this works
    ' regarding collections
    ' http://dailydoseofexcel.com/archives/2010/07/04/custom-collection-class/
    ' Attribute NewEnum.VB_MemberFlags = "40" 'for each/next in collection
    Set NewEnum = This.ValidationErrors.[_NewEnum]
End Property

''' <summary>   Adds (or replaces) a validation error. </summary>
Public Sub Add(ByVal a_validationError As IValidationError)
Attribute Add.VB_Description = "Adds (or replaces) a validation error."
    
    Dim p_propertyName As String
    p_propertyName = a_validationError.Binding.Source.PropertyName
    
    Remove p_propertyName
    This.ValidationErrors.Add a_validationError, p_propertyName
    
End Sub

''' <summary>   Removes validation error for the specified property if it exists. True if successful, False if nothing was done. </summary>
Public Function Remove(ByVal a_propertyName As String) As Boolean
Attribute Remove.VB_Description = "Removes validation error for the specified property if it exists. True if successful, False if nothing was done."
    Dim Existing As IValidationError
    Set Existing = Item(a_propertyName)
    If Not Existing Is Nothing Then
        This.ValidationErrors.Remove a_propertyName
        Remove = True
    End If
End Function

Private Sub Class_Initialize()
    Set This.ValidationErrors = New Collection
End Sub

Private Sub Class_Terminate()
    Set This.ValidationErrors = Nothing
End Sub
