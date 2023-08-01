VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValidationManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Holds the validation errors state for a ViewModel."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Holds the validation errors state for a ViewModel. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit
Implements IValidationManager

Private Type TState
    ErrorAdornerFactory As IDynamicAdornerFactory
    ValidationErrors As Dictionary
    IsValid As Boolean
End Type

Private This As TState

''' <summary>   Initializes and returns this <see cref=ValidationManager"/> object. </summary>
Public Function Initialize(Optional ByVal a_errorAdornerFactory As IDynamicAdornerFactory) As ValidationManager
    
    Me.InitAdornerFactory a_errorAdornerFactory
    
    Set Initialize = Me

End Function

Friend Sub InitAdornerFactory(ByVal a_adornerFactory As cc_isr_MVVM.IDynamicAdornerFactory)
    
    Dim p_factory As cc_isr_MVVM.IDynamicAdornerFactory
    
    If a_adornerFactory Is Nothing Then
        Dim p_defaultFactory As cc_isr_MVVM.IDynamicAdornerFactory
        Set p_defaultFactory = New DefaultErrorAdornerFactory
        Set p_factory = p_defaultFactory
    Else
        'use injected instance (could be a test stub):
        Set p_factory = a_adornerFactory
        
    End If
    
    Set Me.ErrorAdornerFactory = p_factory
    
End Sub

Public Property Get ErrorAdornerFactory() As IDynamicAdornerFactory
    Set ErrorAdornerFactory = This.ErrorAdornerFactory
End Property

Public Property Set ErrorAdornerFactory(ByVal a_value As IDynamicAdornerFactory)
    
    cc_isr_Core.GuardClauses.GuardNullReference a_value, VBA.Information.TypeName(Me)
    cc_isr_Core.GuardClauses.GuardDoubleInitialization This.ErrorAdornerFactory, VBA.Information.TypeName(Me)
    Set This.ErrorAdornerFactory = a_value

End Property

Public Property Get IsValid(ByVal a_bindingContext As Object, Optional ByVal a_propertyName As String) As Boolean
    
    Dim p_result As Boolean
    p_result = True
    
    Dim Key As Variant
    For Each Key In This.ValidationErrors.Keys
        
        Dim Errors As ValidationErrors
        Set Errors = This.ValidationErrors.Item(Key)
        
        If a_bindingContext Is Errors.BindingContext Or Errors.BindingContext Is Nothing Then
            
            If a_propertyName = VBA.vbNullString Then
                p_result = Errors.Count = 0
            Else
                Dim Error As IValidationError
                For Each Error In Errors
                    p_result = (Error.Binding.Source.PropertyName <> a_propertyName)
                    If Not p_result Then Exit For
                Next
            End If
            
        End If
        
        If Not p_result Then Exit For
    Next
    
    IsValid = p_result
    
End Property

Public Property Get ValidationError(ByVal a_bindingPath As IBindingPath) As IValidationError
    Dim ContextKey As String
    ContextKey = CStr(ObjPtr(a_bindingPath.Context))
    If Not This.ValidationErrors.Exists(ContextKey) Then Exit Property
    
    Dim Errors As ValidationErrors
    Set Errors = This.ValidationErrors(ContextKey)
    
    'ValidationErrors.Item getter returns Nothing if there's no error for the given property:
    Set ValidationError = Errors.Item(a_bindingPath.PropertyName)
End Property

Private Sub ClearValidationErrors(ByVal a_bindingPath As IBindingPath)
    Dim ContextKey As String
    ContextKey = CStr(ObjPtr(a_bindingPath.Context))
    If Not This.ValidationErrors.Exists(ContextKey) Then Exit Sub
    
    Dim Errors As ValidationErrors
    Set Errors = This.ValidationErrors(ContextKey)
    Errors.Remove a_bindingPath.PropertyName
End Sub

Private Sub OnValidationError(ByVal a_context As Object, ByVal a_validationError As IValidationError)
    Dim ContextKey As String
    ContextKey = CStr(ObjPtr(a_validationError.Binding.Source.Context))
    
    Dim Errors As ValidationErrors
    If This.ValidationErrors.Exists(ContextKey) Then
        Set Errors = This.ValidationErrors.Item(ContextKey)
    Else
        Set Errors = Factory.CreateValidationErrors(a_context)
    End If
    
    Errors.Add a_validationError
    Set This.ValidationErrors.Item(ContextKey) = Errors
End Sub

Private Sub Class_Initialize()
    Set This.ValidationErrors = New Dictionary
    This.IsValid = True
End Sub

Private Sub Class_Terminate()
    Set This.ValidationErrors = Nothing
End Sub

Private Property Get IValidationManager_AdornerFactory() As IDynamicAdornerFactory
    Set IValidationManager_AdornerFactory = This.ErrorAdornerFactory
End Property

Private Sub IValidationManager_ClearValidationError(ByVal a_bindingPath As IBindingPath)
    ClearValidationErrors a_bindingPath
End Sub

Private Property Get IValidationManager_IsValid(ByVal a_context As Object, Optional ByVal a_propertyName As String) As Boolean
    IValidationManager_IsValid = IsValid(a_context, a_propertyName)
End Property

Private Sub IValidationManager_OnValidationError(ByVal a_context As Object, ByVal a_validationError As IValidationError)
    OnValidationError a_context, a_validationError
End Sub

Private Property Get IValidationManager_ValidationError(ByVal a_bindingPath As IBindingPath) As IValidationError
    Set IValidationManager_ValidationError = ValidationError(a_bindingPath)
End Property
