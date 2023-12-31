VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValidationManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Holds the validation errors state for a ViewModel."
'@Folder rd_MVVM.Infrastructure.Validation
'@ModuleDescription "Holds the validation errors state for a ViewModel."
'@PredeclaredId
'@Exposed
Option Explicit
Implements IValidationManager

Private Type TState
    ErrorAdornerFactory As IDynamicAdornerFactory
    ValidationErrors As Dictionary
    IsValid As Boolean
End Type

Private This As TState

Public Function Create(Optional ByVal ErrorAdornerFactory As IDynamicAdornerFactory) As ValidationManager
    Dim result As ValidationManager
    Set result = New ValidationManager
    InitAdornerFactory result, ErrorAdornerFactory
    Set Create = result
End Function

Private Sub InitAdornerFactory(ByVal Manager As rd_MVVM.ValidationManager, ByVal AdornerFactory As rd_MVVM.IDynamicAdornerFactory)
    Dim Factory As rd_MVVM.IDynamicAdornerFactory
    
    If AdornerFactory Is Nothing Then
        Dim DefaultFactory As rd_MVVM.IDynamicAdornerFactory
        Set DefaultFactory = New DefaultErrorAdornerFactory
        Set Factory = DefaultFactory
    Else
        'use injected instance (could be a test stub):
        Set Factory = AdornerFactory
        
    End If
    
    Set Manager.ErrorAdornerFactory = Factory
End Sub

Public Property Get ErrorAdornerFactory() As IDynamicAdornerFactory
    Set ErrorAdornerFactory = This.ErrorAdornerFactory
End Property

Friend Property Set ErrorAdornerFactory(ByVal RHS As IDynamicAdornerFactory)
    GuardClauses.GuardDefaultInstance Me, ValidationManager
    GuardClauses.GuardNullReference RHS, TypeName(ValidationManager)
    GuardClauses.GuardDoubleInitialization This.ErrorAdornerFactory, TypeName(Me)
    Set This.ErrorAdornerFactory = RHS
End Property

Public Property Get IsValid(ByVal BindingContext As Object, Optional ByVal PropertyName As String) As Boolean
    GuardClauses.GuardDefaultInstance Me, ValidationManager
    
    Dim result As Boolean
    result = True
    
    Dim Key As Variant
    For Each Key In This.ValidationErrors.Keys
        
        Dim Errors As ValidationErrors
        Set Errors = This.ValidationErrors.Item(Key)
        
        If BindingContext Is Errors.BindingContext Or Errors.BindingContext Is Nothing Then
            
            If PropertyName = vbNullString Then
                result = Errors.Count = 0
            Else
                Dim Error As IValidationError
                For Each Error In Errors
                    result = (Error.Binding.Source.PropertyName <> PropertyName)
                    If Not result Then Exit For
                Next
            End If
            
        End If
        
        If Not result Then Exit For
    Next
    
    IsValid = result
    
End Property

Public Property Get ValidationError(ByVal BindingPath As IBindingPath) As IValidationError
    Dim ContextKey As String
    ContextKey = CStr(ObjPtr(BindingPath.Context))
    If Not This.ValidationErrors.Exists(ContextKey) Then Exit Property
    
    Dim Errors As ValidationErrors
    Set Errors = This.ValidationErrors(ContextKey)
    
    'ValidationErrors.Item getter returns Nothing if there's no error for the given property:
    Set ValidationError = Errors.Item(BindingPath.PropertyName)
End Property

Private Sub ClearValidationErrors(ByVal BindingPath As IBindingPath)
    Dim ContextKey As String
    ContextKey = CStr(ObjPtr(BindingPath.Context))
    If Not This.ValidationErrors.Exists(ContextKey) Then Exit Sub
    
    Dim Errors As ValidationErrors
    Set Errors = This.ValidationErrors(ContextKey)
    Errors.Remove BindingPath.PropertyName
End Sub

Private Sub OnValidationError(ByVal Context As Object, ByVal ValidationError As IValidationError)
    Dim ContextKey As String
    ContextKey = CStr(ObjPtr(ValidationError.Binding.Source.Context))
    
    Dim Errors As ValidationErrors
    If This.ValidationErrors.Exists(ContextKey) Then
        Set Errors = This.ValidationErrors.Item(ContextKey)
    Else
        Set Errors = ValidationErrors.Create(Context)
    End If
    
    Errors.Add ValidationError
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

Private Sub IValidationManager_ClearValidationError(ByVal BindingPath As IBindingPath)
    ClearValidationErrors BindingPath
End Sub

Private Property Get IValidationManager_IsValid(ByVal Context As Object, Optional ByVal PropertyName As String) As Boolean
    IValidationManager_IsValid = IsValid(Context, PropertyName)
End Property

Private Sub IValidationManager_OnValidationError(ByVal Context As Object, ByVal ValidationError As IValidationError)
    OnValidationError Context, ValidationError
End Sub

Private Property Get IValidationManager_ValidationError(ByVal BindingPath As IBindingPath) As IValidationError
    Set IValidationManager_ValidationError = ValidationError(BindingPath)
End Property
