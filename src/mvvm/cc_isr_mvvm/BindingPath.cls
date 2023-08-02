VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BindingPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An object that can resolve a string property path to a value."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   An object that can resolve a string property path to a value. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit
Implements IBindingPath

Private Type TState
    Context As Object
    Path As String
    
    Object As Object
    PropertyName As String
End Type

Private This As TState

''' <summary>   Initializes and returns an <see cref="IBindingPath"/> representation of
'''             this <see cref="BindingPath"/> object from the specified
'''             property path string and binding context. </summary>
''' <returns>   [<see cref="IBindingPath"/>]. </returns>
Public Function Initialize(ByVal a_context As Object, ByVal a_path As String) As IBindingPath
Attribute Initialize.VB_Description = "Initializes a new binding path from the specified property path string and binding context."
    
    cc_isr_Core.GuardClauses.GuardNullReference a_context, VBA.Information.TypeName(Me)
    cc_isr_Core.GuardClauses.GuardEmptyString a_path, VBA.Information.TypeName(Me)
    
    Set Me.Context = a_context
    Me.Path = a_path
    
    Me.Resolve
    
    Set Initialize = Me
    
End Function

''' <summary>   Gets/Sets the binding context. </summary>
Public Property Get Context() As Object
Attribute Context.VB_Description = "Gets/Sets the binding context."
    Set Context = This.Context
End Property

Public Property Set Context(ByVal a_value As Object)
    
    cc_isr_Core.GuardClauses.GuardNullReference a_value, VBA.Information.TypeName(Me)
    cc_isr_Core.GuardClauses.GuardDoubleInitialization This.Context, VBA.Information.TypeName(Me)
    Set This.Context = a_value

End Property

''' <summary>   Gets/Sets a string representing a property path against the binding context. </summary>
Public Property Get Path() As String
Attribute Path.VB_Description = "Gets/Sets a string representing a property path against the binding context."
    Path = This.Path
End Property

Public Property Let Path(ByVal a_value As String)
    
    cc_isr_Core.GuardClauses.GuardEmptyString a_value, VBA.Information.TypeName(Me)
    cc_isr_Core.GuardClauses.GuardDoubleInitialization This.Path, VBA.Information.TypeName(Me)
    This.Path = a_value

End Property

''' <summary>   Gets the bound object reference. </summary>
Public Property Get Object() As Object
Attribute Object.VB_Description = "Gets the bound object reference."
    Set Object = This.Object
End Property

''' <summary>   Gets the name of the bound property. </summary>
Public Property Get PropertyName() As String
Attribute PropertyName.VB_Description = "Gets the name of the bound property."
    PropertyName = This.PropertyName
End Property

''' <summary>   Resolves the Path to a bound object and property. </summary>
Public Sub Resolve()
Attribute Resolve.VB_Description = "Resolves the Path to a bound object and property."
    This.PropertyName = ResolvePropertyName(This.Path)
    Set This.Object = ResolvePropertyPath(This.Context, This.Path)
End Sub

Private Function ResolvePropertyName(ByVal a_propertyPath As String) As String
    Dim p_parts As Variant
    p_parts = Strings.Split(a_propertyPath, ".")
    ResolvePropertyName = p_parts(UBound(p_parts))
End Function

Private Function ResolvePropertyPath(ByVal a_context As Object, ByVal a_propertyPath As String) As Object
    
    Dim p_parts As Variant
    p_parts = Strings.Split(a_propertyPath, ".")
    
    If UBound(p_parts) = LBound(p_parts) Then
        Set ResolvePropertyPath = a_context
    Else
        Dim p_recursiveProperty As Object
        Set p_recursiveProperty = CallByName(a_context, p_parts(0), VbGet)
        If p_recursiveProperty Is Nothing Then Exit Function
        Set ResolvePropertyPath = ResolvePropertyPath(p_recursiveProperty, Right$(a_propertyPath, Len(a_propertyPath) - Len(p_parts(0)) - 1))
    End If
    
End Function

Private Property Get IBindingPath_Context() As Object
    Set IBindingPath_Context = This.Context
End Property

Private Property Get IBindingPath_Path() As String
    IBindingPath_Path = This.Path
End Property

Private Property Get IBindingPath_Object() As Object
    Set IBindingPath_Object = This.Object
End Property

Private Property Get IBindingPath_PropertyName() As String
    IBindingPath_PropertyName = This.PropertyName
End Property

Private Sub IBindingPath_Resolve()
    Set This.Object = ResolvePropertyPath(This.Context, This.Path)
End Sub

''' <summary>   Returns a string that represents the current object. </summary>
Private Function IBindingPath_ToString() As String
    IBindingPath_ToString = cc_isr_Core.StringExtensions.StringFormat("Context: {0}; Path: {1}", _
            VBA.Information.TypeName(This.Context), This.Path)
End Function

Private Function IBindingPath_TryReadPropertyValue(ByRef a_outValue As Variant) As Boolean
    If This.Object Is Nothing Then Resolve
    On Error Resume Next
    a_outValue = VBA.Interaction.CallByName(This.Object, This.PropertyName, VbGet)
    IBindingPath_TryReadPropertyValue = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Function IBindingPath_TryWritePropertyValue(ByVal a_value As Variant) As Boolean
    If This.Object Is Nothing Then Resolve
    On Error Resume Next
    VBA.Interaction.CallByName This.Object, This.PropertyName, VbLet, a_value
    IBindingPath_TryWritePropertyValue = (Err.Number = 0)
    On Error GoTo 0
End Function
