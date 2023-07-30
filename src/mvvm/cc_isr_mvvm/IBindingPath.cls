VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IBindingPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An object that can resolve a string property path to a value."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   An object that can resolve a string property path to a value. </summary>
''' <remarks>
'''
''' <c>IBindingPath</c> is implemented by the <see cref="BindingPath"/> object by an initialize
''' method that property-injects the <c>Context</c> and <c>Path</c> values and invokes the
''' <c>Resolve</c> method. Thus, creating a binding path really just looks like this:
''' <code>
''' Dim Path As IBindingPath
''' Set Path = Factory.CreateBindingPath(ViewModel, "PropertyName")
''' <code>
'''
''' </remarks>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit

''' <summary>   Gets the context for resolving the Path. </summary>
''' <remarks>
''' Context refers to the base object for resolving the path, i.e., your ViewModel
''' (for the source), or a UI control (for the target).
''' </remarks>
Public Property Get Context() As Object
Attribute Context.VB_Description = "Gets the context for resolving the Path."
End Property

''' <summary>   Gets the property path of the binding. </summary>
''' <remarks>
''' Path refers to the property path string; usually just a ViewModel or control property name,
''' but this string is resolved recursively so you could bind to “ViewModel.SomeObject.SomeProperty”
''' if you needed to.
''' </remarks>
Public Property Get Path() As String
Attribute Path.VB_Description = "Gets the property path of the binding."
End Property

''' <summary>   Gets a reference to the binding source object resolved from the Path. </summary>
''' <remarks>
''' Object is the object that owns the property we’re referring to. If the path is just a property name,
''' then this is the same reference as the Context.
''' </remarks>
Public Property Get Object() As Object
Attribute Object.VB_Description = "Gets a reference to the binding source object resolved from the Path."
End Property

''' <summary>   Gets the name of the bound property resolved from the Path. </summary>
''' <remarks>
''' PropertyName is the resolved property name. In the example path above, of
''' <c>ViewModel.SomeObject.SomeProperty</c> that would be <c>SomeProperty</c>.
''' </remarks>
Public Property Get PropertyName() As String
Attribute PropertyName.VB_Description = "Gets the name of the bound property resolved from the Path."
End Property

''' <summary>   Resolves the binding source. </summary>
''' <remarks>
''' Invoked by the MVVM infrastructure internals.
''' </remarks>
Public Sub Resolve()
Attribute Resolve.VB_Description = "Resolves the binding source."
End Sub

''' <summary>   Reads and returns the property value; False on error, True if successful. </summary>
''' <remarks>
''' Invoked by the MVVM infrastructure internals.
''' </remarks>
Public Function TryReadPropertyValue(ByRef a_outValue As Variant) As Boolean
Attribute TryReadPropertyValue.VB_Description = "Reads and returns the property value; False on error, True if successful."
End Function

''' <summary>   Writes the value to the bound property; False on error, True if successful. </summary>
''' <remarks>
''' Invoked by the MVVM infrastructure internals.
''' </remarks>
Public Function TryWritePropertyValue(ByVal a_value As Variant) As Boolean
Attribute TryWritePropertyValue.VB_Description = "Writes the value to the bound property; False on error, True if successful."
End Function

''' <summary>   Returns a string representing this binding path for debugging purposes. </summary>
''' <remarks>
''' Invoked by the MVVM infrastructure internals.
''' </remarks>
Public Function ToString() As String
Attribute ToString.VB_Description = "Returns a string representing this binding path for debugging purposes."
End Function
