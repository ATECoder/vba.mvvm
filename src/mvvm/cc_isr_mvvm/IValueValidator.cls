VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IValueValidator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An object responsible for validating the value before a binding is applied."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   An object responsible for validating the value before a binding is applied. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit

''' <summary>   Using this validator in a property binding will coerce it into using this update trigger. </summary>
Public Property Get Trigger() As BindingUpdateSourceTrigger
Attribute Trigger.VB_Description = "Using this validator in a property binding will coerce it into using this update trigger."
End Property

''' <summary>   A user-friendly message describing the validation rule(s). </summary>
Public Property Get Message() As String
Attribute Message.VB_Description = "A user-friendly message describing the validation rule(s)."
End Property

''' <summary>   True if the specified value is valid in itself or in the context of the binding source and target. </summary>
Public Function IsValid(ByVal a_value As Variant, ByVal a_source As IBindingPath, _
    ByVal a_target As IBindingPath) As Boolean
End Function
