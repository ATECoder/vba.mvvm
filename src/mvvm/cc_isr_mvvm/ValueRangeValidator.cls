VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValueRangeValidator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "A validator that accepts values between set min and max values. Uses Text comparison for strings."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   A validator that accepts values between set min and max values. Uses Text comparison for strings. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit
Option Compare Text
Implements IValueValidator

Private Type TState
    Min As Variant
    Max As Variant
End Type

Private This As TState

''' <summary>   Initializes and returns an <see cref="IValueValidator"/> representation of
'''             this <see cref="ValueRangeValidator"/> object. </summary>
''' <returns>   [<see cref="IValueValidator"/>]. </returns>
Public Function Initialize(ByVal a_min As Variant, ByVal a_max As Variant) As IValueValidator
    
    Me.Min = a_min
    Me.Max = a_max
    Set Initialize = Me

End Function

Public Property Get Min() As Variant
    Min = This.Min
End Property

Public Property Let Min(ByVal a_value As Variant)
    This.Min = a_value
End Property

Public Property Get Max() As Variant
    Max = This.Max
End Property

Public Property Let Max(ByVal a_value As Variant)
    This.Max = a_value
End Property

Private Function IValueValidator_IsValid(ByVal a_value As Variant, ByVal a_source As IBindingPath, _
        ByVal a_target As IBindingPath) As Boolean
    If IsObject(a_value) Or VBA.IsArray(a_value) Or VBA.IsEmpty(a_value) Or IsError(a_value) Then Exit Function
    IValueValidator_IsValid = a_value >= This.Min And a_value <= This.Max
End Function

Private Property Get IValueValidator_Message() As String

    IValueValidator_Message = cc_isr_Core.StringExtensions.StringFormat( _
            "A valid value is between {0} and {1}.", CStr(This.Min), CStr(This.Max))

End Property

Private Property Get IValueValidator_Trigger() As BindingUpdateSourceTrigger
    IValueValidator_Trigger = OnExit
End Property
