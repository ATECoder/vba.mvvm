VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringToDateConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "A value converter that can convert a String into a Date, or a Date into a String."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   A value converter that can convert a String into a Date, or a Date into a String. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit
Implements IValueConverter

Public Function Default() As IValueConverter
   
    Set Default = StringToDateConverter

End Function

Private Function IValueConverter_Convert(ByVal a_value As Variant) As Variant
    
    IValueConverter_Convert = Convert(a_value)

End Function

Private Function IValueConverter_ConvertBack(ByVal a_value As Variant) As Variant
    
    IValueConverter_ConvertBack = Convert(a_value)

End Function

Private Function Convert(ByVal a_value As Variant) As Variant
    
    cc_isr_Core.GuardClauses.GuardExpression Not VBA.IsDate(a_value), VBA.Information.TypeName(Me), _
            "Value '" & a_value & "' was not recognized as a valid date.", _
            cc_isr_Core.UserDefinedErrors.InvalidArgumentError
            
    If VBA.VarType(a_value) = VBA.VbVarType.vbDate Then
        Convert = VBA.CDate(a_value)
    Else
        Convert = VBA.CStr(a_value)
    End If
End Function
