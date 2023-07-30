VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IValueConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An object that can convert a binding value to/from a type or value that is applicable to the binding target."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   An object that can convert a binding value to/from a type or value that is applicable to the binding target. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit

''' <summary>   Converts the source-provided value to a value that is applicable to the binding target. </summary>
Public Function Convert(ByVal a_value As Variant) As Variant
Attribute Convert.VB_Description = "Converts the source-provided value to a value that is applicable to the binding target."
End Function

''' <summary>   Converts the target-provided value to a value that is applicable to the binding source, for 2-way bindings. </summary>
Public Function ConvertBack(ByVal a_value As Variant) As Variant
Attribute ConvertBack.VB_Description = "Converts the target-provided value to a value that is applicable to the binding source, for 2-way bindings."
End Function
