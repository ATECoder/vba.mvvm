VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InverseBooleanConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "A value converter that inverts a Boolean value."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   A value converter that inverts a Boolean value. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit
Implements IValueConverter

Private Function IValueConverter_Convert(ByVal a_value As Variant) As Variant

    IValueConverter_Convert = Not CBool(a_value)

End Function

Private Function IValueConverter_ConvertBack(ByVal a_value As Variant) As Variant

    IValueConverter_ConvertBack = Not CBool(a_value)

End Function
