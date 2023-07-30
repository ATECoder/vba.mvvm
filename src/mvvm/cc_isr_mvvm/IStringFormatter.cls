VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IStringFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Represents an object with the ability to format a string value."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Represents an object with the ability to format a string value. </summary>
''' - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   Gets the format string used when formatting values. </summary>
''' <value>   [String]. </value>
Public Property Get FormatString() As String
Attribute FormatString.VB_Description = "Gets the format string used when formatting values."
End Property

''' <summary>   Replaces braced numerical values defined in the <see cref="FormatString">
'''             with <paramref="a_values"/> similar to VB.NET String.Format but without
'''             formatting suffixes. </summary>
''' <param name="a_values">   [Variant, ParamArray] The parameter values matching the
'''                           <see cref="FromatString"/>. </param>
''' <returns>   [String]. </returns>
Public Function Format(ParamArray a_values() As Variant) As String
Attribute Format.VB_Description = "Formats the specified values as per the FormatString."
End Function
