VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An object representing a format string in a notation compatible for use with .NET String.Format()."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>  An object representing a format string in a notation compatible for use with
'''            cc.isr.Core.StringExtension.FormatString syntax, which is compatibvle with the
'''            .NET notation absent in braces formatting. </summary>
''' <remarks> Dependencies:
''' cc_isr_Core.StringExtensions.cls
''' cc_isr_MVVM.IStringFormatter.cls </remarks>
''' - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Implements IStringFormatter

Private Type TState
    FormatString As String
End Type

Private This As TState

''' <summary>   Initializes and returns an <see cref="IStringFormatter"/> represention of
'''             this <see cref="StringFormatter"/> object. </summary>
''' <returns>   [<see cref="IStringFormatter"/>]. </returns>
Public Function Initialize(ByVal a_formatString As String) As IStringFormatter
    
    Me.FormatString = a_formatString
    
    Set Initialize = Me

End Function

Public Property Get FormatString() As String
    FormatString = This.FormatString
End Property

Public Property Let FormatString(ByVal a_value As String)
    
    This.FormatString = a_value

End Property

Private Function IStringFormatter_Format(ParamArray a_values() As Variant) As String
    
    IStringFormatter_Format = cc_isr_Core.StringExtensions.StringFormat(This.FormatString, a_values)
    
End Function

Private Property Get IStringFormatter_FormatString() As String
    IStringFormatter_FormatString = This.FormatString
End Property
