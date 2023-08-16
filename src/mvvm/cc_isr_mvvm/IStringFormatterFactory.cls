VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IStringFormatterFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "A factory that creates instances of a particular IStringFormatter implementation."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Represents a factory that creates instances of a particular IStringFormatter implementation.
''' <remarks> Dependencies:
''' cc_isr_MVVM.IStringFormatter.cls </remarks>
''' - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   Creates a new instance of a string formatter. </summary>
Public Function Create(ByVal a_formatString As String) As IStringFormatter
Attribute Create.VB_Description = "Creates a new instance of a string formatter."
End Function
