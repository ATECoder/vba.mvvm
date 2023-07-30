VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringFormatterFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "A factory that creates new .NET-syntax string formatters."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   A factory that creates new .NET-syntax string formatters.
''' <remarks> Dependencies:
''' cc_isr_MVVM.IStringFormatterFactory.cls, cc_isr_MVVM.IStringFormatter.cls,
''' cc_isr_MVVM.StringFormatter.cls.
''' </remarks>
''' - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Implements IStringFormatterFactory

Private Function IStringFormatterFactory_Create(ByVal a_formatString As String) As IStringFormatter
    Set IStringFormatterFactory_Create = Factory.CreateStringFormatter(a_formatString)
End Function
