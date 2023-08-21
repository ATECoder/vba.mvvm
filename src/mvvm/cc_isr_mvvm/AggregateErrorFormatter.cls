VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AggregateErrorFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "A composite error formatter implementation."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   A composite error formatter implementation. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit
Implements IValidationErrorFormatter

Private Const AutoPropertyName As String = "AUTO"

Private Type TState
    Context As Object
    PropertyName As String
    Formatters As Collection
End Type

Private This As TState

''' <summary>   Initializes and returns an <see cref="IValidationErrorFormatter"/> representation of
'''             this <see cref="AggregateErrorFormatter"/> object that aggregates/combines
'''             all the specified formatters. </summary>
''' <param name="a_formatters">   [Variant] The formatters to aggregate. </param>
''' <returns>   [<see cref="IValidationErrorFormatter"/>]. </returns>
Public Function Initialize(ParamArray a_formatters() As Variant) As IValidationErrorFormatter

    Dim p_source As String
    p_source = ThisWorkbook.VBProject.Name & "." & VBA.Information.TypeName(Me) & ".Initialize"
    cc_isr_Core_IO.GuardClauses.GuardExpression UBound(a_formatters) < 0, p_source, _
            "At least one formatter is required.", _
            cc_isr_Core_IO.UserDefinedErrors.InvalidArgumentError
    
    ' handle ParamArray proxy args:
    
    Dim p_localFormatters() As Variant
    If VBA.IsArray(a_formatters(0)) Then
        p_localFormatters = a_formatters(0)
    Else
        ReDim p_localFormatters(UBound(a_formatters))
        Dim j As Long
        For j = 0 To UBound(p_localFormatters)
            p_localFormatters(j) = a_formatters(j)
        Next
    End If
    
    Dim i As Long
    For i = LBound(p_localFormatters) To UBound(p_localFormatters)
        
        If Not VBA.IsEmpty(p_localFormatters(i)) And Not p_localFormatters(i) Is Nothing Then
            Dim p_formatter As IValidationErrorFormatter
            On Error Resume Next
            Set p_formatter = p_localFormatters(i)
            On Error GoTo 0
            If Not p_formatter Is Nothing Then Me.AddFormatter p_formatter
        End If
        
    Next
    
    Set Initialize = Me
    
End Function

''' <summary>   Adds the specified formatter to this composite formatter. </summary>
Public Sub AddFormatter(ByVal a_formatter As IValidationErrorFormatter)
Attribute AddFormatter.VB_Description = "Adds the specified formatter to this composite formatter."
    
    Dim p_source As String
    p_source = ThisWorkbook.VBProject.Name & "." & VBA.Information.TypeName(Me) & ".AddFormatter"
    cc_isr_Core_IO.GuardClauses.GuardNullReference a_formatter, p_source
    
    This.Formatters.Add a_formatter

End Sub

Private Sub Class_Initialize()
    Set This.Formatters = New Collection
End Sub

Private Sub IValidationErrorFormatter_Apply(ByVal a_target As Object, ByVal a_message As String)
    Dim Formatter As IValidationErrorFormatter
    For Each Formatter In This.Formatters
        Formatter.Apply a_target, a_message
    Next
End Sub

Private Sub IValidationErrorFormatter_Restore(ByVal a_target As Object)
    Dim Formatter As IValidationErrorFormatter
    For Each Formatter In This.Formatters
        Formatter.Restore a_target
    Next
End Sub
