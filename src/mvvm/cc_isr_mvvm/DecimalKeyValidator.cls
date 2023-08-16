VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DecimalKeyValidator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "A key validator that allows numeric and decimal separator inputs."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   A key validator that allows numeric and decimal separator inputs. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit
Implements IValueValidator

Private Type ThisData
    SeparatorChar As String
    UpdateSourceTriggerOption As cc_isr_MVVM.BindingUpdateSourceTrigger 'avoid being called twice
End Type

Private This As ThisData

Private Sub Class_Initialize()
    This.SeparatorChar = VBA.Strings.Format$(0, ".")
End Sub

Private Function IValueValidator_IsValid(ByVal a_value As Variant, ByVal a_source As IBindingPath, ByVal a_target As IBindingPath) As Boolean
    Dim p_result As Boolean
    p_result = IsNumeric(a_value)
    
    If a_value = This.SeparatorChar Then
        
        Dim p_sourceValue As String
        If a_source.TryReadPropertyValue(a_outValue:=p_sourceValue) Then
            'decimal separator character is legal if there aren't any already
            Dim TargetValue As String
            If a_target.TryReadPropertyValue(a_outValue:=TargetValue) Then
                If p_sourceValue = 0 Or IsBindingTargetEmpty(a_target) Then
                    TargetValue = VBA.vbNullString
                End If
                Dim Separators As Long
                Separators = Len(TargetValue) - Len(Replace$(TargetValue, This.SeparatorChar, VBA.vbNullString))
                p_result = Separators = 0
            End If
        End If
    End If
    
    IValueValidator_IsValid = p_result
End Function

Private Function IsBindingTargetEmpty(ByVal a_target As IBindingPath) As Boolean
    'TargetValue may be string-formatted; if all the content is selected, treat it as empty.
    On Error Resume Next
    'late-bound member calls should work against TextBox-like controls
    IsBindingTargetEmpty = a_target.Context.SelLength = a_target.Context.TextLength
    On Error GoTo 0
End Function

Private Property Get IValueValidator_Message() As String

    ' IValueValidator_Message = "Value must be numeric."
    ' VF: reworded to reflect different logic/trigger
    IValueValidator_Message = "Numeric value required."

End Property

Private Property Get IValueValidator_Trigger() As BindingUpdateSourceTrigger
    ' VF: funnel through OnChange to capture cut/paste and backspace/delete
    'avoid being called twice
    If This.UpdateSourceTriggerOption = BindingUpdateSourceTrigger.NotSetYet Then
        Select Case MsgBox("Trigger validation of numeric text boxes 'OnChange'?" & vbCr & vbCr & "No = Trigger 'OnKeyPress' (as originally, which means not capturing cut/paste and backspace/delete in the text boxes)", vbQuestion + vbYesNo, VBA.Information.TypeName(Me))
            Case vbYes
                'but deletes entire string ...as of now
                'IValueValidator_Trigger = OnChange
                This.UpdateSourceTriggerOption = BindingUpdateSourceTrigger.OnChange
            Case Else
                'IValueValidator_Trigger = OnKeyPress
                This.UpdateSourceTriggerOption = BindingUpdateSourceTrigger.OnKeyPress
        End Select
    Else
        IValueValidator_Trigger = This.UpdateSourceTriggerOption
    End If
End Property
