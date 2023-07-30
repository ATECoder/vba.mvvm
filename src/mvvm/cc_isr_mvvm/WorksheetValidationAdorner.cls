VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorksheetValidationAdorner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An object that dynamically decorates a target Excel.Range object."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   An object that dynamically decorates a target Excel.Range object. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit
Implements IDynamicAdorner

Private Type TState
    Target As Excel.Range
    TargetFormatter As cc_isr_MVVM.IValidationErrorFormatter
End Type

Private This As TState

Public Function Create(ByVal a_target As Excel.Range, ByVal Formatter As cc_isr_MVVM.IValidationErrorFormatter) As IDynamicAdorner
    Dim p_result As WorksheetValidationAdorner
    Set p_result = New WorksheetValidationAdorner
    Set p_result.Target = a_target
    Set p_result.Formatter = Formatter
    Set Create = p_result
End Function

Public Property Get Formatter() As cc_isr_MVVM.IValidationErrorFormatter
    Set Formatter = This.TargetFormatter
End Property

Public Property Set Formatter(ByVal a_value As cc_isr_MVVM.IValidationErrorFormatter)
    Set This.TargetFormatter = a_value
End Property

Public Property Get Target() As Excel.Range
    Set Target = This.Target
End Property

Public Property Set Target(ByVal a_value As Excel.Range)
    
    cc_isr_Core.GuardClauses.GuardDoubleInitialization This.Target, VBA.Information.TypeName(Me)
    cc_isr_Core.GuardClauses.GuardNullReference a_value
    Set This.Target = a_value

End Property

Private Sub IDynamicAdorner_Hide()
    This.TargetFormatter.Restore This.Target
End Sub

Private Sub IDynamicAdorner_Show(ByVal a_message As String)
    This.TargetFormatter.Apply This.Target, a_message
End Sub
