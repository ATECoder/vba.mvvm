VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorksheetValidationAdorner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "An object that dynamically decorates a target Excel.Range object."
'@Folder rd_MVVM.Infrastructure.Validation.ErrorFormatting
'@ModuleDescription "An object that dynamically decorates a target Excel.Range object."
'@PredeclaredId
'@Exposed
Option Explicit
Implements IDynamicAdorner

Private Type TState
    Target As Excel.Range
    TargetFormatter As rd_MVVM.IValidationErrorFormatter
End Type

Private This As TState

Public Function Create(ByVal Target As Excel.Range, ByVal Formatter As rd_MVVM.IValidationErrorFormatter) As IDynamicAdorner
    Dim result As WorksheetValidationAdorner
    Set result = New WorksheetValidationAdorner
    Set result.Target = Target
    Set result.Formatter = Formatter
    Set Create = result
End Function

Friend Property Get Formatter() As rd_MVVM.IValidationErrorFormatter
    Set Formatter = This.TargetFormatter
End Property

Friend Property Set Formatter(ByVal RHS As rd_MVVM.IValidationErrorFormatter)
    Set This.TargetFormatter = RHS
End Property

Friend Property Get Target() As Excel.Range
    Set Target = This.Target
End Property

Friend Property Set Target(ByVal RHS As Excel.Range)
    GuardClauses.GuardDefaultInstance Me, WorksheetValidationAdorner
    GuardClauses.GuardDoubleInitialization This.Target, TypeName(Me)
    GuardClauses.GuardNullReference RHS
    Set This.Target = RHS
End Property

Private Sub IDynamicAdorner_Hide()
    This.TargetFormatter.Restore This.Target
End Sub

Private Sub IDynamicAdorner_Show(ByVal Message As String)
    This.TargetFormatter.Apply This.Target, Message
End Sub
