VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IValidationErrorFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An object that describes how a control changes its appearance given a validation error."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   An object that describes how a control changes its appearance given a validation error. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit

''' <summary>   Applies error styling to the target given the specified validation error message. </summary>
Public Sub Apply(ByVal a_target As Object, ByVal a_message As String)
Attribute Apply.VB_Description = "Applies error styling to the target given the specified validation error message."
End Sub

''' <summary>   Removes error styling on the target. </summary>
Public Sub Restore(ByVal a_target As Object)
Attribute Restore.VB_Description = "Removes error styling on the target."
End Sub
