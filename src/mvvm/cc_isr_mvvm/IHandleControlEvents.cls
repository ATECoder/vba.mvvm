VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IHandleControlEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An object that can be registered as a handler for IControlEvents callbacks."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   An object that can be registered as a handler for IControlEvents callbacks. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit

''' <summary>   A callback that handles the 'Enter' event. </summary>
Public Sub HandleEnter()
Attribute HandleEnter.VB_Description = "A callback that handles the 'Enter' event."
End Sub

''' <summary>   A callback that handles the 'Exit' event. </summary>
Public Sub HandleExit(ByRef a_cancel As Boolean)
Attribute HandleExit.VB_Description = "A callback that handles the 'Exit' event."
End Sub

''' <summary>   A callback that handles the 'AfterUpdate' event. </summary>
Public Sub HandleAfterUpdate()
Attribute HandleAfterUpdate.VB_Description = "A callback that handles the 'AfterUpdate' event."
End Sub

''' <summary>   A callback that handles the 'BeforeUpdate' event. </summary>
Public Sub HandleBeforeUpdate(ByRef a_cancel As Boolean)
Attribute HandleBeforeUpdate.VB_Description = "A callback that handles the 'BeforeUpdate' event."
End Sub

''' <summary>   A callback that handles the 'Change' event. </summary>
''' <remarks>   VF: added: required to capture Cut/Paste and Backspace/Delete in TextBox. </remarks>
Public Sub HandleChange()
End Sub
