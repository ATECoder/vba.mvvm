VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IControlEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An observer that provides the infrastructure to relay MSForms.Control events."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   An observer that provides the infrastructure to relay MSForms.Control events. </summary>
''' <remarks>
''' What actually triggers the propagation of, for example, a new TextBox.Text value to the ViewModel property,
''' is the BeforeUpdate and Exit events, both of which are initially handled in the ControlEventsPunk class
''' implementation of this interface; an object that leverages the Win32 IConnectionPoint API to hook up event
''' handlers for the MSForms.Controls interface of our TextBox control (we covered that in the previous post).
'''
''' </remarks>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit

Public Sub RegisterHandler(ByVal a_handler As IHandleControlEvents)
End Sub

Public Sub OnEnter()
End Sub

Public Sub OnExit(ByVal a_cancel As MSForms.IReturnBoolean)
End Sub

Public Sub OnAfterUpdate()
End Sub

''' <summary>  </summary>
''' <remarks>
''' The first thing to run is the OnBeforeUpdate callback, which promptly iterates all registered
''' handlers (“observers”) and invokes their HandleBeforeUpdate method.
''' This callback, thus, propagates the event to the control property binding (e.g., TextBoxPropertyBinding),
''' and the IHandleControlEvents_HandleBeforeUpdate procedure runs as a result… which proceeds to forward it
''' to the PropertyBindingBase class with a call to OnBindingUpdateSourceOpportunity, a method with the
''' responsibility of conditionally invoking the ApplyToSource method.
''' </remarks>
Public Sub OnBeforeUpdate(ByVal a_cancel As MSForms.IReturnBoolean)
End Sub

'VF: event to capture BackSpace/Delete as well as Paste new tex /Cut all text in TextBox
'VF: paste/delete/cut easiest captured by _change, alternatively fiddle with KeyCodes
'see userform ExploreTextboxEvents
'Private Sub TextBox1_Change()
'    Label1.Caption = Label1.Caption & vbLf & "TextBox1_Change"
'End Sub
Public Sub OnChange()
End Sub
