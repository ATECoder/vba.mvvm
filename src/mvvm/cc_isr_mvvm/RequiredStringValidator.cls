VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RequiredStringValidator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "A validator that requires a string value to be non-empty."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   A validator that requires a string value to be non-empty. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit
Implements IValueValidator

' VF: failed space bar/blank spaces between words entered in TextBox that's why TextBox event (IValueValidator_Trigger) switched to onChange from onKeyPress
'but OnChange was also needed to capture Cut and Paste in TextBox
Private Function IValueValidator_IsValid(ByVal a_value As Variant, ByVal a_source As IBindingPath, ByVal a_target As IBindingPath) As Boolean
    IValueValidator_IsValid = Len(Trim$(a_value)) > 0
End Function

Private Property Get IValueValidator_Message() As String
    IValueValidator_Message = "Value cannot be empty."
End Property

Private Property Get IValueValidator_Trigger() As BindingUpdateSourceTrigger
    
    ' VF: must use OnChange (tried OnKeyPres first) despite sucking some performance
    ' IValueValidator_Trigger = OnExit
    ' IValueValidator_Trigger = BindingUpdateSourceTrigger.OnKeyPress 'does not catch BackSpace and Delete so empty TextBox is skipped
    ' tried blindly first:
    ' IValueValidator_Trigger = BindingUpdateSourceTrigger.OnPropertyChanged 'does not catch BackSpace and Delete either
    ' catches cut/paste and backspace/delete
    
    IValueValidator_Trigger = BindingUpdateSourceTrigger.OnChange
    
    ' OnExit and OnPropertyChanged are still somewhat ambiguously defined,
    ' but OnExit gets to keep the target control focused if there�s a validation error,
    ' and OnPropertyChanged is [currently] firing for every keypress in a TextBox, after the control�s Text property updates.
    ' Disambiguating the terms would mean breaking with MSForms� meaning of �updated�� which may actually be a good thing:
    ' OnPropertyChanged would basically fire on exit but without a possible navigation-blocking validation error, and then
    ' OnKeyPress mode would still need to behave like OnPropertyChanged as far as updating the source goes.
    
    ' VF: description on web page does not match code
    ' What actually triggers the propagation of the new TextBox.Text value to the ViewModel property, is the
    ' BeforeUpdate and Exit events, both of which are initially handled in the ControlEventsPunk class,
    ' an object that leverages the Win32 IConnectionPoint API to hook up event handlers for the MSForms.Controls
    ' interface of our TextBox control
    
End Property
