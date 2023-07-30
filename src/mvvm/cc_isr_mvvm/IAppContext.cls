VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAppContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An object keeping command and property bindings in scope."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   An object keeping command and property bindings in scope. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit

''' <summary>   Gets an object responsible for managing command bindings. </summary>
Public Property Get Commands() As cc_isr_MVVM.ICommandManager
Attribute Commands.VB_Description = "Gets an object responsible for managing command bindings."
End Property

''' <summary>   Gets an object responsible for managing property bindings. </summary>
Public Property Get Bindings() As cc_isr_MVVM.IBindingManager
Attribute Bindings.VB_Description = "Gets an object responsible for managing property bindings."
End Property

''' <summary>   Gets an object responsible for managing binding validation errors. </summary>
Public Property Get Validation() As cc_isr_MVVM.IValidationManager
Attribute Validation.VB_Description = "Gets an object responsible for managing binding validation errors."
End Property

''' <summary>   Gets an abstract factory that creates objects that can format strings. </summary>
Public Property Get StringFormatterFactory() As cc_isr_MVVM.IStringFormatterFactory
Attribute StringFormatterFactory.VB_Description = "Gets an abstract factory that creates objects that can format strings."
End Property
