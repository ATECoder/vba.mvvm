VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An object keeping command and property bindings in scope."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   An object keeping command and property bindings in scope. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit
Implements IAppContext
Implements cc_isr_Core.IDisposable

Private Type TState
    Bindings As cc_isr_MVVM.IBindingManager
    Commands As cc_isr_MVVM.ICommandManager
    Validation As cc_isr_MVVM.IValidationManager
    StringFormatterFactory As cc_isr_MVVM.IStringFormatterFactory
    DebugOutput As Boolean
    Disposed As Boolean
End Type

Private This As TState

''' <summary>   Initializes and returns this <see cref="AppContext"/> object. </summary>
''' <returns>   [<see cref="AppContext"/>]. </returns>
Public Function Initialize(Optional ByVal a_bindings As cc_isr_MVVM.IBindingManager, _
        Optional ByVal a_commands As cc_isr_MVVM.ICommandManager, _
        Optional ByVal a_validation As cc_isr_MVVM.IValidationManager, _
        Optional ByVal a_formatterFactory As cc_isr_MVVM.IStringFormatterFactory, _
        Optional ByVal a_adornerFactory As cc_isr_MVVM.IDynamicAdornerFactory, _
        Optional ByVal a_debugOutput As Boolean = False) As AppContext
Attribute Initialize.VB_Description = "Initializes a new MVVM application context."
    
    Me.DebugOutput = a_debugOutput
    
    ' temporal coupling: IStringFormatterVBFactory instance needs to be set before we init the binding manager.
    Me.InitStringFormatterFactory a_formatterFactory
    
    ' more temporal coupling...
    
    Me.InitCommandManager a_commands
    Me.InitBindingManager a_bindings
    Me.InitValidationManager a_validation, a_adornerFactory
    
    Set Initialize = Me
    
End Function

Friend Sub InitStringFormatterFactory(ByVal a_formatterFactory As cc_isr_MVVM.IStringFormatterFactory)
    
    If a_formatterFactory Is Nothing Then
        Set Me.StringFormatterFactory = New cc_isr_MVVM.StringFormatterFactory
    Else
        Set Me.StringFormatterFactory = a_formatterFactory
    End If

End Sub

Friend Sub InitCommandManager(ByVal a_commands As cc_isr_MVVM.ICommandManager)
    
    Dim p_manager As cc_isr_MVVM.ICommandManager
    If a_commands Is Nothing Then
        
        ' default to concrete implementation:
        Dim a_concreteManager As cc_isr_MVVM.CommandManager
        Set a_concreteManager = New cc_isr_MVVM.CommandManager
        a_concreteManager.DebugOutput = Me.DebugOutput
        Set p_manager = a_concreteManager
    
    Else
        
        'use injected instance (could be a test stub):
        Set p_manager = a_commands
    
    End If
    Set Me.Commands = p_manager
End Sub

Friend Sub InitBindingManager(ByVal a_binding As cc_isr_MVVM.IBindingManager)
    
    Dim p_manager As cc_isr_MVVM.IBindingManager
    
    If a_binding Is Nothing Then
    
        Dim a_concreteManager As cc_isr_MVVM.BindingManager
        Set a_concreteManager = Factory.NewBindingManager().Initialize(Me, Me.StringFormatterFactory, _
                                                                       a_debugOutput:=Me.DebugOutput)
        Set p_manager = a_concreteManager
        
    Else
        'use injected instance (could be a test stub):
        Set p_manager = a_binding
    
    End If
    Set Me.Bindings = p_manager
    
End Sub

Friend Sub InitValidationManager(ByVal a_validation As cc_isr_MVVM.IValidationManager, _
        ByVal a_adornerFactory As cc_isr_MVVM.IDynamicAdornerFactory)
    
    Dim p_manager As cc_isr_MVVM.IValidationManager
    
    If a_validation Is Nothing Then
        Dim a_concreteManager As cc_isr_MVVM.ValidationManager
        Set a_concreteManager = Factory.NewValidationManager().Initialize(a_adornerFactory)
        Set p_manager = a_concreteManager
    Else
        'use injected instance (could be a test stub):
        Set p_manager = a_validation
        
    End If
    
    Set Me.Validation = p_manager
    
End Sub

''' <summary>   Gets/sets the binding manager reference. </summary>
Public Property Get Bindings() As cc_isr_MVVM.IBindingManager
Attribute Bindings.VB_Description = "Gets/sets the binding manager reference."
    
    Set Bindings = This.Bindings

End Property

Public Property Set Bindings(ByVal a_value As cc_isr_MVVM.IBindingManager)
    
    cc_isr_Core.GuardClauses.GuardDoubleInitialization This.Bindings, VBA.Information.TypeName(Me)
    cc_isr_Core.GuardClauses.GuardNullReference a_value
    Set This.Bindings = a_value

End Property

''' <summary>   Gets/sets the command manager reference. </summary>
Public Property Get Commands() As ICommandManager
Attribute Commands.VB_Description = "Gets/sets the command manager reference."
    
    Set Commands = This.Commands

End Property

Public Property Set Commands(ByVal a_value As ICommandManager)
    
    cc_isr_Core.GuardClauses.GuardDoubleInitialization This.Commands, VBA.Information.TypeName(Me)
    cc_isr_Core.GuardClauses.GuardNullReference a_value
    Set This.Commands = a_value

End Property

''' <summary>   Gets/sets the validation manager reference. </summary>
Public Property Get Validation() As cc_isr_MVVM.IValidationManager
Attribute Validation.VB_Description = "Gets/sets the validation manager reference."
    
    Set Validation = This.Validation

End Property

Public Property Set Validation(ByVal a_value As cc_isr_MVVM.IValidationManager)
    
    cc_isr_Core.GuardClauses.GuardDoubleInitialization This.Validation, VBA.Information.TypeName(Me)
    cc_isr_Core.GuardClauses.GuardNullReference a_value
    Set This.Validation = a_value

End Property

''' <summary>   Gets/sets the IStringFormatter factory reference. </summary>
Public Property Get StringFormatterFactory() As IStringFormatterFactory
Attribute StringFormatterFactory.VB_Description = "Gets/sets the IStringFormatter factory reference."
    
    Set StringFormatterFactory = This.StringFormatterFactory

End Property

Public Property Set StringFormatterFactory(ByVal a_value As IStringFormatterFactory)
    
    cc_isr_Core.GuardClauses.GuardDoubleInitialization This.StringFormatterFactory, VBA.Information.TypeName(Me)
    cc_isr_Core.GuardClauses.GuardNullReference a_value
    Set This.StringFormatterFactory = a_value

End Property

''' <summary>   Gets/sets a value that determines whether binding managers produce debugger output. </summary>
Public Property Get DebugOutput() As Boolean
Attribute DebugOutput.VB_Description = "Gets/sets a value that determines whether binding managers produce debugger output."
    DebugOutput = This.DebugOutput
End Property

Public Property Let DebugOutput(ByVal a_value As Boolean)
    This.DebugOutput = a_value
End Property

Private Sub Dispose()
    If This.Disposed Then
        Debug.Print VBA.Information.TypeName(Me) & " instance was already disposed."
        Exit Sub
    End If
    If Not This.Bindings Is Nothing Then
        cc_isr_Core.DisposableExtensions.TryDispose This.Bindings
        Set This.Bindings = Nothing
    End If
    If Not This.Commands Is Nothing Then
        cc_isr_Core.DisposableExtensions.TryDispose This.Commands
        Set This.Commands = Nothing
    End If
    This.Disposed = True
End Sub

Private Sub Class_Terminate()
    If Not This.Disposed Then Dispose
End Sub

Private Property Get IAppContext_Bindings() As IBindingManager
    Set IAppContext_Bindings = This.Bindings
End Property

Private Property Get IAppContext_Commands() As ICommandManager
    Set IAppContext_Commands = This.Commands
End Property

Private Sub IAppContext_Dispose()
    Dispose
End Sub

Private Property Get IAppContext_StringFormatterFactory() As IStringFormatterFactory
    Set IAppContext_StringFormatterFactory = This.StringFormatterFactory
End Property

Private Property Get IAppContext_Validation() As IValidationManager
    Set IAppContext_Validation = This.Validation
End Property

' - - - - - - - - - - - - - - - -
'
' Disposable implementation
'
' - - - - - - - - - - - - - - - -

Private Sub IDisposable_Dispose()
    Dispose
End Sub
