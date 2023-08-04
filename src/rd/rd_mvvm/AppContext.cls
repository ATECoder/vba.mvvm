VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "An object keeping command and property bindings in scope."
'@Folder rd_MVVM.Infrastructure
'@ModuleDescription "An object keeping command and property bindings in scope."
'@PredeclaredId
'@Exposed
Option Explicit
Implements IAppContext
Implements IDisposable

Public Enum StringFormatterStrategy
    UseDotNetStringFormatSyntax
    UseVBStringFormatSyntax
End Enum

Private Type TState
    Bindings As rd_MVVM.IBindingManager
    Commands As rd_MVVM.ICommandManager
    Validation As rd_MVVM.IValidationManager
    
    DefaultStringFormatSyntax As StringFormatterStrategy
    StringFormatterFactory As IStringFormatterFactory
    
    DebugOutput As Boolean
    Disposed As Boolean
End Type

Private This As TState

'@Description "Creates a new MVVM application context."
Public Function Create( _
Optional ByVal Bindings As rd_MVVM.IBindingManager, _
Optional ByVal Commands As rd_MVVM.ICommandManager, _
Optional ByVal Validation As rd_MVVM.IValidationManager, _
Optional ByVal FormatterFactory As rd_MVVM.IStringFormatterFactory, _
Optional ByVal AdornerFactory As rd_MVVM.IDynamicAdornerFactory, _
Optional ByVal DebugOutput As Boolean = False) As rd_MVVM.AppContext
Attribute Create.VB_Description = "Creates a new MVVM application context."
    
    GuardClauses.GuardNonDefaultInstance Me, rd_MVVM.AppContext
    
    Dim result As rd_MVVM.AppContext
    Set result = New rd_MVVM.AppContext
    result.DebugOutput = DebugOutput
    
    'temporal coupling: IStringFormatterVBFactory instance needs to be set before we init the binding manager.
    InitStringFormatterFactory result, FormatterFactory
    
    'more temporal coupling...
    InitCommandManager result, Commands
    InitBindingManager result, Bindings
    InitValidationManager result, Validation, AdornerFactory
    
    Set Create = result
End Function

Private Sub InitStringFormatterFactory(ByVal Context As rd_MVVM.AppContext, ByVal FormatterFactory As rd_MVVM.IStringFormatterFactory)
    Dim Factory As rd_MVVM.IStringFormatterFactory
    If FormatterFactory Is Nothing Then
        Select Case This.DefaultStringFormatSyntax
        
            Case StringFormatterStrategy.UseDotNetStringFormatSyntax
                Set Factory = New rd_MVVM.StringFormatterNetFactory
                
            Case StringFormatterStrategy.UseVBStringFormatSyntax
                Set Factory = New rd_MVVM.StringFormatterVBFactory
                
        End Select
    Else
        Set Factory = FormatterFactory
    End If
    Set Context.StringFormatterFactory = Factory
End Sub

Private Sub InitCommandManager(ByVal Context As rd_MVVM.AppContext, ByVal Commands As rd_MVVM.ICommandManager)
    Dim Manager As rd_MVVM.ICommandManager
    If Commands Is Nothing Then
        'default to concrete implementation:
        Dim ConcreteManager As rd_MVVM.CommandManager
        Set ConcreteManager = New rd_MVVM.CommandManager
        ConcreteManager.DebugOutput = Context.DebugOutput
        Set Manager = ConcreteManager
    Else
        'use injected instance (could be a test stub):
        Set Manager = Commands
    End If
    Set Context.Commands = Manager
End Sub

Private Sub InitBindingManager(ByVal Context As rd_MVVM.AppContext, ByVal Bindings As rd_MVVM.IBindingManager)
    Dim Manager As rd_MVVM.IBindingManager
    
    If Bindings Is Nothing Then
        Dim ConcreteManager As rd_MVVM.BindingManager
        Set ConcreteManager = rd_MVVM.BindingManager.Create(Context, Context.StringFormatterFactory, DebugOutput:=Context.DebugOutput)
        Set Manager = ConcreteManager
        
    Else
        'use injected instance (could be a test stub):
        Set Manager = Bindings
    
    End If
    Set Context.Bindings = Manager
End Sub

Private Sub InitValidationManager(ByVal Context As rd_MVVM.AppContext, ByVal Validation As rd_MVVM.IValidationManager, ByVal AdornerFactory As rd_MVVM.IDynamicAdornerFactory)
    Dim Manager As rd_MVVM.IValidationManager
    
    If Validation Is Nothing Then
        Dim ConcreteManager As rd_MVVM.ValidationManager
        Set ConcreteManager = rd_MVVM.ValidationManager.Create(AdornerFactory)
        Set Manager = ConcreteManager
    Else
        'use injected instance (could be a test stub):
        Set Manager = Validation
        
    End If
    
    Set Context.Validation = Manager
End Sub

'@Description "Gets/sets the binding manager reference."
Public Property Get Bindings() As rd_MVVM.IBindingManager
Attribute Bindings.VB_Description = "Gets/sets the binding manager reference."
    GuardClauses.GuardDefaultInstance Me, rd_MVVM.AppContext
    Set Bindings = This.Bindings
End Property

Friend Property Set Bindings(ByVal RHS As rd_MVVM.IBindingManager)
    GuardClauses.GuardDefaultInstance Me, rd_MVVM.AppContext
    GuardClauses.GuardDoubleInitialization This.Bindings, TypeName(Me)
    GuardClauses.GuardNullReference RHS
    Set This.Bindings = RHS
End Property

'@Description "Gets/sets the command manager reference."
Public Property Get Commands() As ICommandManager
Attribute Commands.VB_Description = "Gets/sets the command manager reference."
    GuardClauses.GuardDefaultInstance Me, rd_MVVM.AppContext
    Set Commands = This.Commands
End Property

Friend Property Set Commands(ByVal RHS As ICommandManager)
    GuardClauses.GuardDefaultInstance Me, rd_MVVM.AppContext
    GuardClauses.GuardDoubleInitialization This.Commands, TypeName(Me)
    GuardClauses.GuardNullReference RHS
    Set This.Commands = RHS
End Property

'@Description "Gets/sets the validation manager reference."
Public Property Get Validation() As IValidationManager
Attribute Validation.VB_Description = "Gets/sets the validation manager reference."
    GuardClauses.GuardDefaultInstance Me, rd_MVVM.AppContext
    Set Validation = This.Validation
End Property

Friend Property Set Validation(ByVal RHS As IValidationManager)
    GuardClauses.GuardDefaultInstance Me, rd_MVVM.AppContext
    GuardClauses.GuardDoubleInitialization This.Validation, TypeName(Me)
    GuardClauses.GuardNullReference RHS
    Set This.Validation = RHS
End Property

'@Description "Gets/sets the IStringFormatter factory reference."
Public Property Get StringFormatterFactory() As IStringFormatterFactory
Attribute StringFormatterFactory.VB_Description = "Gets/sets the IStringFormatter factory reference."
    'GuardClauses.GuardDefaultInstance Me, rd_MVVM.AppContext
    Set StringFormatterFactory = This.StringFormatterFactory
End Property

Friend Property Set StringFormatterFactory(ByVal RHS As IStringFormatterFactory)
    'GuardClauses.GuardDefaultInstance Me, rd_MVVM.AppContext
    GuardClauses.GuardDoubleInitialization This.StringFormatterFactory, TypeName(Me)
    GuardClauses.GuardNullReference RHS
    Set This.StringFormatterFactory = RHS
End Property

'@Description "Gets/sets a value that determines the StringFormat syntax to use when no implementation is provided."
Public Property Get DefaultStringFormatSyntax() As StringFormatterStrategy
Attribute DefaultStringFormatSyntax.VB_Description = "Gets/sets a value that determines the StringFormat syntax to use when no implementation is provided."
    GuardClauses.GuardNonDefaultInstance Me, rd_MVVM.AppContext
    DefaultStringFormatSyntax = This.DefaultStringFormatSyntax
End Property

Public Property Let DefaultStringFormatSyntax(ByVal RHS As StringFormatterStrategy)
    GuardClauses.GuardNonDefaultInstance Me, rd_MVVM.AppContext
    This.DefaultStringFormatSyntax = RHS
End Property

'@Description "Gets/sets a value that determines whether binding managers produce debugger output."
Public Property Get DebugOutput() As Boolean
Attribute DebugOutput.VB_Description = "Gets/sets a value that determines whether binding managers produce debugger output."
    DebugOutput = This.DebugOutput
End Property

Public Property Let DebugOutput(ByVal RHS As Boolean)
    This.DebugOutput = RHS
End Property

Private Sub Dispose()
    If This.Disposed Then
        Debug.Print TypeName(Me) & " instance was already disposed."
        Exit Sub
    End If
    If Not This.Bindings Is Nothing Then
        Disposable.TryDispose This.Bindings
        Set This.Bindings = Nothing
    End If
    If Not This.Commands Is Nothing Then
        Disposable.TryDispose This.Commands
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

Private Sub IDisposable_Dispose()
    Dispose
End Sub
