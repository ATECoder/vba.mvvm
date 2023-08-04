Attribute VB_Name = "CommandManagerTests"
'@Folder Tests.Bindings
'@TestModule
Option Explicit
Option Private Module

Private Assert As cc_isr_Test_Fx.Assert

Private Type TState
    ExpectedErrNumber As Long
    ExpectedErrSource As String
    ExpectedErrorCaught As Boolean
    
    ConcreteSUT As CommandManager
    AbstractSUT As ICommandManager
    
    BindingContext As TestBindingObject
    Command As TestCommand
    
End Type

Private Test As TState

'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = cc_isr_Test_Fx.Assert
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    Set Test.ConcreteSUT = New CommandManager
    Set Test.AbstractSUT = Test.ConcreteSUT
    Set Test.BindingContext = New TestBindingObject
    Set Test.Command = New TestCommand
End Sub

'@TestCleanup
Private Sub TestCleanup()
    Set Test.ConcreteSUT = Nothing
    Set Test.AbstractSUT = Nothing
    Set Test.BindingContext = Nothing
    Set Test.Command = Nothing
End Sub

Private Sub ExpectError()
    Dim Message As String
    If Err.Number = Test.ExpectedErrNumber Then
        If (Test.ExpectedErrSource = vbNullString) Or (Err.Source = Test.ExpectedErrSource) Then
            Test.ExpectedErrorCaught = True
        Else
            Message = "An error was raised, but not from the expected source. " & _
                      "Expected: '" & TypeName(Test.ConcreteSUT) & "'; Actual: '" & Err.Source & "'."
        End If
    ElseIf Err.Number <> 0 Then
        Message = "An error was raised, but not with the expected number. Expected: '" & Test.ExpectedErrNumber & "'; Actual: '" & Err.Number & "'."
    Else
        Message = "No error was raised."
    End If
    
    If Not Test.ExpectedErrorCaught Then Assert.Fail Message
End Sub

Private Function DefaultTargetCommandBindingFor(ByVal ProgID As String, ByRef outTarget As Object) As ICommandBinding
    Set outTarget = CreateObject(ProgID)
    Set DefaultTargetCommandBindingFor = Test.AbstractSUT.BindCommand(Test.BindingContext, outTarget, Test.Command)
End Function

'@TestMethod("DefaultCommandTargetBindings")
Private Function BindCommand_BindsCommandButton() As cc_isr_Test_Fx.Assert
    Dim Target As Object
    With DefaultTargetCommandBindingFor(FormsProgID.CommandButtonProgId, outTarget:=Target)
        Set BindCommand_BindsCommandButton = Assert.AreSame(Test.Command, .Command, "")
        Set BindCommand_BindsCommandButton = Assert.AreSame(Target, .Target, "")
    End With
End Function

'@TestMethod("DefaultCommandTargetBindings")
Private Function BindCommand_BindsCheckBox() As cc_isr_Test_Fx.Assert
    Dim Target As Object
    With DefaultTargetCommandBindingFor(FormsProgID.CheckBoxProgId, outTarget:=Target)
        Set BindCommand_BindsCheckBox = Assert.AreSame(Test.Command, .Command, "")
        Set BindCommand_BindsCheckBox = Assert.AreSame(Target, .Target, "")
    End With
End Function

'@TestMethod("DefaultCommandTargetBindings")
Private Function BindCommand_BindsImage() As cc_isr_Test_Fx.Assert
    Dim Target As Object
    With DefaultTargetCommandBindingFor(FormsProgID.ImageProgId, outTarget:=Target)
        Set BindCommand_BindsImage = Assert.AreSame(Test.Command, .Command, "")
        Set BindCommand_BindsImage = Assert.AreSame(Target, .Target, "")
    End With
End Function

'@TestMethod("DefaultCommandTargetBindings")
Private Function BindCommand_BindsLabel() As cc_isr_Test_Fx.Assert
    Dim Target As Object
    With DefaultTargetCommandBindingFor(FormsProgID.LabelProgId, outTarget:=Target)
        Set BindCommand_BindsLabel = Assert.AreSame(Test.Command, .Command, "")
        Set BindCommand_BindsLabel = Assert.AreSame(Target, .Target, "")
    End With
End Function


