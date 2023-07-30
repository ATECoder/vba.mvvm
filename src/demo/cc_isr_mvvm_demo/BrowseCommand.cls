VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BrowseCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   A browse command for the Example ViewModel implementation. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit
Implements ICommand

Private Function ICommand_CanExecute(ByVal a_context As Object) As Boolean
    If TypeOf a_context Is ExampleViewModel Then
        Dim ViewModel As ExampleViewModel
        Set ViewModel = a_context
        ICommand_CanExecute = ViewModel.SomeOtherOption
    End If
End Function

Private Property Get ICommand_Description() As String
    ICommand_Description = "Browse for source file..."
End Property

Private Sub ICommand_Execute(ByVal a_context As Object)
    Dim ViewModel As ExampleViewModel
    Set ViewModel = a_context
    
    ViewModel.SomeFilePath = "TEST"
    
    MsgBox "Implement me!", vbInformation, "Hello from " & VBA.Information.TypeName(Me)
End Sub
