Attribute VB_Name = "Example"
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Runs the example. </summary>
''' <remarks>
''' </remarks>
''' - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Public Type FormSizeType
    Height As Long
    Width As Long
End Type

''' <summary>   Runs the MVVM example UI. </summary>
''' <remarks>
''' VF: Windows 10 is having a hard time handling multiple monitors, especially if different resolutions and more so if
''' legacy applications like the VBE keeps shrinking the userform in the VBE and thus showing the shrunk form.
''' <- must counteract this ugly Windows bug by specifying Height and Width of IView
''' rendering engine was changed from 2010 to 2013
''' should go into IView, shouldn't it?
''' </remarks>
Public Sub Run()
Attribute Run.VB_Description = "Runs the MVVM example UI."

    ' here a more elaborate application would wire-up dependencies for complex commands,
    ' and then property-inject them into the ViewModel via a factory method e.g. SomeViewModel.Create(args).

    Dim a_viewModel As New ExampleViewModel
    
    ' ViewModel properties can be set before or after it's wired to the View.
    ' ViewModel.SourcePath = "TEST"
    a_viewModel.SomeOption = True
    
    Set a_viewModel.SomeCommand = New BrowseCommand
    
    Dim p_appContext As AppContext
    Set p_appContext = cc_isr_MVVM.Factory.NewAppContext().Initialize(a_debugOutput:=True)
    
    a_viewModel.BooleanProperty = False
    a_viewModel.ByteProperty = 240
    a_viewModel.DateProperty = VBA.DateTime.Now + 2
    a_viewModel.DoubleProperty = 85
    a_viewModel.StringProperty = "Beta"
    a_viewModel.LongProperty = -42
    
    Dim p_view As IView
    
    Set p_view = ExampleView.Create(p_appContext, a_viewModel)
    
    If p_view.ShowDialog Then
        Debug.Print a_viewModel.SomeFilePath, a_viewModel.SomeOption, a_viewModel.SomeOtherOption
    Else
        Debug.Print "Dialog was cancelled."
    End If
    
    cc_isr_Core.DisposableExtensions.TryDispose p_appContext
    
End Sub

''' <summary>   Runs the MVVM example dynamic UI. </summary>
Public Sub DynamicRun()

    Dim p_context As IAppContext
    Set p_context = cc_isr_MVVM.Factory.NewAppContext().Initialize()
    
    Dim p_viewModel As New ExampleViewModel
    
    Dim p_view As IView
    Dim p_formSize As FormSizeType
    'VF: in non-dynamic userforms like ExampleView the controls stay put so I use the right bottom most controls as anchor point like Me.Width = LastControl.left+LastControl.width + OffsetWidthPerOfficeVersion <- yes, userform are rendered differently depending on the version of Office 2007, ....
        'if sizing dynamically I would proceed likewise somehow with the (right bottom most) container <- is going to take quite an amount of code :-(
    With p_formSize
        .Height = 180 'some value that work in 2019, and somehow in 2010, too
        .Width = 230
    End With
    
    Set p_view = ExampleDynamicView.Create(p_context, p_viewModel, p_formSize)
    ' or keep factory .Create 'clean'?
'    With ExampleDynamicView.Create(p_context, p_viewModel, p_formSize)
'        .SizeView 'not implemented
'        .ShowDialog
'        'payload DoSomething if not cancelled
'    End With
        
    Debug.Print p_view.ShowDialog
    
End Sub

