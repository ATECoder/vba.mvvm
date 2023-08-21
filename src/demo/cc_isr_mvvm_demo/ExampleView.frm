VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExampleView 
   Caption         =   "ExampleView"
   ClientHeight    =   4695
   ClientLeft      =   -150
   ClientTop       =   -510
   ClientWidth     =   3675
   OleObjectBlob   =   "ExampleView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ExampleView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "An example implementation of a View."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   An non-dynamic view for the example. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Implements IView
Implements ICancellable
Option Explicit

Private Type TView

    Context As cc_isr_MVVM.IAppContext
    
    ' IView state
    ViewModel As ExampleViewModel
    
    ' ICancellable state:
    IsCancelled As Boolean
    
End Type

Private This As TView

''' <summary>   A factory method to create new instances of this View,
'''             already wired-up to a ViewModel. </summary>
Public Function Create(ByVal a_context As IAppContext, ByVal a_viewModel As ExampleViewModel) As IView
Attribute Create.VB_Description = "A factory method to create new instances of this View, already wired-up to a ViewModel."

    Dim p_source As String
    p_source = ThisWorkbook.VBProject.Name & "." & VBA.Information.TypeName(Me) & ".Create"
    cc_isr_Core_IO.GuardClauses.GuardNonDefaultInstance Me, ExampleView, p_source
    cc_isr_Core_IO.GuardClauses.GuardNullReference a_viewModel, p_source
    cc_isr_Core_IO.GuardClauses.GuardNullReference a_context, p_source
    
    Dim result As ExampleView
    Set result = New ExampleView
    
    Set result.Context = a_context
    Set result.ViewModel = a_viewModel
    
    Set Create = result
    
End Function

Private Property Get IsDefaultInstance() As Boolean
    IsDefaultInstance = Me Is ExampleView
End Property

''' <summary>   Gets/sets the ViewModel to use as a context for property and command bindings. </summary>
Public Property Get ViewModel() As ExampleViewModel
Attribute ViewModel.VB_Description = "Gets/sets the ViewModel to use as a context for property and command bindings."
    Set ViewModel = This.ViewModel
End Property

Public Property Set ViewModel(ByVal a_value As ExampleViewModel)

    Dim p_source As String
    p_source = ThisWorkbook.VBProject.Name & "." & VBA.Information.TypeName(Me) & ".ViewModel"
    cc_isr_Core_IO.GuardClauses.GuardDefaultInstance Me, ExampleView, p_source
    cc_isr_Core_IO.GuardClauses.GuardNullReference a_value, p_source
    
    Set This.ViewModel = a_value
    
End Property

''' <summary>   Gets/sets the MVVM application context. </summary>
Public Property Get Context() As cc_isr_MVVM.IAppContext
Attribute Context.VB_Description = "Gets/sets the MVVM application context."
    Set Context = This.Context
End Property

Public Property Set Context(ByVal a_value As cc_isr_MVVM.IAppContext)

    Dim p_source As String
    p_source = ThisWorkbook.VBProject.Name & "." & VBA.Information.TypeName(Me) & ".Context"
    cc_isr_Core_IO.GuardClauses.GuardDefaultInstance Me, ExampleView, p_source
    cc_isr_Core_IO.GuardClauses.GuardDoubleInitialization This.Context, p_source
    cc_isr_Core_IO.GuardClauses.GuardNullReference a_value, p_source
    
    Set This.Context = a_value
    
End Property

Private Sub BindViewModelCommands()
    
    Me.Context.Commands.BindCommand Me.ViewModel, Me.OkButton, _
            cc_isr_MVVM.Factory.NewAcceptCommand().Initialize(Me, This.Context.Validation)
            
    Me.Context.Commands.BindCommand Me.ViewModel, Me.CancelButton, _
            cc_isr_MVVM.Factory.NewCancelCommand().Initialize(Me)
    
    Me.Context.Commands.BindCommand Me.ViewModel, Me.BrowseButton, ViewModel.SomeCommand
    '...

End Sub

Private Sub BindViewModelProperties()

    ' Binding to a Label control without a target property like this creates a CaptionPropertyBinding.
    ' This type of binding defaults to a one-way binding (source property -> target property)
    ' that will update the target caption if the source changes (source property setter must invoke OnPropertyChanged):
    Me.Context.Bindings.BindPropertyPath Me.ViewModel, "Instructions", Me.InstructionsLabel
    
    ' If we know we're not going to change the caption at any point,
    ' we can always make the binding one-time (source property -> target property).
    ' Binding to an OptionButton control without specifying a target property creates an OptionButtonBinding,
    ' and because we're binding to a String property the target property is inferred to be the Caption:
    Me.Context.Bindings.BindPropertyPath Me.ViewModel, "SomeOptionName", Me.OptionButton1, a_mode:=OneTimeBinding
    
    Me.Context.Bindings.BindPropertyPath Me.ViewModel, "SomeOtherOptionName", Me.OptionButton2, a_mode:=OneTimeBinding
    
    ' Binding to a TextBox control we can specify a format string to format the control when
    ' it loses focus, and by setting the binding's UpdateTrigger to OnKeyPress we get to
    ' use a KeyValidator that can prevent invalid (here, non-numeric) inputs.
    ' Without specifying a target property, we're binding to TextBox.Text regardless of the data type of the source property:
    Me.Context.Bindings.BindPropertyPath Me.ViewModel, "SomeAmount", Me.AmountBox, _
        a_formatString:="{0:C2}", _
        a_updateTrigger:=OnKeyPress, _
        a_validator:=cc_isr_MVVM.Factory.NewDecimalKeyValidator
    
    ' Binding a Date property on the ViewModel to a TextBox control works best with a value converter.
    ' The converter must be able to convert the specified format string into a Date, and back.
    ' We can handle validation errors by providing a ValidationErrorAdorner instance:
    Me.Context.Bindings.BindPropertyPath Me.ViewModel, "SomeDate", Me.TextBox1, _
        a_formatString:="{0:MMMM dd, yyyy}", _
        a_converter:=StringToDateConverter.Default, _
        a_validator:=cc_isr_MVVM.Factory.NewRequiredStringValidator, _
        a_validationAdorner:=cc_isr_MVVM.Factory.NewValidationErrorAdorner().Initialize( _
             a_target:=Me.TextBox1, _
            a_targetFormatter:=cc_isr_MVVM.Factory.NewValidationErrorFormatter.WithErrorBorderColor.WithErrorBackgroundColor)
    
    ' OptionButton controls automatically bind their value to a Boolean property:
    
    Me.Context.Bindings.BindPropertyPath Me.ViewModel, "SomeOption", Me.OptionButton1
    
    Me.Context.Bindings.BindPropertyPath Me.ViewModel, "SomeOtherOption", Me.OptionButton2
    
    ' When binding an array property to a ComboBox target, the List property is the implicit target:
    
    Me.Context.Bindings.BindPropertyPath Me.ViewModel, "SomeItems", Me.ComboBox1
    
    ' If we want we can bind a String property to automatically bind to ComboBox.Text:
    
    Me.Context.Bindings.BindPropertyPath Me.ViewModel, "SelectedItemText", Me.ComboBox1
    
    ' Or (and?) we want we can bind a Long property to automatically bind to ComboBox.ListIndex:
    
    Me.Context.Bindings.BindPropertyPath Me.ViewModel, "SelectedItemIndex", Me.ComboBox1
    
    ' Binding to any other source property data type binds to ComboBox.Value;
    ' that's especially useful when the List has multiple columns and the first (the Value!) contains some hidden unique ID.
    
    ' So MVVM works for a MSForms UI.
    ' What if the binding target was something else?
    
    ' a worksheet cell's value?
    ' .BindPropertyPath Me.ViewModel, "SelectedItemText", Sheet1.Range("A1"), "Value"
    
    ' a...chart's title?
    ' .BindPropertyPath Me.ViewModel, "Instructions", Sheet1.ChartObjects("Chart 1"), "Chart.ChartTitle.Text"
    
    ' ...I've created a monster, haven't I?
    
End Sub

Private Sub BindViewControls()
    
    Dim p_LayoutView As ILayout
    Set p_LayoutView = cc_isr_MVVM.Factory.NewLayout().Initialize(Me, 15, 30)
    With p_LayoutView
        .BindControlLayout Me, Me.InstructionsLabel, LeftAnchor + RightAnchor
        .BindControlLayout Me, Me.AmountBox, LeftAnchor
        .BindControlLayout Me, Me.BrowseButton, RightAnchor
        .BindControlLayout Me, Me.TextBox1, LeftAnchor + RightAnchor
        .BindControlLayout Me, Me.ComboBox1, LeftAnchor + RightAnchor
        .BindControlLayout Me, Me.OptionsFrame, LeftAnchor + RightAnchor
        .BindControlLayout Me.OptionsFrame, Me.OptionButton1, LeftAnchor
        .BindControlLayout Me.OptionsFrame, Me.OptionButton2, LeftAnchor
    End With
    p_LayoutView.ResizeLayout
    
End Sub

Private Sub InitializeBindings()
    If ViewModel Is Nothing Then Exit Sub
    BindViewModelProperties
    BindViewModelCommands
    BindViewControls
    This.Context.Bindings.Apply Me.ViewModel
End Sub

Private Sub OnCancel()
    This.IsCancelled = True
    Me.Hide
End Sub

Private Property Get ICancellable_IsCancelled() As Boolean
    ICancellable_IsCancelled = This.IsCancelled
End Property

Private Sub ICancellable_OnCancel()
    OnCancel
End Sub

Private Sub IView_Hide()
    Me.Hide
End Sub

Private Sub IView_Show()
    InitializeBindings
    Me.Show vbModal
End Sub

Private Function IView_ShowDialog() As Boolean
    InitializeBindings
    Me.Show vbModal
    IView_ShowDialog = Not This.IsCancelled
End Function

Private Property Get IView_ViewModel() As Object
    Set IView_ViewModel = This.ViewModel
End Property

Private Sub UserForm_QueryClose(a_cancel As Integer, a_closeMode As Integer)
    If a_closeMode = VbQueryClose.vbFormControlMenu Then
        a_cancel = True
        OnCancel
    End If
End Sub
