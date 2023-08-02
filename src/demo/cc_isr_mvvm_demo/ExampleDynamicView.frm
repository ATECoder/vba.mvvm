VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExampleDynamicView 
   Caption         =   "ExampleDynamicView"
   ClientHeight    =   3015
   ClientLeft      =   -450
   ClientTop       =   -1515
   ClientWidth     =   4560
   OleObjectBlob   =   "ExampleDynamicView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ExampleDynamicView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Dynamic view for the example. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit
Implements IView
Implements ICancellable

Private Type TState
    Context As cc_isr_MVVM.IAppContext
    ViewModel As ExampleViewModel
    IsCancelled As Boolean
    Factory As cc_isr_MVVM.Factory
End Type

Private This As TState

''' <summary>   Creates a new instance of this form. </summary>
Public Function Create(ByVal a_context As cc_isr_MVVM.IAppContext, _
        ByVal a_viewModel As ExampleViewModel, ByRef a_formSize As FormSizeType) As IView
    Dim result As ExampleDynamicView
    Set result = New ExampleDynamicView
    Set result.Context = a_context
    Set result.ViewModel = a_viewModel
    result.Height = a_formSize.Height
    result.Width = a_formSize.Width
    Set Create = result
End Function

Public Property Get Context() As cc_isr_MVVM.IAppContext
    Set Context = This.Context
End Property

Public Property Set Context(ByVal a_value As cc_isr_MVVM.IAppContext)
    Set This.Context = a_value
End Property

Public Property Get ViewModel() As Object
    Set ViewModel = This.ViewModel
End Property

Public Property Set ViewModel(ByVal a_value As Object)
    Set This.ViewModel = a_value
End Property

Public Sub SizeView(ByVal a_height As Long, ByVal a_width As Long)
    Me.Height = a_height
    Me.Width = a_width
End Sub

Private Sub OnCancel()
    This.IsCancelled = True
    Me.Hide
End Sub

Private Sub InitializeView()
    
    Set This.Factory = cc_isr_MVVM.Factory
    
    Dim p_layout As IContainerLayout
    Set p_layout = cc_isr_MVVM.Factory.NewContainerLayout().Initialize(Me.Controls, cc_isr_MVVM.LayoutDirection.TopToBottom)
    
    With This.Factory.NewDynamicControls().Initialize(This.Context, p_layout)
        
        With .LabelFor("All controls on this form are created at run-time.")
            .Font.Bold = True
            .Name = "LabelIndex1"
        End With
        
        With .LabelFor(This.Factory.NewBindingPath().Initialize(This.ViewModel, "Instructions"))
            .Name = "LabelIndex2"
        End With
        
        'VF: refactor free string to some enum PropertyName ("StringProperty", "CurrencyProperty") throughout (?) [when I frame a question mark in parentheses is not really a question but a rhetorical question, meaning I am pretty sure of the correct answer]
        With .TextBoxFor(This.Factory.NewBindingPath().Initialize(This.ViewModel, "StringProperty"), _
                    a_validator:=This.Factory.NewRequiredStringValidator, _
                    a_titleSource:="Some String:")
                .Name = "TextBoxIndex1"
        End With
        
        With .TextBoxFor(This.Factory.NewBindingPath().Initialize(This.ViewModel, "CurrencyProperty"), _
                    a_formatString:="{0:C2}", _
                    a_validator:=This.Factory.NewDecimalKeyValidator, _
                    a_titleSource:="Some Amount:")
                .Name = "TextBoxIndex2"
        End With
        
        ' ToDo: 'VF: needs validation .CanExecute(This.Context) before .Show
        ' (as textbox1 has focus and is empty and when moving to this close button,
        ' tb1 is validated and OnClick is disabled leaving the user out in the rain)
        With .CommandButtonFor(This.Factory.NewAcceptCommand().Initialize(Me, This.Context.Validation), _
                This.ViewModel, "Close")
            .Name = "CommandButtonIndex1"
        End With
        
    End With
    
    This.Context.Bindings.Apply This.ViewModel
    
End Sub

Private Sub BindViewControls()
    
    Dim p_layout As ILayout
    Set p_layout = This.Factory.NewLayout().Initialize(Me, 25, 20)
    With p_layout
        .BindControlLayout Me, Me.Controls("LabelIndex1"), AnchorEdges.LeftAnchor + AnchorEdges.RightAnchor
        .BindControlLayout Me, Me.Controls("LabelIndex2"), AnchorEdges.LeftAnchor + AnchorEdges.RightAnchor
        .BindControlLayout Me, Me.Controls("TextBoxIndex1"), AnchorEdges.LeftAnchor + AnchorEdges.RightAnchor
        .BindControlLayout Me, Me.Controls("TextBoxIndex2"), AnchorEdges.LeftAnchor + AnchorEdges.RightAnchor
        .BindControlLayout Me, Me.Controls("CommandButtonIndex1"), AnchorEdges.LeftAnchor + AnchorEdges.RightAnchor
    End With
    p_layout.ResizeLayout
    
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
    InitializeView
    BindViewControls
    Me.Show vbModal
End Sub

Private Function IView_ShowDialog() As Boolean
    InitializeView
    BindViewControls
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
