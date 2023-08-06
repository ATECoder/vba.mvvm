VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TextRepresentableThingsView 
   Caption         =   "TextRepresentableThingsView"
   ClientHeight    =   2340
   ClientLeft      =   90
   ClientTop       =   435
   ClientWidth     =   5880
   OleObjectBlob   =   "TextRepresentableThingsView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TextRepresentableThingsView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@Folder "ThingComparer"
Option Explicit
Implements IView
Implements ICancellable

Private Type TState
    Context As cc_isr_MVVM.IAppContext
    ViewModel As ThingComparisonViewModel
    IsCancelled As Boolean
End Type

Private this As TState

Implements IThingComparisonViewFactory

'@Description "Creates a new instance of this form."
Private Function IThingComparisonViewFactory_Create(ByVal Context As cc_isr_MVVM.IAppContext, ByVal ViewModel As ThingComparisonViewModel) As IView
Attribute IThingComparisonViewFactory_Create.VB_Description = "Creates a new instance of this form."
    Dim result As TextRepresentableThingsView
    Set result = New TextRepresentableThingsView
    Set result.Context = Context
    Set result.ViewModel = ViewModel
    Set IThingComparisonViewFactory_Create = result
End Function

Public Property Get Context() As cc_isr_MVVM.IAppContext
    Set Context = this.Context
End Property

Public Property Set Context(ByVal a_value As cc_isr_MVVM.IAppContext)
    Set this.Context = a_value
End Property

Friend Property Set ViewModel(ByVal a_value As Object)
    Set this.ViewModel = a_value
End Property

Private Sub OnCancel()
    this.IsCancelled = True
    Me.Hide
End Sub

Private Sub InitializeView()

    With this.Context.Bindings

        'these are just captions so only need one way
        .BindPropertyPath this.ViewModel, "ThingX", Me.ThingXLabel, a_mode:=OneTimeBinding
        .BindPropertyPath this.ViewModel, "ThingY", Me.ThingYLabel, a_mode:=OneTimeBinding
        
    End With
    
    With this.Context.Commands
        .BindCommand this.ViewModel, Me.ThingXLabel, this.ViewModel.SelectXCommand(Me)
        .BindCommand this.ViewModel, Me.ThingYLabel, this.ViewModel.SelectYCommand(Me)
    End With
    
    this.Context.Bindings.Apply this.ViewModel
    
End Sub

Private Property Get ICancellable_IsCancelled() As Boolean
    ICancellable_IsCancelled = this.IsCancelled
End Property

Private Sub ICancellable_OnCancel()
    OnCancel
End Sub

Private Sub IView_Hide()
    Me.Hide
End Sub

Private Sub IView_Show()
    InitializeView
    Me.Show vbModal
End Sub

Private Function IView_ShowDialog() As Boolean
    InitializeView
    Me.Show vbModal
    IView_ShowDialog = Not this.IsCancelled
End Function

Private Property Get IView_ViewModel() As Object
    Set IView_ViewModel = this.ViewModel
End Property

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = True
        OnCancel
    End If
End Sub
