# MVVM in VBA!

This project aims at implementing object-oriented programming in VBA and *Model-View-ViewModel*. 

# MVVM?

*Model-View-ViewModel* (MVVM) is a UI design pattern used in modern software development, both in Win32/desktop (WPF/XAML) and web front-ends (Javascript).
What sets this pattern apart from, say, *Model-View-Presenter*, is property and command bindings: we don't handle control events anymore, so the form's 
code-behind is focused on the only concern that remains - presentation.

> In MVVM, we're going to be referring to a `UserForm` as a *View* to broadly generalize the abstraction, but keep in mind that a *View* could just as well be a
`MSForms.Tab` control in a `MSForms.TabStrip` container, itself a child of a `UserForm`. The "Model-View-ViewModel" triad is about *abstractions*, so think of the
*View* as whichever component is responsible for directly interacting with the user.

This is a significant departure from how VBA traditionally makes you reason about programming. The Visual Basic Editor (VBE) has made a lot of us believe having lots of small, specialized modules was combersome and counter-productive. We are rightfully reluctant to code against interfaces, when there's no IDE support to navigate to their implementations. What if we just ran with it though, and embraced the full breadth of what [**Rubberduck**](https://github.com/rubberduck-vba/Rubberduck) *and VBA as a language* have to offer? This project is what happens then.

We can still drag-and-drop design our forms - but a *View* will only initialize property and command bindings, and MVVM does everything else. Or we can use an API to create the entire UI at run-time and bind the controls to *ViewModel* properties; either way, with MVVM the only code that's needed in a form's code-behind module, is code that configures all the property bindings, and boilerplate `IView` interface implementation.

The *ViewModel* is an object that exposes all the properties needed by the *View*, and implements the `INotifyPropertyChanged` interface to notify listeners (property bindings) when a value needs to be synchronized.

The *Model* is an abstraction representing the object(s) responsible for retrieving and persisting the *ViewModel* data, as applicable. It's arguably also the *commands* you implement that read *ViewModel* properties and pass them to some stored procedure on SQL Server.

## MVVM and Worksheets

This project is aimed also at demonstrating the implementation of MVVM in worksheets.

# Getting Started

1. Get [Rubberduck](https://rubberduckvba.com). Seriously, you'll need it.
2. Download `MVVM.xlsm` from this repository and open it in Microsoft Excel, then press Alt+F11 to bring up the VBE.
3. Add a new *user form* (`UserForm1`) and paste this code in:

```vba
Option Explicit
Implements IView
Implements ICancellable

Private Type TState
    Context As MVVM.IAppContext
    ViewModel As Class1
    IsCancelled As Boolean
End Type

Private This As TState

'@Description "Creates a new instance of this form."
Public Function Create(ByVal Context As MVVM.IAppContext, ByVal ViewModel As Class1) As IView
    Dim Result As UserForm1
    Set Result = New UserForm1
    Set Result.Context = Context
    Set Result.ViewModel = ViewModel
    Set Create = Result
End Function

Public Property Get Context() As MVVM.IAppContext
    Set Context = This.Context
End Property

Public Property Set Context(ByVal RHS As MVVM.IAppContext)
    Set This.Context = RHS
End Property

Public Property Get ViewModel() As Object
    Set ViewModel = This.ViewModel
End Property

Public Property Set ViewModel(ByVal RHS As Object)
    Set This.ViewModel = RHS
End Property

Private Sub OnCancel()
    This.IsCancelled = True
    Me.Hide
End Sub

Private Sub InitializeView()
    With This.Context.Bindings
        'TODO configure property bindings
    End With
    With This.Context.Commands
        'TODO configure command bindings
    End With
    This.Context.Bindings.Apply This.ViewModel
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
    Me.Show vbModal
End Sub

Private Function IView_ShowDialog() As Boolean
    InitializeView
    Me.Show vbModal
    IView_ShowDialog = Not This.IsCancelled
End Function

Private Property Get IView_ViewModel() As Object
    Set IView_ViewModel = This.ViewModel
End Property

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = True
        OnCancel
    End If
End Sub
```

4. Add a new *standard module* (`Module1`) and a new *class module* (`Class1`) to the project, then add a new parameterless `Sub` procedure (say, `DoSomething`) to `Module1`. Inside that procedure scope:
   - Declare a `Context As IAppContext` object reference, and assign it to the output of the `AppContext.Create` *factory method*.
   - Declare a `ViewModel As Class1` object reference, and then `Set ViewModel = New Class1`.
   - Declare a `View As IView` object reference, and then `Set View = UserForm1.Create(Context, ViewModel)`.

5. Add the properties you need in `Class1`; make the class implement the `INotifyPropertyChanged` to support 2-way bindings. Use `.BindPropertyPath` in the `With This.Context.Bindings` block of the `InitializeView` method to configure property bindings and associate a *ViewModel* property with a property of a control on the form.

6. Add a new class (`Class2`) and make it implement the `ICommand` interface; the `Context` parameter in both `CanExecute` and `Execute` methods holds a reference to the *ViewModel*. Use `.BindCommand` in the `With This.Context.Commands` block of the `InitializeView` method to configure command bindings and associate a command object with a `CommandButton` control on the form.

## Features

The 100+ modules solve many problems related to building and programming user interfaces in VBA, and provide an object model that gives an application a solid, decoupled backbone structure.

### Object Model

The `IAppContext` interface, and its `AppContext` implementation, are at the top of the MVVM object model. This *context* object exposes `IBindingManager`, `ICommandManager`, and `IValidationManager` objects (among others), each holding their own piece of the application's state (property bindings, command bindings, and binding validation errors, respectively).

### Property Bindings

The `INotifyPropertyChanged` interface allows property bindings to work both from the source (ViewModel) to the target (UI controls), and from the target to the source. Hence, by implementing this interface on ViewModel classes, UI code can bind a ViewModel property to a `MSForms.TextBox` control (or anything), via the `IBindingManager.BindPropertyPath` method - by letting the manager infer most of everything...

```vba
With Context.Bindings 'where Context is an IAppContext object reference
    ' use IBindingManager.BindPropertyPath to bind a ViewModel property to a property of a MSForms control target.
    .BindPropertyPath ViewModel, "Instructions", Me.InstructionsLabel
End With
```

...or by configuring every aspect of the binding explicitly.

### Validation

Application code may implement the `IValueValidator` interface to supply a property binding with a `Validator` argument. Bindings that fail validation use the default *dynamic error adorner* (that was configured when the top-level `AppContext` is created) to display configurable visual indicators (border, background, font colors, but also dynamic tooltips, icons, and labels); when the binding is valid again, the visual cues are hidden and the `IValidationManager` holds no more `IValidationError` objects in its `ValidationErrors` collection for the ViewModel's binding context (each ViewModel gets its own "validation scope").

By default, an invalid field visually looks like this:

![an invalid string property binding with the default dynamic adorner shown](https://user-images.githubusercontent.com/5751684/97099459-ac19ac80-165f-11eb-9430-7fda96dc4d8b.png)


### Command Bindings

The `ICommand` interface can be implemented for anything that needs to happen in response to the user clicking a button: in MVVM you don't handle `Click` events anymore, instead you *bind* an implementation of the `ICommand` interface to a `MSForms.CommandButton` control: the MVVM infrastructure code automatically takes care to enable or disable that control (you provide the `ICommand.CanExecute` Boolean logic, MVVM automatically invokes it).

```vba
With Context.Commands 'where Context is an IAppContext object reference
    ' use ICommandManager.BindCommand to bind a MSForms.CommandButton to any ICommand object.
    .BindCommand ViewModel, Me.CommandButton1, ViewModel.SomeCommand
End With
```

### Dynamic UI

This part of the API is still very much subject to breaking changes since it's very much alpha-stage, but the idea is to provide an API to make it easy to programmatically *generate* a user interface from VBA code, and automatically create the associated property and command bindings.

Whether your UI is dynamic or made at design-time, the recommendation would be to create the bindings in a dedicated `InitializeView` procedure in the form's code-behind.

This example snippet is from the `ExampleDynamicView` module - remember to invoke `IBindingManager.Apply` to bring it all to life:

```vba
Private Sub InitializeView()
    
    Dim Layout As IContainerLayout
    Set Layout = ContainerLayout.Create(Me.Controls, TopToBottom)
    
    With DynamicControls.Create(This.Context, Layout)
        
        With .LabelFor("All controls on this form are created at run-time.")
            .Font.Bold = True
        End With
        
        .LabelFor BindingPath.Create(This.ViewModel, "Instructions")
        
        .TextBoxFor BindingPath.Create(This.ViewModel, "StringProperty"), _
                    Validator:=New RequiredStringValidator, _
                    TitleSource:="Some String:"
                    
        .TextBoxFor BindingPath.Create(This.ViewModel, "CurrencyProperty"), _
                    FormatString:="{0:C2}", _
                    Validator:=New DecimalKeyValidator, _
                    TitleSource:="Some Amount:"
        
        .CommandButtonFor AcceptCommand.Create(Me, This.Context.Validation), This.ViewModel, "Close"
        
    End With
    
    This.Context.Bindings.Apply This.ViewModel
End Sub
```

![the ExampleDynamicView at run-time](https://user-images.githubusercontent.com/5751684/97261293-e8cadc80-17f4-11eb-9c01-5733632a05fe.png)

[About binding]: https://rubberduckvba.blog/2020/10/25/making-mvvm-work-in-vba-part-3-bindings/
