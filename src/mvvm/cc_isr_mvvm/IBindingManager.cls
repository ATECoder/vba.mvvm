VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IBindingManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Represents an object responsible for managing property bindings in an application context."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Represents an object responsible for managing property bindings in an application context. </summary>
''' <remarks>
''' The binding manager is the object that knows about all the property bindings; each property binding needs
''' a reference to the binding manager that owns it, in order to invoke data validation and trigger command
''' bindings to evaluate whether commands can be executed. This object is automatically created when you create
''' an <c>AppContext</c> instance, but the <c>AppContext</c> can be injected with any <c>IBindingManager</c>
''' implementation as needed.
''' </remarks>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit

''' <summary>   Creates a new property binding for the specified path, source, and target. </summary>
''' <remarks>
''' <see href="https://rubberduckvba.blog/2020/10/25/making-mvvm-work-in-vba-part-3-bindings/">MVVM Binding</see>
''' </remarks>
''' <param name="a_source">    </param>
''' <param name="">    </param>
''' <param name="">    </param>
''' <param name="">    </param>
''' <param name="">    </param>
''' <param name="">    </param>
''' <param name="">    </param>
''' <param name="">    </param>
''' <param name="">    </param>
''' <param name="">    </param>
Public Function BindPropertyPath(ByVal a_source As Object, ByVal a_propertyPath As String, ByVal a_target As Object, _
    Optional ByVal a_targetProperty As String, _
    Optional ByVal a_mode As BindingMode = BindingMode.TwoWayBinding, _
    Optional ByVal a_updateTrigger As BindingUpdateSourceTrigger = BindingUpdateSourceTrigger.OnPropertyChanged, _
    Optional ByVal a_validator As IValueValidator, _
    Optional ByVal a_converter As IValueConverter, _
    Optional ByVal a_formatString As String, _
    Optional ByVal a_validationAdorner As IDynamicAdorner) As IPropertyBinding
Attribute BindPropertyPath.VB_Description = "Creates a new property binding for the specified path, source, and target."
End Function

''' <summary>   Gets a reference to an object that creates notification helper object instances. </summary>
Public Property Get NotifierFactory() As INotifierFactory
Attribute NotifierFactory.VB_Description = "Gets a reference to an object that creates notification helper object instances."
End Property

''' <summary>   Gets a reference to an object that creates string-formatting helper object instances. </summary>
Public Property Get StringFormatterFactory() As IStringFormatterFactory
Attribute StringFormatterFactory.VB_Description = "Gets a reference to an object that creates string-formatting helper object instances."
End Property

''' <summary>   Gets a value indicating whether bindings produce debugger output. </summary>
Public Property Get DebugOutput() As Boolean
Attribute DebugOutput.VB_Description = "Gets a value indicating whether bindings produce debugger output."
End Property

''' <summary>   Applies all property bindings. </summary>
Public Sub Apply(ByVal a_context As Object)
Attribute Apply.VB_Description = "Applies all property bindings."
End Sub
