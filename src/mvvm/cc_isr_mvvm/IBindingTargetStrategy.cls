VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IBindingTargetStrategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A strategy for creating an IPropertyBinding "
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   A strategy for creating an IPropertyBinding. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit

''' <summary>   Infers the target property path from the runtime type of the specified Target. </summary>
Public Function DefaultPropertyBindingFor(ByVal a_context As IAppContext, ByVal a_source As IBindingPath, ByVal a_target As Object, _
    ByVal a_mode As BindingMode, _
    ByVal a_updateTrigger As BindingUpdateSourceTrigger, _
    ByVal a_validator As IValueValidator, _
    ByVal a_converter As IValueConverter, _
    ByVal a_stringFormatter As IStringFormatter, _
    ByVal a_validationAdorner As IDynamicAdorner) As IPropertyBinding
Attribute DefaultPropertyBindingFor.VB_Description = "Infers the target property path from the runtime type of the specified Target."
End Function

''' <summary>   Creates a property binding as specified. </summary>
Public Function PropertyBindingFor(ByVal a_context As IAppContext, ByVal a_source As IBindingPath, ByVal a_target As IBindingPath, _
    ByVal a_mode As BindingMode, _
    ByVal a_updateTrigger As BindingUpdateSourceTrigger, _
    ByVal a_validator As IValueValidator, _
    ByVal a_converter As IValueConverter, _
    ByVal a_stringFormatter As IStringFormatter, _
    ByVal a_validationAdorner As IDynamicAdorner) As IPropertyBinding
Attribute PropertyBindingFor.VB_Description = "Creates a property binding as specified."
End Function
