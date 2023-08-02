VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListBoxBindingStrategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   A target binding strategy that creates a <see cref="ListBoxPropertyBinding"/>
'''             with an inferred or specific target property path. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit
Implements IBindingTargetStrategy

Private Function IBindingTargetStrategy_DefaultPropertyBindingFor(ByVal a_context As IAppContext, ByVal a_source As IBindingPath, ByVal a_target As Object, _
        ByVal a_mode As BindingMode, _
        ByVal a_updateTrigger As BindingUpdateSourceTrigger, _
        ByVal a_validator As IValueValidator, _
        ByVal a_converter As IValueConverter, _
        ByVal a_stringFormatter As IStringFormatter, _
        ByVal a_validationAdorner As IDynamicAdorner) As IPropertyBinding

    ' target property is inferred

    Set IBindingTargetStrategy_DefaultPropertyBindingFor = Factory.NewListBoxPropertyBinding().Initialize(a_context, _
            a_source, a_target, _
            a_validator:=a_validator, _
            a_stringFormatter:=a_stringFormatter, _
            a_converter:=a_converter, _
            a_validationAdorner:=a_validationAdorner)
            
End Function

Private Function IBindingTargetStrategy_PropertyBindingFor(ByVal a_context As IAppContext, ByVal a_source As IBindingPath, _
        ByVal a_target As IBindingPath, _
        ByVal a_mode As BindingMode, _
        ByVal a_updateTrigger As BindingUpdateSourceTrigger, _
        ByVal a_validator As IValueValidator, _
        ByVal a_converter As IValueConverter, _
        ByVal a_stringFormatter As IStringFormatter, _
        ByVal a_validationAdorner As IDynamicAdorner) As IPropertyBinding

    Set IBindingTargetStrategy_PropertyBindingFor = Factory.NewListBoxPropertyBinding().Initialize(a_context, _
            a_source, a_target.Context, a_target.Path, a_mode, a_updateTrigger, a_validator, _
            a_converter, a_stringFormatter, a_validationAdorner)

End Function
