VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CaptionBindingStrategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A target binding strategy that creates a CaptionPropertyBinding by default, or a OneWayPropertyBinding given a target property path."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   A target binding strategy that creates a <see cref="CaptionPropertyBinding"/> by default,
'''             or a <see cref="OneWayPropertyBinding"/> given a target property path. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit
Implements IBindingTargetStrategy

Private Function IBindingTargetStrategy_DefaultPropertyBindingFor(ByVal a_context As IAppContext, _
        ByVal a_source As IBindingPath, ByVal a_target As Object, _
        ByVal a_mode As BindingMode, _
        ByVal a_updateTrigger As BindingUpdateSourceTrigger, _
        ByVal a_validator As IValueValidator, _
        ByVal a_converter As IValueConverter, _
        ByVal a_stringFormatter As IStringFormatter, _
        ByVal a_validationAdorner As IDynamicAdorner) As IPropertyBinding

    Set IBindingTargetStrategy_DefaultPropertyBindingFor = Factory.NewCaptionPropertyBinding().Initialize(a_context, _
            a_source, a_target, a_validator, a_converter, a_stringFormatter, a_validationAdorner)

End Function

Private Function IBindingTargetStrategy_PropertyBindingFor(ByVal a_context As IAppContext, _
        ByVal a_source As IBindingPath, ByVal a_target As IBindingPath, _
        ByVal a_mode As BindingMode, _
        ByVal a_updateTrigger As BindingUpdateSourceTrigger, _
        ByVal a_validator As IValueValidator, _
        ByVal a_converter As IValueConverter, _
        ByVal a_stringFormatter As IStringFormatter, _
        ByVal a_validationAdorner As IDynamicAdorner) As IPropertyBinding

    Set IBindingTargetStrategy_PropertyBindingFor = Factory.NewOneWayPropertyBinding().Initialize(a_context, _
            a_source, a_target, a_validator, a_converter, a_stringFormatter, a_validationAdorner)

End Function
