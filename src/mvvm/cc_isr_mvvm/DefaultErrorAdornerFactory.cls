VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DefaultErrorAdornerFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "A factory that creates default validation error adorners."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   A factory that creates default validation error adorners. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit
Implements IDynamicAdornerFactory

Private Function IDynamicAdornerFactory_Create(ByVal a_target As Object) As IDynamicAdorner

    Dim p_result As IDynamicAdorner
    Select Case True
        
        Case TypeOf a_target Is MSForms.TextBox
            Set p_result = Factory.CreateValidationErrorAdorner(a_target, _
                            a_labelPosition:=Factory.CreateDynamicControlPosition(RelativePosition.AboveLeft, 0))
        
        Case TypeOf a_target Is MSForms.ComboBox
            Set p_result = Factory.CreateValidationErrorAdorner(a_target, _
                            a_iconPosition:=Factory.CreateDynamicControlPosition(RelativePosition.InsideRight, _
                                    0, a_topMargin:=3, a_rightMargin:=12))
        
        Case Else
            Debug.Print "A dynamic adorner was requested, but the target type isn't supported."
        
    End Select
    Set IDynamicAdornerFactory_Create = p_result
    
End Function
