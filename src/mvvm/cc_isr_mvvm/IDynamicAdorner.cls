VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDynamicAdorner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Describes an object that can decorate a binding target with implementation-defined components."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Describes an object that can decorate a binding target with implementation-defined components. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit

''' <summary>   Brings the adorner components into view. </summary>
Public Sub Show(ByVal a_message As String)
Attribute Show.VB_Description = "Brings the adorner components into view."
End Sub

''' <summary>   Hides the adorner components from view. </summary>
Public Sub Hide()
Attribute Hide.VB_Description = "Hides the adorner components from view."
End Sub
