VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDynamicAdornerFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An abstract factory that creates dynamic validation error adorners."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   An abstract factory that creates dynamic validation error adorners. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit

''' <summary>   Creates a new dynamic validation error adorner for the specified target. </summary>
Public Function Create(ByVal a_target As Object) As IDynamicAdorner
Attribute Create.VB_Description = "Creates a new dynamic validation error adorner for the specified target."
End Function
