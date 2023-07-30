VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICancellable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Describes an object that encapsulates 'IsCancelled' state."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Describes an object that encapsulates 'IsCancelled' state. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit

''' <summary>   Gets a value indicating whether the object is in a cancelled state. </summary>
Public Property Get IsCancelled() As Boolean
Attribute IsCancelled.VB_Description = "Gets a value indicating whether the object is in a cancelled state."
End Property

''' <summary>   Sets the IsCancelled flag to True. </summary>
Public Sub OnCancel()
Attribute OnCancel.VB_Description = "Sets the IsCancelled flag to True."
End Sub
