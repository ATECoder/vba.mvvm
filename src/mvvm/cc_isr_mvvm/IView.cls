VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Describes a View."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Describes a View. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit

''' <summary>   Gets the ViewModel / binding Source. </summary>
Public Property Get ViewModel() As Object
Attribute ViewModel.VB_Description = "Gets the ViewModel / binding Source."
End Property

''' <summary>   Displays the View modally and returns a value indicating confirmation (True) or cancellation (False). </summary>
Public Function ShowDialog() As Boolean
Attribute ShowDialog.VB_Description = "Displays the View modally and returns a value indicating confirmation (True) or cancellation (False)."
End Function

''' <summary>   Displays the View. </summary>
Public Sub Show()
Attribute Show.VB_Description = "Displays the View."
End Sub

''' <summary>   Hides the View. </summary>
Public Sub Hide()
Attribute Hide.VB_Description = "Hides the View."
End Sub
