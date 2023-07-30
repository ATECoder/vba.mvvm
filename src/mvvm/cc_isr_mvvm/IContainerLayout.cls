VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IContainerLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Encapsulates the layout logic for dynamic UI components."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Encapsulates the layout logic for dynamic UI components. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit

''' <summary>   Adds a new dynamic control to the layout. </summary>
Public Function Add(ByVal ProgID As String, Optional ByVal a_height As Variant, Optional ByVal a_width As Variant, Optional ByVal a_applyMargin As Boolean = True) As Object
Attribute Add.VB_Description = "Adds a new dynamic control to the layout."
End Function

''' <summary>   Moves the position of the next control by a margin. </summary>
Public Sub SkipPosition(Optional ByVal a_margin As Long)
Attribute SkipPosition.VB_Description = "Moves the position of the next control by a margin."
End Sub
