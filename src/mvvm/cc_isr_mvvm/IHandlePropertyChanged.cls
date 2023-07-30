VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IHandlePropertyChanged"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An object that can be registered as an INotifyPropertyChanged handler."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   An object that can be registered as an INotifyPropertyChanged handler. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit

''' <summary>   A callback invoked when a property value is set. </summary>
Public Sub HandlePropertyChanged(ByVal a_source As Object, ByVal a_propertyName As String)
Attribute HandlePropertyChanged.VB_Description = "A callback invoked when a property value is set."
End Sub
