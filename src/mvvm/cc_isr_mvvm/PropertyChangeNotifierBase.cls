VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PropertyChangeNotifierBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Provides a base implementation for the INotifyPropertyChanged observer interface."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Provides a base implementation for the INotifyPropertyChanged observer interface. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit
Implements INotifyPropertyChanged
Implements cc_isr_Core.IDisposable

Private Type TState
    Handlers As Collection
End Type
Private This As TState

Private Sub Class_Initialize()
    Set This.Handlers = New Collection
End Sub

' - - - - - - - - - - - - - - - -
'
' Disposable implementation
'
' - - - - - - - - - - - - - - - -

Private Sub IDisposable_Dispose()
    Set This.Handlers = Nothing
End Sub

Private Sub INotifyPropertyChanged_OnPropertyChanged(ByVal a_source As Object, ByVal a_propertyName As String)
    Dim p_handler As IHandlePropertyChanged
    For Each p_handler In This.Handlers
        p_handler.HandlePropertyChanged a_source, a_propertyName
    Next
End Sub

Private Sub INotifyPropertyChanged_RegisterHandler(ByVal a_handler As IHandlePropertyChanged)
    This.Handlers.Add a_handler
End Sub
