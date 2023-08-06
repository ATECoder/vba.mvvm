VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThingComparisonViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "ThingComparer"
Option Explicit

'Implements INotifyPropertyChanged

Private Type TState
    ThingX As String
    ThingY As String
    Handler As IHandlePropertyChanged
    Choice As ComparisonResult
End Type

Public Enum ComparisonResult
    xChosen
    yChosen
    CancelChosen
End Enum

Private this As TState

Public Property Get Choice() As ComparisonResult
    Choice = this.Choice
End Property

Public Property Let Choice(ByVal a_value As ComparisonResult)
    this.Choice = a_value
End Property

'@Ignore ProcedureNotUsed: Called by name by cc_isr_MVVM framework
Public Property Get ThingX() As String
    ThingX = this.ThingX
End Property

Public Property Let ThingX(ByVal a_value As String)
    this.ThingX = a_value
End Property

'@Ignore ProcedureNotUsed
Public Property Get ThingY() As String
    ThingY = this.ThingY
End Property

Public Property Let ThingY(ByVal a_value As String)
    this.ThingY = a_value
End Property

Public Property Get SelectXCommand(ByVal View As IView) As ICommand
    Set SelectXCommand = SelectOptionCommand.Create(ThingXClick, View)
End Property

Public Property Get SelectYCommand(ByVal View As IView) As ICommand
    Set SelectYCommand = SelectOptionCommand.Create(ThingYClick, View)
End Property


'Private Sub INotifyPropertyChanged_OnPropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
'    this.Handler.HandlePropertyChanged Source, PropertyName
'End Sub
'
'Private Sub INotifyPropertyChanged_RegisterHandler(ByVal Handler As IHandlePropertyChanged)
'    Set this.Handler = Handler
'End Sub
