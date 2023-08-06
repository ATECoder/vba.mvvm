VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SelectOptionCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "ThingComparer"
'@PredeclaredID
Option Explicit

Implements ICommand

Public Enum CommandType
    ThingXClick
    ThingYClick
End Enum

Private Type TState
    View As IView
    ClickType As CommandType
End Type

Private this As TState

Public Function Create(ByVal ClickType As CommandType, ByVal View As IView) As SelectOptionCommand
    Dim result As SelectOptionCommand
    Set result = New SelectOptionCommand
    
    result.ClickType = ClickType
    Set result.View = View
    Set Create = result
End Function

Friend Property Let ClickType(ByVal RHS As CommandType)
    this.ClickType = RHS
End Property

Friend Property Set View(ByVal RHS As IView)
    GuardClauses.GuardDefaultInstance Me, SelectOptionCommand
    GuardClauses.GuardDoubleInitialization this.View, TypeName(Me)
    Set this.View = RHS
End Property

Private Function ICommand_CanExecute(ByVal Context As Object) As Boolean
    ICommand_CanExecute = True
End Function

Private Property Get ICommand_Description() As String
    ICommand_Description = "Click to select"
End Property

Private Sub ICommand_Execute(ByVal Context As Object)
    Dim ViewModel As ThingComparisonViewModel
    Set ViewModel = Context 'REVIEW: Or this.View.ViewModel
    'just need to save click result
    Select Case this.ClickType
        Case ThingXClick: ViewModel.Choice = xChosen
        Case ThingYClick: ViewModel.Choice = yChosen
    End Select
    this.View.Hide
End Sub
