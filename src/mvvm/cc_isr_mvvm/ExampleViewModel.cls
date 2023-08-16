VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExampleViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An example ViewModel implementation for some dialog."
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   An example ViewModel implementation for some dialog. </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Implements INotifyPropertyChanged
Option Explicit

Private Type TViewModel
    Notifier As INotifyPropertyChanged
    
    SomeCommand As ICommand
    SomeFilePath As String
    SomeAmount As Currency
    SomeDate As Date
    SomeProperty As String
    SomeOption As Boolean
    SomeOtherOption As Boolean
    SomeItems As Variant
    SelectedItemText As String
    SelectedItemIndex As Long
    BooleanProperty As Boolean
    ByteProperty As Byte
    CurrencyProperty As Currency
    DateProperty As Date
    DoubleProperty As Double
    LongProperty As Long
    StringProperty As String
End Type

Private This As TViewModel

Public Property Get BooleanProperty() As Boolean
    BooleanProperty = This.BooleanProperty
End Property

Public Property Let BooleanProperty(ByVal a_value As Boolean)
    If This.BooleanProperty <> a_value Then
        This.BooleanProperty = a_value
        OnPropertyChanged "BooleanProperty"
    End If
End Property

Public Property Get ByteProperty() As Byte
    ByteProperty = This.ByteProperty
End Property

Public Property Let ByteProperty(ByVal a_value As Byte)
    If This.ByteProperty <> a_value Then
        This.ByteProperty = a_value
        OnPropertyChanged "ByteProperty"
    End If
End Property

Public Property Get CurrencyProperty() As Currency
    CurrencyProperty = This.CurrencyProperty
End Property

Public Property Let CurrencyProperty(ByVal a_value As Currency)
    If This.CurrencyProperty <> a_value Then
        This.CurrencyProperty = a_value
        OnPropertyChanged "CurrencyProperty"
    End If
End Property

Public Property Get DateProperty() As Date
    DateProperty = This.DateProperty
End Property

Public Property Let DateProperty(ByVal a_value As Date)
    If This.DateProperty <> a_value Then
        This.DateProperty = a_value
        OnPropertyChanged "DateProperty"
    End If
End Property

Public Property Get DoubleProperty() As Double
    DoubleProperty = This.DoubleProperty
End Property

Public Property Let DoubleProperty(ByVal a_value As Double)
    If This.DoubleProperty <> a_value Then
        This.DoubleProperty = a_value
        OnPropertyChanged "DoubleProperty"
    End If
End Property

Public Property Get LongProperty() As Long
    LongProperty = This.LongProperty
End Property

Public Property Let LongProperty(ByVal a_value As Long)
    If This.LongProperty <> a_value Then
        This.LongProperty = a_value
        OnPropertyChanged "LongProperty"
    End If
End Property

Public Property Get StringProperty() As String
    StringProperty = This.StringProperty
End Property

Public Property Let StringProperty(ByVal a_value As String)
    If This.StringProperty <> a_value Then
        This.StringProperty = a_value
        OnPropertyChanged "StringProperty"
    End If
End Property

Public Property Get SomeFilePath() As String
    SomeFilePath = This.SomeFilePath
End Property

Public Property Let SomeFilePath(ByVal a_value As String)
    If This.SomeFilePath <> a_value Then
        This.SomeFilePath = a_value
        OnPropertyChanged "SomeFilePath"
    End If
End Property

Public Property Get SomeItems() As Variant
    SomeItems = This.SomeItems
End Property

Public Property Get SelectedItemText() As String
    SelectedItemText = This.SelectedItemText
End Property

Public Property Let SelectedItemText(ByVal a_value As String)
    If This.SelectedItemText <> a_value Then
        This.SelectedItemText = a_value
        OnPropertyChanged "SelectedItemText"
    End If
End Property

Public Property Get SelectedItemIndex() As Long
    SelectedItemIndex = This.SelectedItemIndex
End Property

Public Property Let SelectedItemIndex(ByVal a_value As Long)
    If This.SelectedItemIndex <> a_value Then
        This.SelectedItemIndex = a_value
        OnPropertyChanged "SelectedItemIndex"
    End If
End Property

Public Property Get SomeAmount() As Currency
    SomeAmount = This.SomeAmount
End Property

Public Property Let SomeAmount(ByVal a_value As Currency)
    If This.SomeAmount <> a_value Then
        This.SomeAmount = a_value
        OnPropertyChanged "SomeAmount"
    End If
End Property

Public Property Get SomeDate() As Date
    SomeDate = This.SomeDate
End Property

Public Property Let SomeDate(ByVal a_value As Date)
    If This.SomeDate <> a_value Then
        This.SomeDate = a_value
        OnPropertyChanged "SomeDate"
    End If
End Property

Public Property Get SomeProperty() As String
    SomeProperty = This.SomeProperty
End Property

Public Property Let SomeProperty(ByVal a_value As String)
    If This.SomeProperty <> a_value Then
        This.SomeProperty = a_value
        OnPropertyChanged "SomeProperty"
    End If
End Property

Public Property Get SomeOption() As Boolean
    SomeOption = This.SomeOption
End Property

Public Property Let SomeOption(ByVal a_value As Boolean)
    If This.SomeOption <> a_value Then
        This.SomeOption = a_value
        OnPropertyChanged "SomeOption"
    End If
End Property

Public Property Get SomeOtherOption() As Boolean
    SomeOtherOption = This.SomeOtherOption
End Property

Public Property Let SomeOtherOption(ByVal a_value As Boolean)
    If This.SomeOtherOption <> a_value Then
        This.SomeOtherOption = a_value
        OnPropertyChanged "SomeOtherOption"
    End If
End Property

Public Property Get SomeCommand() As ICommand
    Set SomeCommand = This.SomeCommand
End Property

Public Property Set SomeCommand(ByVal a_value As ICommand)
    Set This.SomeCommand = a_value
End Property

Public Property Get SomeOptionName() As String
    SomeOptionName = "Auto"
End Property

Public Property Get SomeOtherOptionName() As String
    SomeOtherOptionName = "Manual/Browse"
End Property

Public Property Get Instructions() As String
    Instructions = "Example instructions."
End Property

Private Sub OnPropertyChanged(ByVal a_propertyName As String)
    This.Notifier.OnPropertyChanged Me, a_propertyName
End Sub

Private Sub Class_Initialize()
    Set This.Notifier = New PropertyChangeNotifierBase
    ReDim This.SomeItems(0 To 4)
    Dim i As Long
    For i = LBound(This.SomeItems) To UBound(This.SomeItems)
        This.SomeItems(i) = "Selection " & i
    Next
End Sub

' - - - - - - - - - - - - - - - -
'
' Notify Property Change implementation
'
' - - - - - - - - - - - - - - - -

Private Sub INotifyPropertyChanged_OnPropertyChanged(ByVal a_source As Object, ByVal a_propertyName As String)
    This.Notifier.OnPropertyChanged a_source, a_propertyName
End Sub

Private Sub INotifyPropertyChanged_RegisterHandler(ByVal a_handler As IHandlePropertyChanged)
    This.Notifier.RegisterHandler a_handler
End Sub


