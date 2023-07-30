VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BindingDefaults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Contains binding property defaults"
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Binding properties and controls defaults"
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit

Private Type ThisData
    MarginDefault As Integer
    LayoutDefaultWidth As Integer
    LayoutDefaultHeight As Integer
End Type

Private This As ThisData

''' <summary>   Initializes this <see cref="BindingDefaults"/> object. </summary>
Public Sub Initialize()
    This.MarginDefault = 3
    This.LayoutDefaultWidth = 240
    This.LayoutDefaultHeight = 180
End Sub

''' <summary>   Gets the layout default height. </summary>
''' <value>   [Integer]. </value>
Public Property Get LayoutDefaultHeight() As Integer
    LayoutDefaultHeight = This.LayoutDefaultHeight
End Property

''' <summary>   Sets the layout default height. </summary>
''' <param naem ="a_value">   the value. </param>
Public Property Let LayoutDefaultHeight(ByVal a_value As Integer)
    This.LayoutDefaultHeight = a_value
End Property

''' <summary>   Gets the layout default width. </summary>
''' <value>   [Integer]. </value>
Public Property Get LayoutDefaultWidth() As Integer
    LayoutDefaultWidth = This.LayoutDefaultWidth
End Property

''' <summary>   Sets the layout default width. </summary>
''' <param naem ="a_value">   the value. </param>
Public Property Let LayoutDefaultWidth(ByVal a_value As Integer)
    This.LayoutDefaultWidth = a_value
End Property

''' <summary>   Gets the margin default size. </summary>
''' <value>   [Integer]. </value>
Public Property Get MarginDefault() As Integer
    MarginDefault = This.MarginDefault
End Property

''' <summary>   Sets the margin default size. </summary>
''' <param naem ="a_value">   the value. </param>
Public Property Let MarginDefault(ByVal a_value As Integer)
    This.MarginDefault = a_value
End Property



' Private Const Default TargetControlProperty As String = "Caption"

Public Property Get CaptionDefaultProperty() As String
    CaptionDefaultProperty = "Caption"
End Property

Public Property Get CheckBoxProgId() As String
    CheckBoxProgId = "Forms.CheckBox.1"
End Property

' Private Const Default TargetControlProperty As String = "Value"
' Private Const FormsControlProgId As String = "Forms.CheckBox.1"

Public Property Get CheckBoxDefaultProperty() As String
    CheckBoxDefaultProperty = "Value"
End Property

Public Property Get ComboBoxProgId() As String
    ComboBoxProgId = "Forms.ComboBox.1"
End Property

' Private Const Default TargetControlProperty As String = "Value"
' Private Const FormsControlProgId As String = "Forms.ComboBox.1"

Public Property Get ComboBoxDefaultProperty() As String
    ComboBoxDefaultProperty = "Value"
End Property

Public Property Get CommandButtonProgId() As String
    CommandButtonProgId = "Forms.CommandButton.1"
End Property

Public Property Get CommandButtonDefaultProperty() As String
    CommandButtonDefaultProperty = "Text"
End Property

Public Property Get FrameProgId() As String
    FrameProgId = "Forms.Frame.1"
End Property

Public Property Get FrameDefaultProperty() As String
    FrameDefaultProperty = "Caption"
End Property

Public Property Get ImageProgId() As String
    ImageProgId = "Forms.Image.1"
End Property

Public Property Get ImageDefaultProperty() As String
    ImageDefaultProperty = "Image"
End Property

Public Property Get LabelProgId() As String
    LabelProgId = "Forms.Label.1"
End Property

Public Property Get LabelDefaultProperty() As String
    LabelDefaultProperty = "Text"
End Property

Public Property Get ListBoxProgId() As String
    ListBoxProgId = "Forms.ListBox.1"
End Property

Public Property Get ListBoxDefaultProperty() As String
    ListBoxDefaultProperty = "Value"
End Property

'Private Const Default TargetControlProperty As String = "Value"
'Private Const FormsControlProgId As String = "Forms.ListBox.1"

Public Property Get MultiPageProgId() As String
    MultiPageProgId = "Forms.MultiPage.1"
End Property

' Private Const Default TargetControlProperty As String = "Value"
' Private Const FormsControlProgId As String = "Forms.MultiPage.1"

Public Property Get MultiPageDefaultProperty() As String
    MultiPageDefaultProperty = "Value"
End Property

Public Property Get OptionButtonProgId() As String
    OptionButtonProgId = "Forms.OptionButton.1"
End Property

' Private Const Default TargetControlProperty As String = "Value"
' Private Const FormsControlProgId As String = "Forms.OptionButton.1"

Public Property Get OptionButtonDefaultProperty() As String
    OptionButtonDefaultProperty = "Value"
End Property


' Private Const Default TargetControlProperty As String = "Value"

Public Property Get BindingBaseDefaultProperty() As String
    BindingBaseDefaultProperty = "value"
End Property


Public Property Get ScrollBarProgId() As String
    ScrollBarProgId = "Forms.ScrollBar.1"
End Property

' Private Const Default TargetControlProperty As String = "Value"
' Private Const FormsControlProgId As String = "Forms.ScrollBar.1"

Public Property Get ScrollBarDefaultProperty() As String
    ScrollBarDefaultProperty = "Value"
End Property

Public Property Get SpinButtonProgId() As String
    SpinButtonProgId = "Forms.SpinButton.1"
End Property

' Private Const Default TargetControlProperty As String = "Value"
' Private Const FormsControlProgId As String = "Forms.SpinButton.1"

Public Property Get SpinButtonDefaultProperty() As String
    SpinButtonDefaultProperty = "Value"
End Property

Public Property Get TabStripProgId() As String
    TabStripProgId = "Forms.TabStrip.1"
End Property

' Private Const Default TargetControlProperty As String = "Value"
' Private Const FormsControlProgId As String = "Forms.TabStrip.1"

Public Property Get TabStripDefaultProperty() As String
    TabStripDefaultProperty = "Value"
End Property

Public Property Get TextBoxProgId() As String
    TextBoxProgId = "Forms.TextBox.1"
End Property

'Private Const Default TargetControlProperty As String = "Text"
'Private Const FormsControlProgId As String = "Forms.TextBox.1"

Public Property Get TextBoxDefaultProperty() As String
    TextBoxDefaultProperty = "Text"
End Property

Public Property Get WorksheetAreaDefaultProperty() As String
    WorksheetAreaDefaultProperty = "Value"
End Property

Public Property Get WorksheetCellDefaultProperty() As String
    WorksheetCellDefaultProperty = "Value"
End Property



