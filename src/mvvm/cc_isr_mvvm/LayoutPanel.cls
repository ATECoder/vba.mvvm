VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LayoutPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - -
''' <summary>   </summary>
''' - - - - - - - - - - - - - - - - - - - - -
Option Explicit

Private Type TState
    Frame As MSForms.Frame
    Direction As LayoutDirection
    NextControlTop As Long
    NextControlLeft As Long
    Margin As Long
End Type

Private This As TState

''' <summary>   Initializes and returns this <see cref="LayoutPanel"/> object. </summary>
Public Function Initialize(ByVal a_parentForm As MSForms.UserForm, _
        Optional ByVal a_parent As MSForms.Controls, _
        Optional ByVal a_height As Single, _
        Optional ByVal a_width As Single) As LayoutPanel
    
    If a_parent Is Nothing Then Set a_parent = a_parentForm.Controls
    
    Dim p_frameControl As MSForms.Control
    Set p_frameControl = a_parent.Add(BindingDefaults.FrameProgId, Visible:=False)
    If a_height <> 0 Then
        p_frameControl.Height = a_height
    Else
        p_frameControl.Height = a_parentForm.Height '<~ a_value member call is late-bound
    End If
    
    If a_width <> 0 Then
        p_frameControl.Width = a_width
    Else
        p_frameControl.Width = a_parentForm.Width '<~ a_value member call is late-bound
    End If
    
    Set Me.Frame = p_frameControl
    Me.Frame.BorderStyle = MSForms.fmBorderStyle.fmBorderStyleNone
    Me.Frame.Caption = VBA.vbNullString
    Me.Frame.BackColor = a_parentForm.BackColor
    Me.Frame.ScrollBars = MSForms.fmScrollBars.fmScrollBarsBoth
    Me.Frame.KeepScrollBarsVisible = MSForms.fmScrollBars.fmScrollBarsNone
    
    Set Initialize = Me
    
End Function

Public Property Get Frame() As MSForms.Frame
    Set Frame = This.Frame
End Property

Public Property Set Frame(ByVal a_value As MSForms.Frame)
    Set This.Frame = a_value
End Property
