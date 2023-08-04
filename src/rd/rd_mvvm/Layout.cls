VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Layout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "An object that encapsulates the UserForm Resizer"
'@ModuleAttribute VB_Ext_KEY, "Rubberduck", "Predeclared Class Module"
'@ModuleDescription "An object that encapsulates the UserForm Layout"
'@Folder MVVM.Infrastructure.View.Dynamic.Layout
'@PredeclaredId
'@Exposed
Option Explicit

Implements ILayout

#If VBA7 Then
    Private Declare PtrSafe Function GetSystemMetrics32 Lib "user32" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As LongPtr) As Long
#Else
    Private Declare Function GetSystemMetrics32 Lib "user32" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long
    Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
    Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
#End If

Private Const SM_CXSCREEN As Long = 0               'Horizontal Resolution
Private Const SM_CYSCREEN As Long = 1               'Vertical Resolution
Private Const LOGPIXELSX As Long = 88               'Pixels/inch in X
Private Const POINTS_PER_INCH As Long = 72          'A point is defined as 1/72 inches
Private Const DEFAULTWIDTH As Long = 240
Private Const DEFAULTHEIGHT As Long = 180


Private Type TState
    TotalMonitors As Long
    HorizontalResInPixel As Long
    VerticalResInPixel As Long
    
    Object As Object
    Width As Single
    Height As Single
    LayoutBindings As Collection
    
End Type

Private This As TState

Public Property Get HorizontalResInPixel() As Long
    HorizontalResInPixel = GetSystemMetrics32(SM_CXSCREEN) * PointsPerPixel
End Property

Public Property Get VerticalResInPixel() As Long
    VerticalResInPixel = GetSystemMetrics32(SM_CYSCREEN) * PointsPerPixel
End Property

Public Property Get Object() As Object
    Set Object = This.Object
End Property

Public Property Set Object(ByVal RHS As Object)
    Set This.Object = RHS
End Property

Public Property Get Width() As Single
    Width = This.Object.Width
End Property

Public Property Let Width(ByVal RHS As Single)
    GuardClauses.GuardExpression Throw:=Sgn(RHS - 0) + Sgn(RHS - 100), Message:="Value not between 0-100"
    This.Width = HorizontalResInPixel * (RHS / 100)
End Property

Public Property Get Height() As Single
    Height = This.Object.Height
End Property

Public Property Let Height(ByVal RHS As Single)
    GuardClauses.GuardExpression Throw:=Sgn(RHS - 0) + Sgn(RHS - 100), Message:="Value not between 0-100"
    This.Height = VerticalResInPixel * (RHS / 100)
End Property

Public Property Get LayoutBindings() As Collection
    Set LayoutBindings = This.LayoutBindings
End Property

Public Property Set LayoutBindings(ByVal RHS As Collection)
    GuardClauses.GuardNullReference RHS
    Set This.LayoutBindings = RHS
End Property

Public Function Create(ByVal Object As Object, Optional ByVal NewWidthPercent As Long = 50, Optional ByVal NewHeightPercent As Long = 50) As Layout
    GuardClauses.GuardNonDefaultInstance Me, Layout, TypeName(Me)
    GuardClauses.GuardNullReference Object, VBA.Information.TypeName(Me)
    
    Dim result As Layout
    Set result = New Layout
    Set result.Object = Object
        result.Width = NewWidthPercent
        result.Height = NewHeightPercent
    Set result.LayoutBindings = New Collection
    Set Create = result
    
End Function

Public Sub BindControlLayout(ByVal Parent As Object, ByVal Child As MSForms.Control, ByVal Anchor As AnchorEdges)
    GuardClauses.GuardNullReference Parent
    GuardClauses.GuardNullReference Child
    
    Dim Layout As ControlLayout
    Set Layout = New ControlLayout
    Layout.Bind Parent, Child, Anchor

    This.LayoutBindings.Add Layout
End Sub

Public Sub ResizeLayout()
    
    If This.Width < Object.Width Then Object.Width = DEFAULTWIDTH Else Object.Width = This.Width
    If This.Height < Object.Height Then Object.Height = DEFAULTHEIGHT Else Object.Height = This.Height
    
    On Error GoTo CleanFail
    Dim Layout As ControlLayout
    For Each Layout In This.LayoutBindings
        Layout.Resize Object
    Next

CleanExit:
    Exit Sub

CleanFail:
    MsgBox VBA.Err.Description, Title:=VBA.Err.Number
    Resume CleanExit
    Resume
    
End Sub

'@Description "Get Points Per Pixel Screen resloution."
Private Function PointsPerPixel() As Double

    #If VBA7 Then
        Dim HandleContex As LongPtr
        Dim DotsPerInch As LongPtr
    #Else
        Dim HandleContex As Long
        Dim DotsPerInch As Long
    #End If

    HandleContex = GetDC(0)
    DotsPerInch = GetDeviceCaps(HandleContex, LOGPIXELSX)
    PointsPerPixel = POINTS_PER_INCH / DotsPerInch
    ReleaseDC 0, HandleContex

End Function

Private Sub ILayout_BindControlLayout(ByVal Parent As Object, ByVal Child As MSForms.Control, ByVal Anchor As AnchorEdges)
    BindControlLayout Parent, Child, Anchor
End Sub

Private Sub ILayout_ResizeLayout()
    ResizeLayout
End Sub
